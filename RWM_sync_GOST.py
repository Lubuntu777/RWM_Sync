#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Радиосинхронизация времени по сигналам RWM (ВНИИФТРИ)
Параметры сигнала: Rep. ITU-R TF.2487-0, п. 8.2 (RWM service)

ИСТОЧНИКИ:
  ① Аудио (PyAudio)   — звук с трансивера в USB, тон ~2000 Гц
  ② RTL-SDR Direct Q  — приём в USB: настройка (RWM − 2000 Гц) → тон 2000 Гц
  ③ Авто              — оба работают, выбирается лучший по SNR

ДЕКОДЕР: причинный КИХ-тракт, начало переднего фронта (экстраполяция),
         сетка моментов UTC из циклограммы, учёт пропусков меток,
         компенсация всех задержек включая ионосферную.
МОНИТОР: звук активного источника на колонки ПК (48 кГц, sinc ×4).
NTP-СЕРВЕР: раздача дисциплинированного времени.
"""

import os
import sys
import json
import time
import queue
import threading
import socket
import struct
import subprocess
import importlib
import math
import ctypes
import ipaddress
from collections import deque, defaultdict
from datetime import datetime, timedelta, timezone


# ------------------------------------------------------------
# Windows: поиск DLL для librtlsdr
# ------------------------------------------------------------
def _collect_dll_dirs():
    dirs = []
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
        exe = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
        exe = base
    for p in (os.path.join(base, "dll"), os.path.join(exe, "dll"), base, exe):
        if os.path.isdir(p) and p not in dirs:
            dirs.append(p)
    return dirs


def _find_dll(name, dirs):
    for d in dirs:
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
    return None


def _setup_dll_path():
    if sys.platform != "win32":
        return
    dirs = _collect_dll_dirs()
    for p in dirs:
        try:
            os.add_dll_directory(p)
        except (AttributeError, OSError):
            pass
        if p not in os.environ.get("PATH", ""):
            os.environ["PATH"] = p + os.pathsep + os.environ.get("PATH", "")
    try:
        import ctypes.util as cu
        orig = cu.find_library

        def patched(name):
            if name in ("rtlsdr", "librtlsdr"):
                for c in ("librtlsdr.dll", "rtlsdr.dll"):
                    p = _find_dll(c, dirs)
                    if p:
                        return p
            return orig(name)

        cu.find_library = patched
    except Exception:
        pass
    for n in ("libwinpthread-1.dll", "libusb-1.0.dll", "rtlsdr.dll", "librtlsdr.dll"):
        p = _find_dll(n, dirs)
        if p:
            try:
                ctypes.CDLL(p)
            except OSError:
                pass


_setup_dll_path()


def ensure_packages():
    req = {"numpy": "numpy", "pyaudio": "pyaudio", "serial": "pyserial",
           "matplotlib": "matplotlib", "rtlsdr": "pyrtlsdr[lib]"}
    if sys.platform.startswith("win"):
        req["win32com"] = "pywin32"
    for mod, pkg in req.items():
        try:
            importlib.import_module(mod)
        except Exception:
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", pkg],
                                      stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass


ensure_packages()

import tkinter as tk
from tkinter import ttk, messagebox

try:
    import numpy as np
except Exception:
    np = None
try:
    import pyaudio
except Exception:
    pyaudio = None
AUDIO_AVAILABLE = (np is not None) and (pyaudio is not None)
try:
    from rtlsdr import RtlSdr
    SDR_AVAILABLE = (np is not None)
except Exception:
    RtlSdr = None
    SDR_AVAILABLE = False
try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MPL = True
except Exception:
    MPL = False
    Figure = FigureCanvasTkAgg = None
try:
    import serial
    import serial.tools.list_ports
    SERIAL_AVAILABLE = True
except Exception:
    SERIAL_AVAILABLE = False


C = {"bg": "#0b0f16", "surface": "#131a24", "surface2": "#1a2331",
     "raise": "#212d3d", "border": "#2b3a4d", "text": "#e4ecf5",
     "muted": "#7d8fa3", "dim": "#556579", "blue": "#4c9fff",
     "cyan": "#2fd4c4", "green": "#3ddc84", "amber": "#f0b429",
     "red": "#ff5f56", "violet": "#b07cff", "magenta": "#ff5fa2"}

SEG_COLORS = {"CARRIER": "#31506e", "OFF": "#5a2a2a", "ID": "#7a4a9e",
              "A1X": "#1f7a4d", "A1N": "#1a6f78"}
SEG_LABELS = {"CARRIER": "несущая", "OFF": "пауза", "ID": "позывной",
              "A1X": "A1X 1 Гц", "A1N": "A1N 10 Гц"}

AGC_TARGET_PEAK = 0.80
AGC_MAX_GAIN = 200.0
AGC_MIN_GAIN = 0.02
AGC_PEAK_DECAY = 0.995

# ============================================================
# RWM — Rep. ITU-R TF.2487-0, п. 8.2
# ------------------------------------------------------------
# Координаты: 56°44'N 37°38'E.  Мощность 10 кВт.
# Несущие: 4996.0 / 9996.0 / 14996.0 кГц.
# Относительная погрешность несущей: 5e-12 (усреднение за сутки).
#
# Длительности (п. 8.2.1):
#   A1X (1 Гц):  секундная 100 мс,  минутная 500 мс
#   A1N (10 Гц): 0.1 с — 20 мс, начало секунды — 40 мс,
#                начало минуты — 500 мс
# Характерная точка A1X и A1N — НАЧАЛО ПЕРЕДНЕГО ФРОНТА.
#
# Часовая программа (табл. 23), внутри каждых 30 минут:
#   00:00–07:55 несущая      08:00–09:00 нет сигнала
#   09:00–10:00 позывной     10:00–19:55 A1X
#   20:00–29:55 A1N
# Сноска: метки секунд 56,57,58,59 в минутах
#   9,14,19,24,29,39,44,49,54,59 ПРОПУСКАЮТСЯ  (minute % 5 == 4)
#
# DUT1: маркировка доп. импульсами секунд 1..8 → +0.1·n,
#       9..16 → −0.1·m;  dUT1: 21..24 → +0.02·p, 31..34 → −0.02·q
# ============================================================
RWM_LAT = 56.0 + 44.0 / 60.0
RWM_LON = 37.0 + 38.0 / 60.0
RWM_FREQS_HZ = (4996000, 9996000, 14996000)
RWM_CARRIER_UNCERTAINTY = 5e-12

A1X_SECOND_MS = 100.0
A1X_MINUTE_MS = 500.0
A1N_TENTH_MS = 20.0
A1N_SECOND_MS = 40.0
A1N_MINUTE_MS = 500.0

LIGHT_KM_S = 299792.458
EARTH_R_KM = 6371.0088


def maidenhead_to_latlon(loc):
    """Maidenhead → (lat, lon) центра квадрата. Поддержка 4/6/8 знаков."""
    s = str(loc or "").strip().upper().replace(" ", "")
    if len(s) < 4 or len(s) % 2 != 0 or len(s) > 8:
        return None
    try:
        if not (s[0].isalpha() and s[1].isalpha() and s[2].isdigit() and s[3].isdigit()):
            return None
        lon = (ord(s[0]) - 65) * 20.0 - 180.0
        lat = (ord(s[1]) - 65) * 10.0 - 90.0
        if not (0 <= ord(s[0]) - 65 < 18 and 0 <= ord(s[1]) - 65 < 18):
            return None
        lon += int(s[2]) * 2.0
        lat += int(s[3]) * 1.0
        lon_sz, lat_sz = 2.0, 1.0
        if len(s) >= 6:
            if not (s[4].isalpha() and s[5].isalpha()):
                return None
            lon += (ord(s[4]) - 65) * (5.0 / 60.0)
            lat += (ord(s[5]) - 65) * (2.5 / 60.0)
            lon_sz, lat_sz = 5.0 / 60.0, 2.5 / 60.0
        if len(s) >= 8:
            if not (s[6].isdigit() and s[7].isdigit()):
                return None
            lon += int(s[6]) * (0.5 / 60.0)
            lat += int(s[7]) * (0.25 / 60.0)
            lon_sz, lat_sz = 0.5 / 60.0, 0.25 / 60.0
        return lat + lat_sz / 2.0, lon + lon_sz / 2.0
    except Exception:
        return None


def haversine_km(lat1, lon1, lat2, lon2):
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = p2 - p1
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2.0 * EARTH_R_KM * math.asin(min(1.0, math.sqrt(a)))


def path_delays_ms(dist_km, iono_height_km=300.0, hops=1):
    """Наземная волна и ионосферная (n-hop, плоская модель)."""
    ground = dist_km / LIGHT_KM_S * 1000.0
    seg = dist_km / max(1, hops)
    iono = hops * 2.0 * math.sqrt((seg / 2.0) ** 2 + iono_height_km ** 2)
    return ground, iono / LIGHT_KM_S * 1000.0


def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
CONFIG_FILE = os.path.join(APP_DIR, "config.json")

DEFAULT_CONFIG = {
    "audio_input": "",
    "cat_mode": "Direct CAT",
    "cat_com_port": "",
    "cat_virtual_port": "COM9",
    "cat_baudrate": 19200,
    "cat_freq_offset": -2000,
    "agc_enabled": True,
    "rx_freq": 9996000,
    "use_ntp_fallback": True,
    "ntp_servers": ["0.ru.pool.ntp.org", "1.ru.pool.ntp.org"],

    # --- геолокация и задержка пути ---
    "locator": "LN23HA",
    "iono_height_km": 300.0,
    "iono_hops": 1,
    "path_mode": "Iono",          # Ground | Iono | Manual
    "path_delay_ms": 0.0,         # 0 = авто из локатора
    "timing_trim_ms": 0.0,

    # --- тракт и детекция ---
    "pulse_center_hz": 2000,
    "pulse_band_hz": 600,
    "pulse_smooth_ms": 1.0,
    "edge_lo_frac": 0.2,
    "edge_hi_frac": 0.8,
    "gate_coarse_ms": 450.0,
    "gate_fine_ms": 80.0,
    "pulse_min_ms": 8.0,
    "pulse_max_ms": 700.0,
    "dur_tolerance": 0.45,
    "sync_threshold_ms": 3.0,
    "sync_min_pulses": 6,
    "decode_dut1": True,

    # --- RTL-SDR ---
    "source_mode": "SDR",
    "sdr_index": 0,
    "sdr_sample_rate": 240000,
    "sdr_if_bw_hz": 1200,
    "sdr_lock_span_hz": 1500,
    "sdr_lock_span_locked_hz": 300,
    "sdr_carrier_fft": 32768,
    "sdr_ppm_cal": 0.0,
    "sdr_autorestart": True,
    "audio_latency_ms": 0,
    "sdr_latency_ms": 0,

    # --- монитор ---
    "monitor_enabled": False,
    "monitor_volume": 0.75,
    "monitor_output": "",
    "monitor_out_rate": 48000,

    # --- NTP-сервер ---
    "ntp_server_enabled": False,
    "ntp_server_port": 123,
    "ntp_server_bind": "0.0.0.0",
    "ntp_server_require_sync": False,
    "ntp_server_allow": "",
}


def _clean(v):
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, list):
        return [_clean(x) for x in v]
    if isinstance(v, dict):
        return {str(k).strip(): _clean(x) for k, x in v.items()}
    return v


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                d = json.load(f)
            if isinstance(d, dict):
                d = {str(k).strip(): _clean(v) for k, v in d.items()}
                m = DEFAULT_CONFIG.copy()
                m.update(d)
                return m
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


CONFIG = load_config()


def resize_1d(a, n):
    if np is None or a is None or len(a) == 0:
        return np.zeros(n) if np is not None else None
    if len(a) == n:
        return a.astype(np.float64)
    if len(a) == 1:
        return np.full(n, float(a[0]))
    return np.interp(np.linspace(0, 1, n), np.linspace(0, 1, len(a)), a)


def local_ipv4_list():
    out = ["0.0.0.0", "127.0.0.1"]
    try:
        for i in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            if i[4][0] not in out:
                out.append(i[4][0])
    except Exception:
        pass
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        if ip not in out:
            out.append(ip)
    except Exception:
        pass
    return out


# ------------------------------------------------------------
# Циклограмма RWM (ITU-R TF.2487-0, табл. 23)
# ------------------------------------------------------------
class RWMSchedule:
    CARRIER = "CARRIER"
    OFF = "OFF"
    ID = "ID"
    A1X = "A1X"
    A1N = "A1N"

    # (нач_мин, кон_мин, кон_сек, режим) внутри 30-минутной половины
    TABLE = [
        (0, 7, 55, CARRIER),
        (8, 9, 0, OFF),
        (9, 10, 0, ID),
        (10, 19, 55, A1X),
        (20, 29, 55, A1N),
    ]

    @classmethod
    def get_mode(cls, dt=None):
        dt = dt or datetime.now(timezone.utc)
        m = dt.minute % 30
        s = dt.second
        for a, b, bs, mode in cls.TABLE:
            if a <= m < b:
                return mode
            if m == b and s < bs:
                return mode
        # хвост 7:55–8:00, 19:55–20:00, 29:55–30:00 — переход
        if m in (7, 19, 29) and s >= 55:
            return cls.CARRIER if m == 7 else cls.OFF
        return cls.CARRIER

    @classmethod
    def desc(cls, mode):
        return {cls.CARRIER: "Немодулированная несущая (N0N)",
                cls.OFF: "Передатчик выключен",
                cls.ID: "Опознавательные сигналы (позывной)",
                cls.A1X: "A1X 1 Гц: 100 мс / минутная 500 мс",
                cls.A1N: "A1N 10 Гц: 20 / 40 / 500 мс"}.get(mode, mode)

    @classmethod
    def is_sync_possible(cls, mode):
        return mode in (cls.A1X, cls.A1N)

    @classmethod
    def grid_step(cls, mode):
        if mode == cls.A1X:
            return 1.0
        if mode == cls.A1N:
            return 0.1
        return None

    @staticmethod
    def seconds_skipped(minute):
        """Сноска к табл. 23: минуты 9,14,19,24,29,39,44,49,54,59."""
        return (minute % 5) == 4

    @classmethod
    def pulse_expected(cls, t_grid):
        """Ожидается ли метка в этот момент сетки (по UTC)."""
        dt = datetime.fromtimestamp(t_grid, tz=timezone.utc)
        if cls.seconds_skipped(dt.minute) and dt.second >= 56:
            return False
        return True

    @classmethod
    def expected_duration_ms(cls, mode, t_grid):
        """Класс метки определяется положением в сетке UTC."""
        frac_min = t_grid - math.floor(t_grid / 60.0) * 60.0
        on_minute = min(frac_min, 60.0 - frac_min) < 1e-6
        frac_sec = t_grid - math.floor(t_grid)
        on_second = min(frac_sec, 1.0 - frac_sec) < 1e-6
        if mode == cls.A1X:
            return (A1X_MINUTE_MS, "minute") if on_minute else (A1X_SECOND_MS, "second")
        if on_minute:
            return A1N_MINUTE_MS, "minute"
        if on_second:
            return A1N_SECOND_MS, "second"
        return A1N_TENTH_MS, "tenth"

    @classmethod
    def seconds_to_next_sync(cls, dt=None):
        dt = dt or datetime.now(timezone.utc)
        pos = (dt.minute % 30) * 60 + dt.second + dt.microsecond / 1e6
        for a, b, bs, mode in cls.TABLE:
            if cls.is_sync_possible(mode) and a * 60 > pos:
                return int(a * 60 - pos)
        return int(30 * 60 - pos + 10 * 60)


# ------------------------------------------------------------
# DSP
# ------------------------------------------------------------
def design_fir_lp(cutoff, fs, n=None, transition=None):
    cutoff = float(min(max(cutoff, 20.0), fs * 0.45))
    if n is None:
        tr = float(transition) if transition else max(fs * 0.02, 200.0)
        n = int(3.5 * fs / max(tr, 50.0))
        n = max(31, min(511, n))
    n = int(n) | 1
    k = np.arange(n) - (n - 1) / 2.0
    h = np.sinc(2.0 * cutoff * k / fs) * np.hamming(n)
    s = float(np.sum(h))
    return (h / s if abs(s) > 1e-12 else np.ones(n) / n).astype(np.float64)


def design_fir_bp(center, band, fs, n):
    """Причинный полосовой КИХ, линейная фаза."""
    n = int(n) | 1
    lp = design_fir_lp(band / 2.0, fs, n=n)
    k = np.arange(n) - (n - 1) / 2.0
    return (2.0 * lp * np.cos(2.0 * math.pi * center * k / fs)).astype(np.float64)


def design_hilbert(n=201):
    n = max(int(n) | 1, 7)
    c = (n - 1) // 2
    h = np.zeros(n)
    for i in range(n):
        k = i - c
        if k != 0 and k % 2 != 0:
            h[i] = 2.0 / (math.pi * k)
    return h * np.blackman(n)


class FIRFilter:
    """Причинный КИХ, overlap-save. Задержка (N−1)/2 сэмплов."""

    def __init__(self, taps, is_complex=False):
        self.h = np.asarray(taps, dtype=np.float64)
        self.n = len(self.h)
        self.is_complex = bool(is_complex)
        self.dt = np.complex128 if is_complex else np.float64
        self.state = np.zeros(max(0, self.n - 1), dtype=self.dt)
        self._nfft = 0
        self._H = None
        self.delay_samples = (self.n - 1) // 2

    def reset(self):
        self.state = np.zeros(max(0, self.n - 1), dtype=self.dt)

    def process(self, x):
        x = np.asarray(x, dtype=self.dt)
        L = len(x)
        if L == 0:
            return np.zeros(0, dtype=self.dt)
        buf = np.concatenate([self.state, x])
        if len(buf) < self.n:
            self.state = np.concatenate(
                [np.zeros(self.n - 1 - len(buf), dtype=self.dt), buf])
            return np.zeros(0, dtype=self.dt)
        nlin = len(buf) + self.n - 1
        nfft = 1 << (nlin - 1).bit_length()
        if nfft != self._nfft:
            self._nfft = nfft
            self._H = np.fft.fft(self.h, nfft)
        y = np.fft.ifft(np.fft.fft(buf, nfft) * self._H, nfft)
        y = y[self.n - 1: self.n - 1 + L]
        self.state = buf[-(self.n - 1):].copy() if self.n > 1 else np.zeros(0, dtype=self.dt)
        return y if self.is_complex else y.real


class SincResampler:
    HALF = 12
    TAPS = 2 * HALF + 1
    HIST = 2 * HALF
    PHASES = 512
    _k = None

    @classmethod
    def kernels(cls):
        if cls._k is None:
            k = np.zeros((cls.PHASES, cls.TAPS), dtype=np.float32)
            for p in range(cls.PHASES):
                x = np.arange(-cls.HALF, cls.HALF + 1, dtype=np.float64) - p / cls.PHASES
                h = np.sinc(x) * np.kaiser(cls.TAPS, 8.0)
                s = np.sum(h)
                h = h / s if abs(s) > 1e-12 else np.eye(1, cls.TAPS, cls.HALF)[0]
                k[p] = h.astype(np.float32)
            cls._k = k
        return cls._k

    def __init__(self, ratio=1.0):
        self.ratio = float(ratio)
        self._kernels = self.kernels()
        self.reset()

    def reset(self):
        self.in_next = 0
        self.out_next = 0.0
        self.hist = np.zeros(self.HIST, dtype=np.float32)

    def process(self, y):
        y = np.asarray(y, dtype=np.float32)
        L = len(y)
        if L == 0:
            return np.zeros(0, dtype=np.float32)
        if abs(self.ratio - 1.0) < 1e-9:
            self.hist = (y[-self.HIST:].copy() if L >= self.HIST
                         else np.concatenate([self.hist, y])[-self.HIST:].copy())
            self.in_next += L
            self.out_next += L
            return y.copy()
        H = self.HIST
        start_abs = self.in_next
        full = np.concatenate([self.hist, y])
        full_start = start_abs - H
        n_full = len(full)
        self.in_next = start_abs + L
        self.hist = (y[-H:].copy() if L >= H else full[-H:].copy())
        lo = full_start + self.HALF
        hi = full_start + n_full - 1 - self.HALF
        if hi < lo:
            return np.zeros(0, dtype=np.float32)
        p0 = max(self.out_next, float(lo))
        pmax = float(hi) + 1.0 - 1e-9
        if p0 >= pmax:
            return np.zeros(0, dtype=np.float32)
        M = int(np.ceil((pmax - p0) / self.ratio))
        if M <= 0:
            return np.zeros(0, dtype=np.float32)
        pos = p0 + self.ratio * np.arange(M, dtype=np.float64)
        pos = pos[pos < pmax]
        if pos.size == 0:
            return np.zeros(0, dtype=np.float32)
        n = np.floor(pos).astype(np.int64)
        fr = pos - n
        neg = fr < 0
        if np.any(neg):
            n[neg] -= 1
            fr[neg] += 1.0
        ph = np.clip((fr * self.PHASES).astype(np.int64), 0, self.PHASES - 1)
        idx = (n - full_start)[:, None] + np.arange(-self.HALF, self.HALF + 1)[None, :]
        np.clip(idx, 0, n_full - 1, out=idx)
        out = np.sum(full[idx] * self._kernels[ph], axis=1)
        self.out_next = pos[-1] + self.ratio
        return out.astype(np.float32)


# ------------------------------------------------------------
# Источник: буфер + метки времени сэмплов
# ------------------------------------------------------------
class BufferedSource:
    SAMPLE_RATE = 12000
    BUFFER_SECONDS = 6
    name = "source"
    title = "Источник"

    def _init_buffer(self):
        self.audio_buffer = np.zeros(0, dtype=np.int16)
        self.buffer_lock = threading.Lock()
        self.running = False
        self.agc_gain = 1.0
        self.latency_sec = 0.0
        self.last_data_ts = 0.0
        self.total_samples = 0
        self._blocks = deque(maxlen=1500)
        self._peak_hold = 0.0

    def _append(self, arr):
        if arr is None or len(arr) == 0:
            return
        t = time.time()
        with self.buffer_lock:
            s = self.total_samples
            self.audio_buffer = np.concatenate((self.audio_buffer, arr))
            self.total_samples += len(arr)
            mx = self.SAMPLE_RATE * self.BUFFER_SECONDS
            if len(self.audio_buffer) > mx:
                self.audio_buffer = self.audio_buffer[-mx:]
            self._blocks.append((s, self.total_samples, t))
        self.last_data_ts = t

    def sample_to_time(self, idx):
        with self.buffer_lock:
            for s, e, t in reversed(self._blocks):
                if s <= idx < e:
                    return t - (e - idx) / float(self.SAMPLE_RATE)
        return None

    def read_from(self, since, max_samples=None):
        with self.buffer_lock:
            bstart = self.total_samples - len(self.audio_buffer)
            if since < bstart:
                since = bstart
            avail = self.total_samples - since
            if avail <= 0:
                return None, since
            if max_samples:
                avail = min(avail, max_samples)
            off = since - bstart
            return self.audio_buffer[off:off + avail].copy(), since + avail

    def get_new_samples(self, since, max_chunk=None):
        return self.read_from(since, max_chunk or self.SAMPLE_RATE // 4)

    def data_alive(self, age=3.0):
        return self.running and (time.time() - self.last_data_ts) < age

    def get_buffer(self, dur):
        need = int(dur * self.SAMPLE_RATE)
        with self.buffer_lock:
            if len(self.audio_buffer) < need:
                return None
            return self.audio_buffer[-need:].copy()

    def get_rms(self):
        with self.buffer_lock:
            if len(self.audio_buffer) == 0:
                return 0.0
            return min(1.0, float(np.sqrt(np.mean(
                np.square(self.audio_buffer.astype(np.float64))))) / 32768.0)

    def get_raw_rms(self):
        with self.buffer_lock:
            if len(self.audio_buffer) == 0:
                return 0.0
            return float(np.sqrt(np.mean(np.square(self.audio_buffer.astype(np.float64)))))

    def quality(self, tone=2000.0):
        buf = self.get_buffer(0.5)
        if buf is None or len(buf) < 1024:
            return 0.0
        try:
            x = buf.astype(np.float64) / 32768.0
            x -= float(np.mean(x))
            if float(np.sqrt(np.mean(x * x))) < 1e-4:
                return 0.0
            sp = np.abs(np.fft.rfft(x * np.hanning(len(x)))) ** 2
            f = np.fft.rfftfreq(len(x), 1.0 / self.SAMPLE_RATE)
            b = (f > tone - 200) & (f < tone + 200)
            r = (f > 100) & (f < 4000) & (~b)
            return float(np.sum(sp[b])) / (float(np.sum(sp[r])) + 1e-12)
        except Exception:
            return 0.0

    def _apply_peak_agc(self, x):
        if len(x) == 0:
            return x
        pk = float(np.max(np.abs(x)))
        if pk > self._peak_hold:
            self._peak_hold = pk
        else:
            self._peak_hold *= AGC_PEAK_DECAY
        if self._peak_hold < 1e-6:
            return x
        g = min(max(AGC_TARGET_PEAK / self._peak_hold, AGC_MIN_GAIN), AGC_MAX_GAIN)
        self.agc_gain = g
        return x * g


class AudioInput(BufferedSource):
    CHUNK = 512
    name = "audio"
    title = "Аудио (PyAudio)"

    def __init__(self, keywords="", agc=True):
        self._init_buffer()
        self.keywords = keywords
        self.agc_enabled = agc
        self.pya = pyaudio.PyAudio() if AUDIO_AVAILABLE else None
        self.stream = None
        self.device_index = None
        self.device_name = ""
        self.latency_sec = float(CONFIG.get("audio_latency_ms", 0)) / 1000.0
        self._find()

    def _find(self):
        if self.pya is None:
            return
        keys = [k.strip().lower() for k in self.keywords.split(",") if k.strip()]
        for i in range(self.pya.get_device_count()):
            try:
                inf = self.pya.get_device_info_by_index(i)
                nm = inf.get("name", "")
                if inf.get("maxInputChannels", 0) > 0:
                    if not keys or all(k in nm.lower() for k in keys):
                        self.device_index, self.device_name = i, nm
                        return
            except Exception:
                continue
        try:
            inf = self.pya.get_default_input_device_info()
            self.device_index = inf["index"]
            self.device_name = inf.get("name", "default")
        except Exception:
            self.device_index = None

    def start(self):
        if not AUDIO_AVAILABLE or self.pya is None or self.device_index is None:
            return False
        try:
            self.running = True
            self.stream = self.pya.open(
                format=pyaudio.paInt16, channels=1, rate=self.SAMPLE_RATE,
                input=True, input_device_index=self.device_index,
                frames_per_buffer=self.CHUNK, stream_callback=self._cb)
            self.stream.start_stream()
            return True
        except Exception as e:
            self.running = False
            print(f"[AudioInput] {e}")
            return False

    def _cb(self, data, frames, tinfo, status):
        if not self.running:
            return (None, pyaudio.paAbort)
        a = np.frombuffer(data, dtype=np.int16).astype(np.float64) / 32768.0
        if self.agc_enabled:
            a = self._apply_peak_agc(a)
        else:
            self.agc_gain = 1.0
        self._append(np.clip(a * 32767.0, -32768, 32767).astype(np.int16))
        return (None, pyaudio.paContinue)

    def stop(self):
        self.running = False
        if self.stream:
            try:
                self.stream.stop_stream()
                self.stream.close()
            except Exception:
                pass
            self.stream = None

    def close(self):
        self.stop()
        if self.pya:
            try:
                self.pya.terminate()
            except Exception:
                pass
            self.pya = None


class SDRInput(BufferedSource):
    name = "sdr"
    title = "RTL-SDR · USB"
    RATES = [240000, 288000, 300000]
    DEFAULT_RATE = 240000
    READ_BLOCK_SEC = 0.010
    READ_MAX = 16384
    READ_MIN = 1024

    def __init__(self, cfg=None, log_func=None):
        self._init_buffer()
        cfg = cfg or {}
        self.log = log_func or (lambda m: print(f"[SDR] {m}"))
        self.device_index = int(cfg.get("sdr_index", 0))
        r = int(cfg.get("sdr_sample_rate", self.DEFAULT_RATE))
        if r not in self.RATES:
            self.log(f"SR {r} → {self.DEFAULT_RATE} (кратность 12000)")
            r = self.DEFAULT_RATE
        self.req_rate = r
        self.tone_hz = float(CONFIG.get("pulse_center_hz", 2000))
        self.if_bw = float(cfg.get("sdr_if_bw_hz", 1200))
        self.lock_span = float(cfg.get("sdr_lock_span_hz", 1500))
        self.lock_span_locked = float(cfg.get("sdr_lock_span_locked_hz", 300))
        self.carrier_fft = int(cfg.get("sdr_carrier_fft", 32768))
        self.ppm_cal = float(cfg.get("sdr_ppm_cal", 0.0))
        self.extra_latency = float(cfg.get("sdr_latency_ms", 0)) / 1000.0
        self.rx_freq = float(CONFIG.get("rx_freq", 9996000))
        self.center_freq = self.rx_freq - self.tone_hz
        self.fs_in = float(self.req_rate)
        self.fs_dec = 12000.0
        self.decim = 20
        self.sdr = None
        self.thread = None
        self._ready = threading.Event()
        self._lock = threading.Lock()
        self._pending_freq = None
        self._last_err = ""
        self.carrier_offset_hz = 0.0
        self.carrier_level_db = -120.0
        self.locked = False
        self.ppm_measured = 0.0
        self.blocks_ok = 0
        self.read_errors = 0
        self._shift_ph = 0.0
        self._tone_ph = 0.0
        self._dc = 0.0
        self._dc_a = math.exp(-2.0 * math.pi * 20.0 / self.SAMPLE_RATE)
        self._carr = np.zeros(0, dtype=np.complex64)
        self._carr_ts = 0.0
        self.hilbert_h = design_hilbert(201)
        self.hilbert_fir = None
        self.lp_fir = None
        self.hilbert_mode = None
        self.resampler = SincResampler(1.0)

    @staticmethod
    def device_list(log=None):
        if not SDR_AVAILABLE:
            return []
        try:
            return [f"#{i}  {s}" for i, s in
                    enumerate(RtlSdr.get_device_serial_addresses())]
        except Exception as e:
            if log:
                log(f"Список устройств: {e}")
            return ["#0"]

    def get_info(self):
        with self._lock:
            return {"rx_hz": self.rx_freq, "center_hz": self.center_freq,
                    "tone_hz": self.tone_hz, "fs_in": self.fs_in,
                    "fs_dec": self.fs_dec, "if_bw": self.if_bw,
                    "offset_hz": self.carrier_offset_hz,
                    "level_db": self.carrier_level_db, "locked": self.locked,
                    "gain": self.agc_gain, "blocks": self.blocks_ok,
                    "usb_err": self.read_errors, "alive": self.data_alive(),
                    "ppm": self.ppm_measured, "ppm_cal": self.ppm_cal,
                    "fft_bin": self.fs_in / max(1, self.carrier_fft)}

    def set_frequency(self, rx):
        with self._lock:
            self.rx_freq = float(rx)
            self.center_freq = self.rx_freq - self.tone_hz
            self._pending_freq = self.center_freq

    def start(self):
        if not SDR_AVAILABLE:
            self._last_err = "pyrtlsdr недоступен"
            return False
        self._ready.clear()
        self.running = True
        self.thread = threading.Thread(target=self._worker, daemon=True, name="sdr")
        self.thread.start()
        if not self._ready.wait(timeout=12.0):
            self._last_err = self._last_err or "таймаут инициализации"
            self.running = False
            return False
        return self.sdr is not None

    def stop(self):
        self.running = False
        if self.thread:
            try:
                self.thread.join(timeout=3.0)
            except Exception:
                pass
            self.thread = None

    def close(self):
        self.stop()

    def _build_dsp(self):
        self.decim = max(1, int(round(self.fs_in / self.SAMPLE_RATE)))
        self.fs_dec = self.fs_in / self.decim
        nyq = self.fs_dec * 0.5
        cut = min(self.if_bw, nyq * 0.8)
        tr = max(300.0, min(nyq - cut, 2.0 * self.tone_hz - cut))
        self.lp_fir = FIRFilter(design_fir_lp(cut, self.fs_in, transition=tr), True)
        self.hilbert_fir = FIRFilter(self.hilbert_h, False)
        self.hilbert_mode = None
        self.resampler = SincResampler(self.fs_dec / self.SAMPLE_RATE)
        self._shift_ph = self._tone_ph = self._dc = 0.0
        self._peak_hold = 0.0
        self._carr = np.zeros(0, dtype=np.complex64)
        pipe = (self.lp_fir.delay_samples + (len(self.hilbert_h) - 1) // 2) / self.fs_in
        self.latency_sec = pipe + self.extra_latency

    def _analytic(self, iq):
        re, im = iq.real.astype(np.float64), iq.imag.astype(np.float64)
        rr = float(np.sqrt(np.mean(re * re)))
        ri = float(np.sqrt(np.mean(im * im)))
        if rr <= 1e-12 and ri <= 1e-12:
            return iq.astype(np.complex128)
        big, small = max(rr, ri), min(rr, ri)
        if big > 0 and small / big >= 0.20:
            if self.hilbert_mode != "c":
                self.hilbert_mode = "c"
                self.hilbert_fir.reset()
            return iq.astype(np.complex128)
        x = im if ri > rr else re
        if self.hilbert_mode != "r":
            self.hilbert_mode = "r"
            self.hilbert_fir.reset()
        hil = self.hilbert_fir.process(x)
        if len(hil) == 0:
            return np.zeros(0, dtype=np.complex128)
        d = self.hilbert_fir.delay_samples
        buf = np.concatenate([np.zeros(d), x])
        xd = buf[d:d + len(hil)]
        n = min(len(xd), len(hil))
        return (xd[:n] + 1j * hil[:n]).astype(np.complex128)

    def _carrier_lock(self, z):
        try:
            self._carr = np.concatenate([self._carr, z.astype(np.complex64)])
            need = max(4096, self.carrier_fft)
            if len(self._carr) > need * 2:
                self._carr = self._carr[-need:]
            if len(self._carr) < need:
                return
            now = time.time()
            if now - self._carr_ts < 0.25:
                return
            self._carr_ts = now
            seg = self._carr[-need:]
            sp = np.abs(np.fft.fft(seg * np.hanning(need)))
            fr = np.fft.fftfreq(need, 1.0 / self.fs_in)
            with self._lock:
                span = self.lock_span_locked if self.locked else self.lock_span
                cur = self.carrier_offset_hz
            idx = np.nonzero(np.abs(fr - (self.tone_hz + cur)) <= span)[0]
            if idx.size == 0:
                return
            k = int(idx[int(np.argmax(sp[idx]))])
            delta = 0.0
            if 0 < k < need - 1:
                a, b, c = sp[k - 1], sp[k], sp[k + 1]
                den = a - 2 * b + c
                if abs(den) > 1e-20:
                    delta = max(-0.5, min(0.5, 0.5 * (a - c) / den))
            fpk = (k + delta) * self.fs_in / need
            snr = float(sp[k] / (float(np.mean(sp)) + 1e-15))
            with self._lock:
                self.carrier_level_db = 20.0 * math.log10(float(sp[k]) + 1e-15)
                if snr > 8.0:
                    off = fpk - self.tone_hz - self.ppm_cal * self.rx_freq / 1e6
                    self.carrier_offset_hz = 0.9 * self.carrier_offset_hz + 0.1 * off
                    self.locked = True
                elif snr < 3.0:
                    self.locked = False
                self.ppm_measured = (self.carrier_offset_hz +
                                     self.ppm_cal * self.rx_freq / 1e6) / self.rx_freq * 1e6
        except Exception:
            pass

    def _dc_block(self, x):
        out = np.empty(len(x))
        dc, a = self._dc, self._dc_a
        oma = 1.0 - a
        for i in range(len(x)):
            v = float(x[i])
            dc = a * dc + oma * v
            out[i] = v - dc
        self._dc = dc
        return out

    def _worker(self):
        tp = 2.0 * math.pi
        try:
            try:
                self.sdr = RtlSdr(device_index=self.device_index)
            except Exception:
                self.sdr = RtlSdr()
            self.sdr.sample_rate = self.req_rate
            self.fs_in = float(self.sdr.sample_rate)
            self.sdr.set_direct_sampling(2)
            self.sdr.center_freq = int(round(self.center_freq))
            try:
                self.sdr.set_tuner_gain_mode(0)
            except Exception:
                pass
            self._build_dsp()
            self.sdr.read_samples(256)
            self.last_data_ts = time.time()
            ex = "ровно" if abs(self.fs_dec - self.SAMPLE_RATE) < 1e-6 else "дробно!"
            self.log(f"✓ USB {self.center_freq/1000:.1f} ← RWM {self.rx_freq/1000:.1f} кГц "
                     f"→ {self.tone_hz:.0f} Гц | ФНЧ ±{self.if_bw:.0f} Гц | "
                     f"SR {self.fs_in:.0f} → {self.fs_dec:.0f} ({ex}) | "
                     f"БПФ {self.carrier_fft} → {self.fs_in/self.carrier_fft:.1f} Гц/бин | "
                     f"задержка тракта {self.latency_sec*1000:.2f} мс")
        except Exception as e:
            self._last_err = str(e)
            self.log(f"Ошибка инициализации: {e}")
            try:
                if self.sdr:
                    self.sdr.close()
            except Exception:
                pass
            self.sdr = None
            self.running = False
            self._ready.set()
            return
        self._ready.set()

        while self.running:
            with self._lock:
                f, self._pending_freq = self._pending_freq, None
            if f is not None:
                try:
                    self.sdr.center_freq = int(round(f))
                    self.sdr.read_samples(256)
                    with self._lock:
                        self.carrier_offset_hz = 0.0
                        self.locked = False
                    self._shift_ph = self._tone_ph = 0.0
                    self._peak_hold = 0.0
                    self._carr = np.zeros(0, dtype=np.complex64)
                    self.lp_fir.reset()
                    self.log(f"USB {f/1000:.1f} ← RWM {(f+self.tone_hz)/1000:.1f} кГц")
                except Exception as e:
                    self.log(f"Перестройка: {e}")
            n = min(self.READ_MAX, max(self.READ_MIN, int(self.fs_in * self.READ_BLOCK_SEC)))
            try:
                smp = self.sdr.read_samples(n)
            except Exception as e:
                self.read_errors += 1
                self._last_err = str(e)
                if self.read_errors <= 3:
                    self.log(f"Ошибка чтения #{self.read_errors}: {e}")
                if self.read_errors >= 10:
                    self.log("Много ошибок чтения — остановка")
                    break
                time.sleep(0.1)
                continue
            if smp is None or len(smp) == 0:
                continue
            self.read_errors = 0
            self.blocks_ok += 1
            try:
                z = self._analytic(np.asarray(smp, dtype=np.complex64))
                if len(z) == 0:
                    continue
                self._carrier_lock(z)
                with self._lock:
                    off, ppm = self.carrier_offset_hz, self.ppm_cal
                shift = self.tone_hz + off + ppm * self.rx_freq / 1e6
                N = len(z)
                inc = -tp * shift / self.fs_in
                ph = self._shift_ph + inc * np.arange(N, dtype=np.float64)
                self._shift_ph = (self._shift_ph + inc * N) % tp
                z = z * np.exp(1j * ph)
                zf = self.lp_fir.process(z)
                if len(zf) < self.decim + 2:
                    continue
                zd = zf[::self.decim]
                m = len(zd)
                ti = tp * self.tone_hz / self.fs_dec
                tph = self._tone_ph + ti * np.arange(m, dtype=np.float64)
                self._tone_ph = (self._tone_ph + ti * m) % tp
                a = (zd * np.exp(1j * tph)).real
                if abs(self.fs_dec - self.SAMPLE_RATE) > 1e-6:
                    a = self.resampler.process(a)
                if len(a) < 8:
                    continue
                a = self._dc_block(a)
                a = self._apply_peak_agc(a)
                self._append(np.clip(a * 32767.0, -32768, 32767).astype(np.int16))
            except Exception:
                continue
        try:
            if self.sdr:
                self.sdr.close()
        except Exception:
            pass
        self.sdr = None
        self.running = False


class SourceManager:
    MODES = ("Audio", "SDR", "Auto")

    def __init__(self, log_func=None):
        self.log = log_func or (lambda m: print(m))
        self.audio = self.sdr = None
        self.mode = CONFIG.get("source_mode", "SDR")
        self.active_name = "sdr" if self.mode == "SDR" else "audio"
        self._hold = 0
        self._ts = 0.0
        self._lock = threading.Lock()

    def start_audio(self, dev, agc):
        self.stop_audio()
        if not AUDIO_AVAILABLE or not dev:
            self.log("[ИСТОЧНИК] аудио недоступно или устройство не выбрано")
            return False
        s = AudioInput(dev, agc=agc)
        if s.start():
            self.audio = s
            self.log(f"[ИСТОЧНИК] ① Аудио: {s.device_name}")
            return True
        s.close()
        self.log(f"[ИСТОЧНИК] не удалось запустить аудио: {dev}")
        return False

    def stop_audio(self):
        if self.audio:
            try:
                self.audio.close()
            except Exception:
                pass
            self.audio = None

    def start_sdr(self):
        self.stop_sdr()
        if not SDR_AVAILABLE:
            self.log("[ИСТОЧНИК] pyrtlsdr недоступен")
            return False
        s = SDRInput(CONFIG, log_func=lambda m: self.log(f"[SDR] {m}"))
        if s.start():
            self.sdr = s
            self.log("[ИСТОЧНИК] ② RTL-SDR USB запущен")
            return True
        err = s._last_err
        s.close()
        self.log(f"[ИСТОЧНИК] RTL-SDR: {err}")
        return False

    def stop_sdr(self):
        if self.sdr:
            try:
                self.sdr.close()
            except Exception:
                pass
            self.sdr = None

    def stop_all(self):
        self.stop_audio()
        self.stop_sdr()

    def set_mode(self, m):
        if m not in self.MODES:
            return
        self.mode = m
        CONFIG["source_mode"] = m
        save_config(CONFIG)
        with self._lock:
            self.active_name = {"Audio": "audio", "SDR": "sdr"}.get(m, self.active_name)
            if m == "Auto":
                self._hold = 0
        self.log(f"[ИСТОЧНИК] режим: {m}")

    def tick(self):
        if self.mode != "Auto":
            return
        now = time.time()
        if now - self._ts < 1.0:
            return
        self._ts = now
        t = float(CONFIG.get("pulse_center_hz", 2000))
        qa = self.audio.quality(t) if (self.audio and self.audio.running) else -1
        qs = self.sdr.quality(t) if (self.sdr and self.sdr.running) else -1
        with self._lock:
            if qa < 0 and qs < 0:
                return
            want = "audio" if qa >= qs else "sdr"
            if want == self.active_name:
                self._hold = 0
                return
            if max(qa, qs) > min(qa, qs) * 1.25 + 1e-9:
                self._hold += 1
                if self._hold >= 3:
                    self.active_name = want
                    self._hold = 0
                    self.log(f"[ИСТОЧНИК] авто → {'① Аудио' if want=='audio' else '② RTL-SDR'}")
            else:
                self._hold = 0

    def get_active(self):
        with self._lock:
            n = self.active_name
        if n == "sdr" and self.sdr and self.sdr.running:
            return self.sdr
        if self.audio and self.audio.running:
            return self.audio
        if self.sdr and self.sdr.running:
            return self.sdr
        return None

    def active_title(self):
        s = self.get_active()
        return s.title if s else "нет источника"

    def any_running(self):
        return bool((self.audio and self.audio.running) or (self.sdr and self.sdr.running))


class AudioMonitor:
    RATES_TRY = [48000, 44100, 32000, 24000, 12000]
    CHUNK = 1024

    def __init__(self, sources, log_func=None):
        self.sources = sources
        self.log = log_func or (lambda m: print(f"[МОНИТОР] {m}"))
        self.enabled = bool(CONFIG.get("monitor_enabled", False))
        self.volume = float(CONFIG.get("monitor_volume", 0.75))
        self.device_name = CONFIG.get("monitor_output", "")
        self.out_rate = int(CONFIG.get("monitor_out_rate", 48000))
        self.playing = False
        self.out_rms = 0.0
        self.last_error = ""
        self._stop = threading.Event()
        self._thread = None
        self._pya = None
        self._stream = None
        self._cur = None
        self._pos = 0
        self._rs = SincResampler(0.25)
        self._lock = threading.Lock()

    def start(self):
        if self._thread:
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True, name="monitor")
        self._thread.start()

    def set_enabled(self, on):
        self.enabled = bool(on)
        if on:
            self.last_error = ""
        else:
            self.playing = False
            self.out_rms = 0.0

    def set_volume(self, v):
        self.volume = max(0.0, min(1.0, float(v)))

    def set_device(self, n):
        with self._lock:
            self.device_name = n or ""
            self._close()
            self._cur = None

    def _dev_idx(self):
        if self._pya is None or not self.device_name:
            return None
        try:
            for i in range(self._pya.get_device_count()):
                inf = self._pya.get_device_info_by_index(i)
                if inf.get("maxOutputChannels", 0) > 0 and inf.get("name") == self.device_name:
                    return i
        except Exception:
            pass
        return None

    def _open(self):
        self._pya = pyaudio.PyAudio()
        idx = self._dev_idx()
        pref = int(CONFIG.get("monitor_out_rate", 48000))
        last = None
        for r in [pref] + [x for x in self.RATES_TRY if x != pref]:
            try:
                self._stream = self._pya.open(
                    format=pyaudio.paInt16, channels=1, rate=int(r), output=True,
                    output_device_index=idx, frames_per_buffer=self.CHUNK)
                self.out_rate = int(r)
                self._rs = SincResampler(BufferedSource.SAMPLE_RATE / float(r))
                self.log(f"Вывод → {self.device_name or 'по умолчанию'} · {r} Гц "
                         f"(×{r/BufferedSource.SAMPLE_RATE:g}, sinc {SincResampler.TAPS} отв.)")
                return
            except Exception as e:
                last = e
                self._stream = None
        raise RuntimeError(f"выход не открыт: {last}")

    def _close(self):
        if self._stream:
            try:
                self._stream.stop_stream()
            except Exception:
                pass
            try:
                self._stream.close()
            except Exception:
                pass
            self._stream = None
        if self._pya:
            try:
                self._pya.terminate()
            except Exception:
                pass
            self._pya = None
        self.playing = False
        self.out_rms = 0.0

    def _loop(self):
        while not self._stop.is_set():
            if not self.enabled or not AUDIO_AVAILABLE:
                self._close()
                time.sleep(0.1)
                continue
            if self._stream is None:
                try:
                    self._open()
                    self._cur = None
                except Exception as e:
                    self.last_error = str(e)
                    self.log(f"Ошибка выхода: {e}")
                    self.enabled = False
                    time.sleep(0.5)
                    continue
            src = self.sources.get_active()
            if src is None or not src.running:
                self.playing = False
                self.out_rms = 0.0
                time.sleep(0.05)
                continue
            if src.name != self._cur:
                self._cur = src.name
                self._pos = src.total_samples
                self._rs.reset()
                self.log(f"Источник монитора: {src.title}")
            data, self._pos = src.get_new_samples(self._pos)
            if data is None or len(data) == 0:
                self.playing = False
                time.sleep(0.012)
                continue
            y = self._rs.process(data.astype(np.float32) * np.float32(self.volume))
            if len(y) == 0:
                continue
            self.out_rms = min(1.0, float(np.sqrt(np.mean(
                np.square(y.astype(np.float64))))) / 32768.0)
            try:
                self._stream.write(np.clip(y, -32768, 32767).astype(np.int16).tobytes())
                self.playing = True
            except Exception as e:
                self.last_error = str(e)
                self.log(f"Ошибка вывода: {e}")
                self._close()
                time.sleep(0.3)
        self._close()

    def stop(self):
        self._stop.set()
        self.enabled = False
        if self._thread:
            try:
                self._thread.join(timeout=2.0)
            except Exception:
                pass
            self._thread = None
        self._close()


class NTPServer:
    EPOCH = 2208988800

    def __init__(self, log_func=None, status_func=None):
        self.log = log_func or (lambda m: print(f"[NTP-SRV] {m}"))
        self.status_func = status_func or (lambda: {})
        self.enabled = bool(CONFIG.get("ntp_server_enabled", False))
        self.port = int(CONFIG.get("ntp_server_port", 123))
        self.bind = CONFIG.get("ntp_server_bind", "0.0.0.0")
        self.require_sync = bool(CONFIG.get("ntp_server_require_sync", False))
        self.allow_raw = CONFIG.get("ntp_server_allow", "")
        self._nets = self._parse(self.allow_raw)
        self._stop = threading.Event()
        self._thread = None
        self._sock = None
        self._lock = threading.Lock()
        self.running = False
        self.last_error = ""
        self.requests = self.served = self.denied = 0
        self.last_client = ""
        try:
            res = float(time.get_clock_info("time").resolution)
            p = int(math.floor(math.log2(res))) if res > 0 else -10
        except Exception:
            p = -10
        self.precision = max(-30, min(0, p))

    @staticmethod
    def _parse(raw):
        nets = []
        for t in str(raw or "").replace(";", ",").split(","):
            t = t.strip()
            if not t:
                continue
            try:
                if "/" not in t:
                    t += "/32" if ":" not in t else "/128"
                nets.append(ipaddress.ip_network(t, strict=False))
            except Exception:
                pass
        return nets

    def _allowed(self, ip):
        if not self._nets:
            return True
        try:
            a = ipaddress.ip_address(ip)
        except Exception:
            return False
        return any(a in n for n in self._nets)

    def configure(self, **kw):
        if "enabled" in kw:
            self.enabled = bool(kw["enabled"])
        if "port" in kw:
            self.port = int(kw["port"])
        if "bind" in kw:
            self.bind = str(kw["bind"])
        if "require_sync" in kw:
            self.require_sync = bool(kw["require_sync"])
        if "allow" in kw:
            self.allow_raw = str(kw["allow"])
            self._nets = self._parse(self.allow_raw)

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True, name="ntp")
        self._thread.start()

    def restart(self):
        self.stop()
        time.sleep(0.15)
        self.start()

    def stop(self):
        self._stop.set()
        with self._lock:
            if self._sock:
                try:
                    self._sock.close()
                except Exception:
                    pass
                self._sock = None
        if self._thread:
            try:
                self._thread.join(timeout=2.0)
            except Exception:
                pass
            self._thread = None
        self.running = False

    def get_stats(self):
        with self._lock:
            return {"enabled": self.enabled, "running": self.running,
                    "port": self.port, "bind": self.bind, "requests": self.requests,
                    "served": self.served, "denied": self.denied,
                    "last_client": self.last_client, "error": self.last_error,
                    "precision": self.precision}

    @classmethod
    def _ts(cls, t):
        if not t or t <= 0:
            return 0
        w = math.floor(t)
        return (((int(w) + cls.EPOCH) & 0xFFFFFFFF) << 32) | \
            (int((t - w) * 4294967296.0) & 0xFFFFFFFF)

    def _reply(self, data, t_recv):
        if len(data) < 48:
            return None
        fl = data[0]
        vn = (fl >> 3) & 7
        if (fl & 7) not in (0, 1, 3):
            return None
        vn = min(max(vn or 4, 1), 4)
        try:
            poll = struct.unpack("!b", data[2:3])[0]
        except Exception:
            poll = 6
        poll = max(4, min(17, poll or 6))
        orig = int.from_bytes(data[40:48], "big")
        st = self.status_func() or {}
        ra, na, off = st.get("rwm_age"), st.get("ntp_age"), st.get("offset_ms")
        if ra is not None and ra < 180:
            stratum, ref, rd, bd, age = 1, b"RWM", 0.0, 0.002, ra
        elif na is not None and na < 900:
            stratum, ref, rd, bd, age = 2, b"NTP", 0.02, 0.010, na
        else:
            stratum, ref, rd, bd, age = 3, b"LOCL", 0.0, 0.050, 3600.0
        li = 0
        if self.require_sync and stratum >= 3:
            li, stratum = 3, 16
        disp = bd + age * 0.00005
        if off is not None:
            disp += abs(float(off)) / 1000.0
        disp = min(disp, 16.0)
        return struct.pack("!BBbbII4sQQQQ", (li << 6) | (vn << 3) | 4,
                           stratum, poll, self.precision,
                           int(rd * 65536) & 0xFFFFFFFF, int(disp * 65536) & 0xFFFFFFFF,
                           ref, self._ts(st.get("last_sync_ts") or t_recv), orig,
                           self._ts(t_recv), self._ts(time.time()))

    def _loop(self):
        while not self._stop.is_set():
            if not self.enabled:
                with self._lock:
                    if self._sock:
                        try:
                            self._sock.close()
                        except Exception:
                            pass
                        self._sock = None
                    self.running = False
                time.sleep(0.2)
                continue
            if self._sock is None:
                try:
                    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                    s.bind((self.bind, self.port))
                    s.settimeout(0.4)
                    with self._lock:
                        self._sock = s
                        self.running = True
                        self.last_error = ""
                    self.log(f"Слушаю {self.bind}:{self.port} UDP · 2^{self.precision} c"
                             + (f" · фильтр {self.allow_raw}" if self._nets else ""))
                except Exception as e:
                    self.last_error = str(e)
                    self.log(f"Не могу занять {self.bind}:{self.port} — {e}")
                    if self.port == 123 and sys.platform == "win32":
                        self.log("Порт 123 держит Windows Time: net stop w32time "
                                 "(или порт 1123)")
                    self.enabled = False
                    with self._lock:
                        self.running = False
                    time.sleep(1.0)
                    continue
            try:
                data, addr = self._sock.recvfrom(1024)
            except socket.timeout:
                continue
            except OSError:
                with self._lock:
                    self._sock = None
                    self.running = False
                time.sleep(0.2)
                continue
            except Exception:
                time.sleep(0.05)
                continue
            tr = time.time()
            with self._lock:
                self.requests += 1
                self.last_client = f"{addr[0]}:{addr[1]}"
            if not self._allowed(addr[0]):
                with self._lock:
                    self.denied += 1
                continue
            try:
                r = self._reply(data, tr)
                if r:
                    self._sock.sendto(r, addr)
                    with self._lock:
                        self.served += 1
            except Exception:
                pass
        with self._lock:
            if self._sock:
                try:
                    self._sock.close()
                except Exception:
                    pass
                self._sock = None
            self.running = False

    def self_test(self, host="127.0.0.1", timeout=2.0):
        p = bytearray(48)
        p[0] = (4 << 3) | 3
        t0 = time.time()
        p[40:48] = self._ts(t0).to_bytes(8, "big")
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(timeout)
        try:
            s.sendto(bytes(p), (host, self.port))
            d, _ = s.recvfrom(1024)
        finally:
            s.close()
        t3 = time.time()
        if len(d) < 48:
            raise RuntimeError("короткий ответ")
        u = struct.unpack("!BBbbII4sQQQQ", d[:48])
        t1 = ((u[9] >> 32) - self.EPOCH) + (u[9] & 0xFFFFFFFF) / 2**32
        t2 = ((u[10] >> 32) - self.EPOCH) + (u[10] & 0xFFFFFFFF) / 2**32
        return {"stratum": u[1], "refid": u[6].decode("ascii", "replace").strip("\x00"),
                "li": (u[0] >> 6) & 3,
                "offset_ms": ((t1 - t0) + (t2 - t3)) / 2 * 1000,
                "delay_ms": ((t3 - t0) - (t2 - t1)) * 1000}


class DirectCAT:
    def __init__(self):
        self.ser = None
        self.lock = threading.Lock()

    def configure(self, port, baud=19200):
        self.close()
        if not port:
            return True
        if not SERIAL_AVAILABLE:
            return False
        try:
            self.ser = serial.Serial(port=port, baudrate=baud, timeout=0.1, write_timeout=0.5)
            return True
        except Exception:
            return False

    def write(self, d):
        if self.ser is None or not self.ser.is_open:
            return False
        with self.lock:
            try:
                self.ser.write(d)
                return True
            except Exception:
                return False

    def close(self):
        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
        self.ser = None


class SerialSniffer:
    def __init__(self, real, virt, baud, log):
        self.real_port, self.virtual_port, self.baudrate, self.log = real, virt, baud, log
        self.real = self.virtual = None
        self.running = False

    def start(self):
        if not SERIAL_AVAILABLE:
            raise RuntimeError("pyserial недоступен")
        try:
            self.real = serial.Serial(self.real_port, self.baudrate, timeout=0.05, write_timeout=0.5)
            self.virtual = serial.Serial(self.virtual_port, self.baudrate, timeout=0.05, write_timeout=0.5)
        except Exception:
            self.stop()
            raise
        self.running = True
        threading.Thread(target=self._v2r, daemon=True).start()
        threading.Thread(target=self._r2v, daemon=True).start()
        if self.log:
            self.log(f"Sniffer: {self.virtual_port} ↔ {self.real_port}")

    def write_real(self, d):
        if self.real and self.real.is_open:
            try:
                self.real.write(d)
                return True
            except Exception:
                return False
        return False

    def _v2r(self):
        while self.running:
            try:
                d = self.virtual.read(4096)
                if d:
                    self.write_real(d)
            except Exception:
                time.sleep(0.2)

    def _r2v(self):
        while self.running:
            try:
                d = self.real.read(4096)
                if d and self.virtual and self.virtual.is_open:
                    self.virtual.write(d)
            except Exception:
                time.sleep(0.2)

    def stop(self):
        self.running = False
        time.sleep(0.1)
        for p in (self.real, self.virtual):
            try:
                if p and p.is_open:
                    p.close()
            except Exception:
                pass
        self.real = self.virtual = None


# ------------------------------------------------------------
# ДЕКОДЕР RWM по ITU-R TF.2487-0
# ------------------------------------------------------------
class RWMDecoder:
    """
    Причинный тракт: КИХ-полосовик → |·| → КИХ-сглаживание.
    Групповые задержки известны и вычитаются.
    Опорная точка — НАЧАЛО ПЕРЕДНЕГО ФРОНТА: линейная экстраполяция
    по точкам 20 % и 80 % амплитуды к нулю (субсэмплово).
    Валидация — по сетке UTC, класс метки — по положению в сетке,
    пропуски секунд 56–59 при minute % 5 == 4 учитываются.
    DUT1/dUT1 — по числу импульсов в секундных слотах.
    """
    SR = BufferedSource.SAMPLE_RATE

    def __init__(self, log_func=None):
        self.log = log_func
        self.tone = float(CONFIG.get("pulse_center_hz", 2000))
        self.band = float(CONFIG.get("pulse_band_hz", 600))
        self.smooth_ms = float(CONFIG.get("pulse_smooth_ms", 1.0))
        self.lo_frac = float(CONFIG.get("edge_lo_frac", 0.2))
        self.hi_frac = float(CONFIG.get("edge_hi_frac", 0.8))
        self.gate_coarse = float(CONFIG.get("gate_coarse_ms", 450)) / 1000.0
        self.gate_fine = float(CONFIG.get("gate_fine_ms", 80)) / 1000.0
        self.min_ms = float(CONFIG.get("pulse_min_ms", 8))
        self.max_ms = float(CONFIG.get("pulse_max_ms", 700))
        self.dur_tol = float(CONFIG.get("dur_tolerance", 0.45))
        self.path_ms = float(CONFIG.get("path_delay_ms", 0))
        self.trim_ms = float(CONFIG.get("timing_trim_ms", 0))
        self.decode_dut1 = bool(CONFIG.get("decode_dut1", True))

        self._lock = threading.Lock()
        self._pos = {}
        self._env = {}
        self._env_start = {}
        self._last_end = {}
        self._resid = deque(maxlen=400)
        self._minute = deque(maxlen=20)
        self._dur = {"second": deque(maxlen=200), "tenth": deque(maxlen=300),
                     "minute": deque(maxlen=60)}
        self._fine = False
        self._settle = 0.0

        self.accepted = self.rejected = self.minute_marks = self.skipped_ok = 0
        self.last_snr = self.last_thr = self.last_noise = self.last_max = 0.0
        self.last_dur_ms = 0.0
        self.last_class = "—"
        self.last_pulse_ts = self.last_minute_ts = 0.0
        self.last_source = "—"
        self.hit_rate = 0.0
        self._diag = 0.0
        self.measure_until = 0.0

        # DUT1
        self._sec_counts = defaultdict(int)
        self._minute_anchor = None
        self.dut1 = None
        self.dut1_detail = ""

        self.bp = None
        self.ma = None
        self.fixed_delay_s = 0.0
        self._build()

    def _build(self):
        tr = max(self.band * 0.5, 150.0)
        n = max(31, min(511, int(3.5 * self.SR / tr))) | 1
        self.bp = FIRFilter(design_fir_bp(self.tone, self.band, self.SR, n), False)
        w = max(3, int(self.SR * self.smooth_ms / 1000.0))
        self.ma = FIRFilter(np.ones(w) / w, False)
        self.fixed_delay_s = (self.bp.delay_samples + self.ma.delay_samples) / float(self.SR)
        with self._lock:
            self._env.clear()
            self._env_start.clear()

    def total_delay(self, src):
        return (self.fixed_delay_s + (src.latency_sec if src else 0.0)
                + self.path_ms / 1000.0 + self.trim_ms / 1000.0)

    @staticmethod
    def _hyst(env, hi, lo, gap):
        st, en = [], []
        inp, s = False, 0
        for i, v in enumerate(env):
            if not inp:
                if v >= hi:
                    inp, s = True, i
            elif v <= lo:
                inp = False
                st.append(s)
                en.append(i)
        if not st:
            return [], []
        ms, me = [], []
        for a, b in zip(st, en):
            if ms and (a - me[-1]) <= gap:
                me[-1] = max(me[-1], b)
            else:
                ms.append(a)
                me.append(b)
        return ms, me

    def _cross(self, seg, idx_from, level, backward=True):
        """Субсэмпловый поиск пересечения уровня при движении от пика назад."""
        i = idx_from
        if backward:
            while i > 0 and seg[i] > level:
                i -= 1
            if i >= idx_from:
                return float(idx_from)
            a, b = float(seg[i]), float(seg[i + 1])
            fr = (level - a) / (b - a) if b != a else 0.0
            return i + max(0.0, min(1.0, fr))
        return float(idx_from)

    def _rising_start(self, env, s, e):
        """Начало переднего фронта: экстраполяция 20/80 % → 0."""
        seg = env[s:e]
        if len(seg) < 4:
            return float(s)
        pk = int(np.argmax(seg))
        pv = float(seg[pk])
        if pv <= 0:
            return float(s)
        i_hi = self._cross(seg, pk, self.hi_frac * pv)
        i_lo = self._cross(seg, pk, self.lo_frac * pv)
        if i_hi <= i_lo:
            return s + i_lo
        span = i_hi - i_lo
        return s + i_lo - span * (self.lo_frac / max(1e-9, self.hi_frac - self.lo_frac))

    def _finalize_dut1(self, anchor):
        """DUT1/dUT1 по маркированным секундам завершившейся минуты."""
        try:
            n_pos = sum(1 for k in range(1, 9) if self._sec_counts.get(anchor + k, 0) >= 2)
            n_neg = sum(1 for k in range(9, 17) if self._sec_counts.get(anchor + k, 0) >= 2)
            p_pos = sum(1 for k in range(21, 25) if self._sec_counts.get(anchor + k, 0) >= 2)
            q_neg = sum(1 for k in range(31, 35) if self._sec_counts.get(anchor + k, 0) >= 2)
            dut1 = 0.1 * n_pos - 0.1 * n_neg
            dut1_fine = 0.02 * p_pos - 0.02 * q_neg
            self.dut1 = dut1 + dut1_fine
            self.dut1_detail = (f"DUT1={dut1:+.1f} c (n={n_pos}, m={n_neg}) · "
                                f"dUT1={dut1_fine:+.2f} c (p={p_pos}, q={q_neg})")
            if self.log:
                self.log(f"[DUT1] UT1−UTC = {self.dut1:+.3f} c · {self.dut1_detail}")
        except Exception:
            pass

    def process(self, src, mode):
        if src is None or not src.running:
            return
        step = RWMSchedule.grid_step(mode)
        if step is None:
            return
        if time.time() < self._settle:
            return
        name = src.name
        with self._lock:
            if name not in self._pos:
                self._pos[name] = max(0, src.total_samples - self.SR // 2)
                self._env[name] = np.zeros(0)
                self._env_start[name] = self._pos[name]
                self._last_end[name] = -1
            pos = self._pos[name]
        new, pos2 = src.read_from(pos, self.SR)
        if new is None or len(new) < 32:
            return
        with self._lock:
            self._pos[name] = pos2

        x = new.astype(np.float64) / 32768.0
        e = self.ma.process(np.abs(self.bp.process(x)))
        if len(e) == 0:
            return
        with self._lock:
            env = np.concatenate([self._env.get(name, np.zeros(0)), e])
            est = self._env_start.get(name, pos)
            if len(env) > self.SR * 3:
                cut = len(env) - self.SR * 3
                env = env[cut:]
                est += cut
            self._env[name] = env
            self._env_start[name] = est

        mx = float(np.max(env))
        noise = float(np.median(env))
        snr = mx / (noise + 1e-6)
        thr = max(noise + 0.25 * (mx - noise), noise * 2.0, 0.008)
        with self._lock:
            self.last_max, self.last_noise = mx, noise
            self.last_snr, self.last_thr = snr, thr
            self.last_source = name

        starts, ends = self._hyst(env, thr, thr * 0.5, int(0.002 * self.SR))
        with self._lock:
            gate = self.gate_fine if self._fine else self.gate_coarse
        gate = min(gate, step * 0.45)
        delay = self.total_delay(src)
        measuring = time.time() < self.measure_until
        n_acc = 0
        last_end = None

        for s, eidx in zip(starts, ends):
            dur_ms = (eidx - s) / self.SR * 1000.0
            abs_s = est + s
            with self._lock:
                if abs_s <= self._last_end.get(name, -1):
                    continue
            last_end = est + eidx
            if not (self.min_ms <= dur_ms <= self.max_ms):
                with self._lock:
                    self.rejected += 1
                continue
            edge = self._rising_start(env, s, eidx)
            abs_edge = est + edge
            i0 = int(math.floor(abs_edge))
            t = src.sample_to_time(i0)
            if t is None:
                continue
            t += (abs_edge - i0) / self.SR
            t -= delay
            k = round(t / step)
            grid_t = k * step
            resid = grid_t - t
            if abs(resid) > gate:
                with self._lock:
                    self.rejected += 1
                continue
            if not RWMSchedule.pulse_expected(grid_t):
                with self._lock:
                    self.skipped_ok += 1
                    self.rejected += 1
                continue
            exp_ms, cls = RWMSchedule.expected_duration_ms(mode, grid_t)
            if abs(dur_ms - exp_ms) > exp_ms * self.dur_tol + 8.0:
                with self._lock:
                    self.rejected += 1
                continue
            now = time.time()
            with self._lock:
                self._last_end[name] = est + eidx
                self._resid.append((now, resid))
                self._dur[cls].append(dur_ms)
                self.accepted += 1
                self.last_pulse_ts = now
                self.last_dur_ms = dur_ms
                self.last_class = cls
                if self.decode_dut1 and mode == RWMSchedule.A1X:
                    self._sec_counts[int(round(grid_t))] += 1
                if cls == "minute":
                    mres = round(t / 60.0) * 60.0 - t
                    if abs(mres) < 1.0:
                        self._minute.append((now, mres))
                        self.minute_marks += 1
                        self.last_minute_ts = now
                        anchor = int(round(grid_t))
                        if self.decode_dut1 and mode == RWMSchedule.A1X:
                            if self._minute_anchor is not None:
                                self._finalize_dut1(self._minute_anchor)
                            self._minute_anchor = anchor
                            for kk in list(self._sec_counts):
                                if kk < anchor - 120:
                                    del self._sec_counts[kk]
            n_acc += 1
            if cls == "minute" and self.log:
                self.log(f"[ДЕКОДЕР:{name}] ★ МИНУТНАЯ МЕТКА 500 мс ★ "
                         f"{datetime.fromtimestamp(t, tz=timezone.utc).strftime('%H:%M:%S')} · "
                         f"изм {dur_ms:.0f} мс · SNR {snr:.1f}")

        with self._lock:
            if last_end is not None and last_end > self._env_start[name]:
                cut = int(last_end - self._env_start[name])
                self._env[name] = self._env[name][cut:]
                self._env_start[name] += cut
            exp_n = max(1.0, (len(new) / self.SR) / step)
            self.hit_rate = 0.8 * self.hit_rate + 0.2 * min(1.0, n_acc / exp_n)
            if not self._fine and len(self._resid) >= 8:
                arr = np.array([r for _, r in self._resid])
                med = float(np.median(arr))
                mad = float(np.median(np.abs(arr - med)))
                if mad < 0.020:
                    self._fine = True
                    if self.log:
                        self.log(f"[ДЕКОДЕР] захват: окно ±{self.gate_fine*1000:.0f} мс "
                                 f"(MAD {mad*1000:.1f} мс)")

        if measuring and time.time() >= self.measure_until:
            self.measure_until = 0.0
            self._report_durations()

        now = time.time()
        if self.log and now - self._diag > 10.0 and n_acc == 0:
            self._diag = now
            self.log(f"[ДЕКОДЕР:{name}] меток нет · SNR {snr:.1f} · порог {thr:.4f} · "
                     f"окно ±{gate*1000:.0f} мс · пропусков учтено {self.skipped_ok}")

    def _report_durations(self):
        with self._lock:
            d = {k: list(v) for k, v in self._dur.items()}
        parts = []
        for k, exp in (("second", A1X_SECOND_MS), ("tenth", A1N_TENTH_MS),
                       ("minute", A1X_MINUTE_MS)):
            if d[k]:
                a = np.array(d[k])
                parts.append(f"{k}: {float(np.median(a)):.1f} мс "
                             f"(n={len(a)}, ITU {exp:.0f})")
        if parts:
            self.log("[ИЗМЕР] " + " │ ".join(parts))
        else:
            self.log("[ИЗМЕР] меток не набрано")
        self.log(f"[ИЗМЕР] ITU: A1X 100/500 мс · A1N 20/40/500 мс")

    def _fresh(self, dq, age=120.0):
        now = time.time()
        return [v for t, v in dq if now - t <= age]

    def get_fine_offset(self):
        with self._lock:
            v = self._fresh(self._resid)
        if len(v) < 3:
            return None, 0.0, 0
        a = np.array(v)
        off = float(np.median(a))
        mad = float(np.median(np.abs(a - off)))
        q = max(0.0, min(100.0, 100.0 * (1.0 - mad / max(self.gate_fine, 1e-6))))
        return off, q, len(a)

    def get_minute_offset(self):
        with self._lock:
            v = self._fresh(self._minute, 240.0)
        return float(np.median(v)) if v else None

    def get_status(self):
        off, q, n = self.get_fine_offset()
        with self._lock:
            return {"offset_ms": off * 1000 if off is not None else None,
                    "quality": q, "n": n, "accepted": self.accepted,
                    "rejected": self.rejected, "minute_marks": self.minute_marks,
                    "skipped": self.skipped_ok, "hit_rate": self.hit_rate,
                    "last_snr": self.last_snr, "last_dur_ms": self.last_dur_ms,
                    "last_class": self.last_class,
                    "last_pulse_ts": self.last_pulse_ts,
                    "last_minute_ts": self.last_minute_ts,
                    "source": self.last_source, "fine": self._fine,
                    "delay_ms": self.fixed_delay_s * 1000,
                    "path_ms": self.path_ms, "trim_ms": self.trim_ms,
                    "dut1": self.dut1, "dut1_detail": self.dut1_detail}

    def reset(self):
        with self._lock:
            self._pos.clear()
            self._env.clear()
            self._env_start.clear()
            self._last_end.clear()
            self._resid.clear()
            self._minute.clear()
            self._fine = False
            self.accepted = self.rejected = self.minute_marks = self.skipped_ok = 0
            self.hit_rate = 0.0
        if self.bp:
            self.bp.reset()
        if self.ma:
            self.ma.reset()

    def after_correction(self):
        self.reset()
        with self._lock:
            for v in self._dur.values():
                v.clear()
            self._sec_counts.clear()
            self._minute_anchor = None
        self._settle = time.time() + 1.5


class TimeSyncController:
    def __init__(self, dec, log, use_ntp=True, servers=None):
        self.decoder = dec
        self.log = log
        self.use_ntp = use_ntp
        self.servers = servers or ["0.ru.pool.ntp.org"]
        self.running = False
        self.thread = None
        self.correction_count = 0
        self.last_correction_time = 0.0
        self.last_offset_ms = None
        self.sync_quality = 0.0
        self.last_rwm_ts = self.last_ntp_ts = self.last_sync_ts = 0.0
        self._st = self._big = 0.0

    def start(self):
        if self.running:
            return
        self.running = True
        self.thread = threading.Thread(target=self._loop, daemon=True, name="sync")
        self.thread.start()
        self.log("[СИНХРОНИЗАТОР] запущен")

    def stop(self):
        self.running = False
        if self.thread:
            self.thread.join(timeout=1.5)
            self.thread = None
        self.log("[СИНХРОНИЗАТОР] остановлен")

    def server_status(self):
        now = time.time()
        st = self.decoder.get_status()
        r = st.get("last_pulse_ts") or self.last_rwm_ts
        return {"rwm_age": (now - r) if r else None,
                "ntp_age": (now - self.last_ntp_ts) if self.last_ntp_ts else None,
                "last_sync_ts": self.last_sync_ts or r or self.last_ntp_ts or now,
                "offset_ms": self.last_offset_ms, "quality": self.sync_quality,
                "corrections": self.correction_count}

    def _loop(self):
        thr = float(CONFIG.get("sync_threshold_ms", 3.0))
        min_n = int(CONFIG.get("sync_min_pulses", 6))
        while self.running:
            now = time.time()
            st = self.decoder.get_status()
            mode = RWMSchedule.get_mode()
            if now - self._st > 5.0:
                self._st = now
                o = st["offset_ms"]
                ostr = f"{o:+.2f} мс" if o is not None else "—"
                self.log(f"[СТАТУС] {st['source']} | {mode} | "
                         f"меток {st['accepted']}/{st['accepted']+st['rejected']} "
                         f"({st['hit_rate']*100:.0f}%) | минутных {st['minute_marks']} | "
                         f"пропуск {st['skipped']} | SNR {st['last_snr']:.1f} | "
                         f"{st['last_class']} {st['last_dur_ms']:.0f} мс | Δ {ostr}")
            if RWMSchedule.is_sync_possible(mode):
                if st["last_pulse_ts"]:
                    self.last_rwm_ts = st["last_pulse_ts"]
                    self.last_sync_ts = st["last_pulse_ts"]
                fine, q, n = self.decoder.get_fine_offset()
                moff = self.decoder.get_minute_offset()
                self.sync_quality = q
                target = None
                if moff is not None and abs(moff) > 0.4:
                    target = moff
                    self.log(f"[СИНХРОНИЗАТОР] минутная метка: {moff*1000:+.1f} мс (грубо)")
                elif fine is not None and n >= min_n:
                    target = fine
                    self.last_offset_ms = fine * 1000
                if target is not None:
                    if abs(target) > 30:
                        if now - self._big > 30:
                            self._big = now
                            self.log(f"[СИНХРОНИЗАТОР] расхождение {target:.3f} с — нужен NTP")
                    elif abs(target) * 1000 > thr and now - self.last_correction_time > 10:
                        self.log(f"[СИНХРОНИЗАТОР] Δ {target*1000:+.2f} мс · "
                                 f"качество {q:.0f}% · меток {n} · коррекция…")
                        if self._adjust(target):
                            self.decoder.after_correction()
                            self.last_correction_time = time.time()
                            self.last_sync_ts = time.time()
                            self.correction_count += 1
                            self.log(f"[СИНХРОНИЗАТОР] ✓ коррекция "
                                     f"(всего {self.correction_count})")
            if self.use_ntp and now - self.last_correction_time > 300 and \
                    now - self.last_rwm_ts > 120:
                o = self._ntp()
                if o is not None:
                    self.last_ntp_ts = time.time()
                    self.log(f"[NTP] расхождение {o*1000:+.1f} мс")
                    if abs(o) > 0.05 and self._adjust(o):
                        self.decoder.after_correction()
                        self.last_correction_time = time.time()
                        self.last_sync_ts = time.time()
                        self.correction_count += 1
                        self.log(f"[NTP] ✓ коррекция (всего {self.correction_count})")
            time.sleep(0.5)

    def _adjust(self, off):
        try:
            nt = datetime.now(timezone.utc) + timedelta(seconds=off)
            try:
                import win32api
                win32api.SetSystemTime(nt.year, nt.month, nt.isoweekday() % 7, nt.day,
                                       nt.hour, nt.minute, nt.second, nt.microsecond // 1000)
                self.log("[СИСТЕМА] " + nt.strftime('%H:%M:%S.%f')[:-3] + " UTC")
                return True
            except Exception:
                pass
            if not hasattr(ctypes, "windll"):
                raise OSError("установка времени только в Windows")

            class ST(ctypes.Structure):
                _fields_ = [("wYear", ctypes.c_ushort), ("wMonth", ctypes.c_ushort),
                            ("wDayOfWeek", ctypes.c_ushort), ("wDay", ctypes.c_ushort),
                            ("wHour", ctypes.c_ushort), ("wMinute", ctypes.c_ushort),
                            ("wSecond", ctypes.c_ushort), ("wMilliseconds", ctypes.c_ushort)]

            s = ST()
            s.wYear, s.wMonth = nt.year, nt.month
            s.wDayOfWeek = nt.isoweekday() % 7
            s.wDay, s.wHour = nt.day, nt.hour
            s.wMinute, s.wSecond = nt.minute, nt.second
            s.wMilliseconds = nt.microsecond // 1000
            if ctypes.windll.kernel32.SetSystemTime(ctypes.byref(s)) == 0:
                raise ctypes.WinError()
            self.log("[СИСТЕМА] " + nt.strftime('%H:%M:%S.%f')[:-3] + " UTC")
            return True
        except Exception as e:
            self.log(f"[СИСТЕМА] ошибка: {e}")
            return False

    def _ntp(self):
        p = bytearray(48)
        p[0] = 0b00100011
        for srv in self.servers:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.settimeout(2)
            d = None
            try:
                s.sendto(p, (srv, 123))
                d, _ = s.recvfrom(1024)
            except Exception:
                continue
            finally:
                s.close()
            if d and len(d) >= 48:
                try:
                    u = struct.unpack("!12I", d[:48])
                    return (float(u[10]) + float(u[11]) / 2**32 - 2208988800) - time.time()
                except Exception:
                    continue
        return None


# ------------------------------------------------------------
# Виджеты
# ------------------------------------------------------------
class LevelBar(tk.Canvas):
    """НЕ использовать _w/_h — служебные атрибуты Tkinter."""

    def __init__(self, parent, width=250, height=16, segments=32):
        super().__init__(parent, width=width, height=height, bg=C["surface"],
                         highlightthickness=1, highlightbackground=C["border"], bd=0)
        self.segments = segments
        self._cw, self._ch = float(width), float(height)
        self._v = self._pk = 0.0
        self._pk_ts = 0.0
        self._draw()

    def set(self, v):
        v = max(0.0, min(1.0, float(v)))
        self._v = v
        if v >= self._pk:
            self._pk, self._pk_ts = v, time.time()
        elif time.time() - self._pk_ts > 1.2:
            self._pk = max(v, self._pk - 0.02)
        self._draw()

    def _draw(self):
        self.delete("all")
        pad, gap = 2.0, 2.0
        n = self.segments
        w = (self._cw - 2 * pad - (n - 1) * gap) / n
        if w <= 0:
            return
        lit = int(round(self._v * n))
        pki = int(round(self._pk * n)) - 1
        for i in range(n):
            x0 = pad + i * (w + gap)
            f = i / float(n)
            on, off = ((C["red"], "#3a1f1f") if f > 0.85 else
                       (C["amber"], "#3a3018") if f > 0.65 else (C["green"], "#173226"))
            col = on if i < lit else off
            if i == pki and i >= lit:
                col = C["text"]
            self.create_rectangle(x0, pad, x0 + w, self._ch - pad, fill=col, outline="")


class ScheduleStrip(tk.Canvas):
    HEIGHT = 62

    def __init__(self, parent):
        super().__init__(parent, height=self.HEIGHT, bg=C["surface"],
                         highlightthickness=1, highlightbackground=C["border"], bd=0)
        self._pos = 0.0
        self.bind("<Configure>", lambda e: self.redraw())
        self.redraw()

    def update_position(self):
        d = datetime.now(timezone.utc)
        self._pos = (d.minute % 30) * 60 + d.second + d.microsecond / 1e6
        self.redraw()

    def redraw(self):
        self.delete("all")
        w = self.winfo_width()
        if w < 60:
            w = 900
        pad, top, bot = 10, 22, 42
        span, total = w - 2 * pad, 1800
        for a, b, bs, m in RWMSchedule.TABLE:
            t0, t1 = a * 60, b * 60 + bs
            x0, x1 = pad + span * t0 / total, pad + span * t1 / total
            act = t0 <= self._pos < t1
            self.create_rectangle(x0, top, x1, bot, fill=SEG_COLORS[m],
                                  outline=C["text"] if act else C["border"],
                                  width=2 if act else 1)
            if x1 - x0 > 44:
                self.create_text((x0 + x1) / 2, (top + bot) / 2, text=SEG_LABELS[m],
                                 fill=C["text"] if act else "#b9c6d4",
                                 font=("Segoe UI", 8, "bold" if act else "normal"))
        # зоны пропуска меток 56–59 при minute % 5 == 4
        for mm in range(30):
            if RWMSchedule.seconds_skipped(mm):
                x0 = pad + span * (mm * 60 + 56) / total
                x1 = pad + span * (mm * 60 + 60) / total
                self.create_rectangle(x0, top, x1, bot, fill="#2a2a35",
                                      outline="", stipple="gray50")
        for mm in range(0, 31, 5):
            x = pad + span * mm * 60 / total
            self.create_line(x, bot, x, bot + 4, fill=C["dim"])
            self.create_text(x, bot + 12, text=str(mm), fill=C["dim"], font=("Consolas", 8))
        x = pad + span * self._pos / total
        self.create_line(x, top - 9, x, bot + 4, fill=C["amber"], width=2)
        self.create_polygon(x - 5, top - 11, x + 5, top - 11, x, top - 4,
                            fill=C["amber"], outline="")
        self.create_text(pad, 10, anchor="w",
                         text="ЦИКЛОГРАММА RWM · ITU-R TF.2487-0 · 30 МИН",
                         fill=C["muted"], font=("Segoe UI", 8, "bold"))
        self.create_text(w - pad, 10, anchor="e",
                         text=f"{int(self._pos//60):02d}:{int(self._pos%60):02d}",
                         fill=C["amber"], font=("Consolas", 9, "bold"))


# ------------------------------------------------------------
# GUI
# ------------------------------------------------------------
class TimeSyncApp:
    RWM_FREQS = [("4996", 4996000), ("9996", 9996000), ("14996", 14996000)]
    WF_MS = 150
    WF_MAX = 3000
    WF_BINS = 160
    FFT = 1024
    SRC_LABELS = {"Audio": "① Аудио (PyAudio)", "SDR": "② RTL-SDR · USB",
                  "Auto": "③ Авто (лучший SNR)"}
    LBL2MODE = {v: k for k, v in SRC_LABELS.items()}

    def __init__(self, root):
        self.root = root
        self.root.title("RWM · ITU-R TF.2487-0 · синхронизация времени и NTP-сервер")
        self.root.configure(bg=C["bg"])
        try:
            self.root.state("zoomed")
        except Exception:
            pass
        try:
            self.root.minsize(1280, 820)
        except Exception:
            pass
        self._style()
        self.log_q = queue.Queue()
        self.running = True
        self.cat_mode = CONFIG.get("cat_mode", "Direct CAT")
        self.rx_freq = int(CONFIG.get("rx_freq", 9996000))
        self.direct_cat = DirectCAT()
        self.sniffer = None
        self.sources = SourceManager(log_func=self._log)
        self.monitor = AudioMonitor(self.sources, log_func=self._log)
        self.wf_data = None
        self.wf_lines = 200
        self._sdr_restart = 0.0
        self._wf_nodata = False
        self._fbtn = {}
        self.geo = None
        self._vars()
        self._ui()
        self.decoder = RWMDecoder(log_func=self._log)
        self.sync = TimeSyncController(self.decoder, self._log,
                                       CONFIG.get("use_ntp_fallback", True),
                                       CONFIG.get("ntp_servers"))
        self.ntp = NTPServer(log_func=self._log, status_func=self.sync.server_status)
        self._settings()
        self.monitor.start()
        self.ntp.start()
        self._calc_geo(apply_now=True, quiet=True)
        self._apply_sources(startup=True)
        self._apply_cat(startup=True)
        self._mon_btn()
        self._paint_all_freq()
        self._usb_label()
        for ms, fn in ((100, self._t_log), (120, self._t_vu), (self.WF_MS, self._t_wf),
                       (400, self._t_status), (200, self._t_decode),
                       (1000, self._t_sched), (4000, self._t_diag), (5000, self._watch)):
            self.root.after(ms, fn)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._log("═" * 88)
        self._log("RWM · параметры по Rep. ITU-R TF.2487-0 п. 8.2 · причинный тракт · NTP")
        self._log("A1X 1 Гц: секундная 100 мс, минутная 500 мс")
        self._log("A1N 10 Гц: 0.1 с — 20 мс, начало секунды — 40 мс, начало минуты — 500 мс")
        self._log("Опорная точка — начало переднего фронта. Пропуски сек 56–59 при minute%5==4")
        if self.geo:
            self._log(f"Локатор {CONFIG['locator']} → {self.geo[0]:.4f}°N {self.geo[1]:.4f}°E · "
                      f"до RWM {self.dist_km:.0f} км · наземная {self.gnd_ms:.2f} мс · "
                      f"ионо {self.iono_ms:.2f} мс")
        self._log(f"Тракт: полоса {CONFIG['pulse_band_hz']:.0f} Гц · "
                  f"сглаж {CONFIG['pulse_smooth_ms']:.1f} мс · "
                  f"задержка КИХ {self.decoder.fixed_delay_s*1000:.2f} мс · "
                  f"путь {self.decoder.path_ms:.2f} мс")
        self._log("═" * 88)

    def _style(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure(".", background=C["surface"], foreground=C["text"],
                     fieldbackground=C["surface2"], bordercolor=C["border"],
                     lightcolor=C["surface2"], darkcolor=C["surface"])
        st.configure("TFrame", background=C["surface"])
        st.configure("TLabel", background=C["surface"], foreground=C["text"])
        st.configure("TNotebook", background=C["bg"], bordercolor=C["border"])
        st.configure("TNotebook.Tab", background=C["surface2"], foreground=C["muted"],
                     padding=(12, 7), font=("Segoe UI", 9))
        st.map("TNotebook.Tab", background=[("selected", C["raise"])],
               foreground=[("selected", C["text"])])
        st.configure("TCombobox", fieldbackground=C["surface2"], background=C["raise"],
                     foreground=C["text"], arrowcolor=C["blue"],
                     bordercolor=C["border"], padding=4)
        st.map("TCombobox", fieldbackground=[("readonly", C["surface2"])],
               foreground=[("readonly", C["text"])])
        for k, v in (("*TCombobox*Listbox.background", C["surface2"]),
                     ("*TCombobox*Listbox.foreground", C["text"]),
                     ("*TCombobox*Listbox.selectBackground", C["blue"]),
                     ("*TCombobox*Listbox.selectForeground", "#04121f")):
            self.root.option_add(k, v)
        st.configure("TCheckbutton", background=C["surface"], foreground=C["text"])
        st.map("TCheckbutton", background=[("active", C["surface"])])
        st.configure("Vertical.TScrollbar", background=C["raise"], troughcolor=C["bg"],
                     bordercolor=C["border"], arrowcolor=C["muted"], width=12)
        st.map("Vertical.TScrollbar", background=[("active", C["blue"])])

    def _btn(self, p, text, cmd, kind="normal", small=False):
        pal = {"normal": (C["raise"], C["text"], C["border"], C["blue"]),
               "green": ("#16452c", "#c6ffdf", "#2b7a4d", C["green"]),
               "red": ("#4a1f1f", "#ffd2ce", "#8a3a34", C["red"]),
               "amber": ("#4a3714", "#ffe9b8", "#8a6a24", C["amber"]),
               "blue": ("#123350", "#cfe6ff", "#2a6496", C["blue"]),
               "violet": ("#2f2050", "#e2d3ff", "#5b3f96", C["violet"]),
               "cyan": ("#0f3b3a", "#c9fff8", "#237a72", C["cyan"]),
               "ghost": (C["surface2"], C["muted"], C["border"], C["blue"])}
        bg, fg, bd, hi = pal.get(kind, pal["normal"])
        b = tk.Button(p, text=text, command=cmd, bg=bg, fg=fg,
                      activebackground=hi, activeforeground="#06121d",
                      relief="flat", bd=0, highlightthickness=1,
                      highlightbackground=bd, highlightcolor=hi,
                      font=("Segoe UI", 9 if small else 10, "normal" if small else "bold"),
                      padx=10 if small else 14, pady=5 if small else 7, cursor="hand2")
        b._pal = (bg, fg, bd, hi)
        b.bind("<Enter>", lambda e: b.configure(bg=b._pal[3], fg="#06121d")
               if b["state"] != "disabled" else None)
        b.bind("<Leave>", lambda e: b.configure(bg=b._pal[0], fg=b._pal[1]))
        return b

    def _lbl(self, p, t, muted=False, bold=False, size=9):
        return tk.Label(p, text=t, bg=C["surface"],
                        fg=C["muted"] if muted else C["text"],
                        font=("Segoe UI", size, "bold" if bold else "normal"))

    def _ent(self, p, var, w):
        return tk.Entry(p, textvariable=var, width=w, bg=C["surface2"], fg=C["text"],
                        insertbackground=C["blue"], relief="flat", highlightthickness=1,
                        highlightbackground=C["border"], highlightcolor=C["blue"],
                        font=("Consolas", 10), justify="center")

    def _vars(self):
        g = CONFIG.get
        self.cat_mode_var = tk.StringVar(value=g("cat_mode", "Direct CAT"))
        self.cat_com_var = tk.StringVar(value=g("cat_com_port", ""))
        self.cat_baud_var = tk.StringVar(value=str(g("cat_baudrate", 19200)))
        self.cat_off_var = tk.StringVar(value=str(g("cat_freq_offset", -2000)))
        self.audio_var = tk.StringVar(value=g("audio_input", ""))
        self.agc_var = tk.BooleanVar(value=g("agc_enabled", True))
        self.freq_var = tk.StringVar(value=str(self.rx_freq))
        self.loc_var = tk.StringVar(value=g("locator", "LN23HA"))
        self.iono_h_var = tk.StringVar(value=str(g("iono_height_km", 300)))
        self.iono_n_var = tk.StringVar(value=str(g("iono_hops", 1)))
        self.path_mode_var = tk.StringVar(value=g("path_mode", "Iono"))
        self.path_var = tk.StringVar(value=str(g("path_delay_ms", 0)))
        self.trim_var = tk.StringVar(value=str(g("timing_trim_ms", 0)))
        self.tone_var = tk.StringVar(value=str(g("pulse_center_hz", 2000)))
        self.band_var = tk.StringVar(value=str(g("pulse_band_hz", 600)))
        self.smooth_var = tk.StringVar(value=str(g("pulse_smooth_ms", 1.0)))
        self.lo_var = tk.StringVar(value=str(g("edge_lo_frac", 0.2)))
        self.hi_var = tk.StringVar(value=str(g("edge_hi_frac", 0.8)))
        self.gc_var = tk.StringVar(value=str(g("gate_coarse_ms", 450)))
        self.gf_var = tk.StringVar(value=str(g("gate_fine_ms", 80)))
        self.dtol_var = tk.StringVar(value=str(g("dur_tolerance", 0.45)))
        self.thr_var = tk.StringVar(value=str(g("sync_threshold_ms", 3)))
        self.dut1_var = tk.BooleanVar(value=g("decode_dut1", True))
        m = g("source_mode", "SDR")
        self.src_var = tk.StringVar(value=self.SRC_LABELS.get(m, self.SRC_LABELS["SDR"]))
        self.rate_var = tk.StringVar(value=str(g("sdr_sample_rate", 240000)))
        self.bw_var = tk.StringVar(value=str(g("sdr_if_bw_hz", 1200)))
        self.span_var = tk.StringVar(value=str(g("sdr_lock_span_hz", 1500)))
        self.fft_var = tk.StringVar(value=str(g("sdr_carrier_fft", 32768)))
        self.ar_var = tk.BooleanVar(value=g("sdr_autorestart", True))
        self.vol_var = tk.DoubleVar(value=float(g("monitor_volume", 0.75)) * 100)
        self.mout_var = tk.StringVar(value=g("monitor_output", ""))
        self.ntp_var = tk.BooleanVar(value=g("use_ntp_fallback", True))
        self.ne_var = tk.BooleanVar(value=g("ntp_server_enabled", False))
        self.np_var = tk.StringVar(value=str(g("ntp_server_port", 123)))
        self.nb_var = tk.StringVar(value=g("ntp_server_bind", "0.0.0.0"))
        self.nr_var = tk.BooleanVar(value=g("ntp_server_require_sync", False))
        self.na_var = tk.StringVar(value=g("ntp_server_allow", ""))

    def _ui(self):
        head = tk.Frame(self.root, bg=C["bg"])
        head.pack(fill="x", padx=10, pady=(10, 6))
        lf = tk.Frame(head, bg=C["bg"])
        lf.pack(side="left")
        tk.Label(lf, text="◈ RWM", bg=C["bg"], fg=C["blue"],
                 font=("Segoe UI", 17, "bold")).pack(side="left")
        tk.Label(lf, text="  ITU-R TF.2487-0 · ВНИИФТРИ · синхронизация и NTP",
                 bg=C["bg"], fg=C["muted"], font=("Segoe UI", 10)).pack(side="left", pady=(6, 0))
        rf = tk.Frame(head, bg=C["bg"])
        rf.pack(side="right")
        self.utc_lbl = tk.Label(rf, text="--:--:-- UTC", bg=C["bg"], fg=C["cyan"],
                                font=("Consolas", 13, "bold"))
        self.utc_lbl.pack(side="right", padx=(10, 0))
        self.mode_badge = tk.Label(rf, text="—", bg=C["surface2"], fg=C["text"],
                                   font=("Segoe UI", 9, "bold"), padx=12, pady=5)
        self.mode_badge.pack(side="right")
        self._btn(rf, "⚙ Настройки", self.open_settings, "ghost", True
                  ).pack(side="right", padx=(0, 10))

        ctrl = tk.Frame(self.root, bg=C["surface"], highlightthickness=1,
                        highlightbackground=C["border"])
        ctrl.pack(fill="x", padx=10, pady=(0, 6))
        inner = tk.Frame(ctrl, bg=C["surface"])
        inner.pack(fill="x", padx=12, pady=10)

        r1 = tk.Frame(inner, bg=C["surface"])
        r1.pack(fill="x")
        g1 = tk.Frame(r1, bg=C["surface"])
        g1.pack(side="left")
        self._btn(g1, "▶  Старт", self.start_sync, "green").pack(side="left", padx=(0, 4))
        self._btn(g1, "■  Стоп", self.stop_sync, "red").pack(side="left", padx=4)
        self._btn(g1, "Тест сигнала", self.test_signal, "amber").pack(side="left", padx=4)
        tk.Frame(r1, bg=C["border"], width=1, height=30).pack(side="left", padx=14, fill="y")
        g2 = tk.Frame(r1, bg=C["surface"])
        g2.pack(side="left")
        self.mon_button = self._btn(g2, "🔇  Звук на ПК", self._toggle_mon, "ghost")
        self.mon_button.pack(side="left", padx=(0, 10))
        self._lbl(g2, "громкость", muted=True).pack(side="left", padx=(0, 6))
        tk.Scale(g2, from_=0, to=100, orient="horizontal", length=110, variable=self.vol_var,
                 showvalue=False, command=self._on_vol, bg=C["surface"],
                 troughcolor=C["surface2"], highlightthickness=0,
                 activebackground=C["blue"], sliderrelief="flat", sliderlength=16,
                 width=10, bd=0, fg=C["text"], cursor="hand2").pack(side="left")
        self.vol_lbl = tk.Label(g2, text=f"{int(self.vol_var.get())}%", bg=C["surface"],
                                fg=C["cyan"], width=5, font=("Consolas", 10, "bold"))
        self.vol_lbl.pack(side="left", padx=(6, 0))
        tk.Frame(r1, bg=C["border"], width=1, height=30).pack(side="left", padx=14, fill="y")
        self.ntp_button = self._btn(r1, "🕓  NTP-сервер", self._toggle_ntp, "ghost")
        self.ntp_button.pack(side="left")
        self.ntp_badge = tk.Label(r1, text="выкл", bg=C["surface2"], fg=C["dim"],
                                  font=("Consolas", 9), padx=10, pady=5)
        self.ntp_badge.pack(side="left", padx=(8, 0))
        g3 = tk.Frame(r1, bg=C["surface"])
        g3.pack(side="right")
        self.corr_lbl = tk.Label(g3, text="коррекций 0", bg=C["surface2"], fg=C["muted"],
                                 font=("Consolas", 9), padx=10, pady=5)
        self.corr_lbl.pack(side="right", padx=(8, 0))
        self.off_lbl = tk.Label(g3, text="Δ —", bg=C["surface2"], fg=C["muted"],
                                font=("Consolas", 10, "bold"), padx=12, pady=5)
        self.off_lbl.pack(side="right")

        r2 = tk.Frame(inner, bg=C["surface"])
        r2.pack(fill="x", pady=(10, 0))
        fq = tk.Frame(r2, bg=C["surface"])
        fq.pack(side="left")
        self._lbl(fq, "Несущая RWM", muted=True).pack(side="left", padx=(0, 6))
        e = tk.Entry(fq, textvariable=self.freq_var, width=11, bg=C["surface2"], fg=C["text"],
                     insertbackground=C["blue"], relief="flat", highlightthickness=1,
                     highlightbackground=C["border"], highlightcolor=C["blue"],
                     font=("Consolas", 12, "bold"), justify="center")
        e.pack(side="left", ipady=5)
        e.bind("<Return>", lambda ev: self.set_freq())
        self._lbl(fq, "Гц", muted=True).pack(side="left", padx=(5, 8))
        self._btn(fq, "Установить", self.set_freq, "blue", True).pack(side="left")
        tk.Frame(r2, bg=C["border"], width=1, height=28).pack(side="left", padx=14, fill="y")
        pre = tk.Frame(r2, bg=C["surface2"], highlightthickness=1,
                       highlightbackground=C["border"])
        pre.pack(side="left")
        tk.Label(pre, text=" RWM ", bg=C["surface2"], fg=C["dim"],
                 font=("Segoe UI", 8, "bold")).pack(side="left", padx=(4, 2))
        for lab, fhz in self.RWM_FREQS:
            b = tk.Button(pre, text=f"{lab} кГц",
                          command=lambda f=fhz: self.set_rwm(f),
                          bg=C["surface2"], fg=C["text"], activebackground=C["blue"],
                          activeforeground="#06121d", relief="flat", bd=0,
                          highlightthickness=0, font=("Consolas", 10, "bold"),
                          padx=12, pady=6, cursor="hand2")
            b.pack(side="left", padx=1, pady=1)
            b.bind("<Enter>", lambda ev, bb=b, f=fhz: bb.configure(
                bg="#2a8a55" if abs(f - self.rx_freq) < 1 else C["blue"], fg="#06121d"))
            b.bind("<Leave>", lambda ev, bb=b, f=fhz: self._paint_freq(bb, f))
            self._fbtn[fhz] = b
        stf = tk.Frame(r2, bg=C["surface"])
        stf.pack(side="left", padx=(10, 0))
        for t, d in (("−100", -100), ("−10", -10), ("+10", 10), ("+100", 100)):
            self._btn(stf, t, lambda dd=d: self.set_freq(self.rx_freq + dd),
                      "ghost", True).pack(side="left", padx=1)
        self.geo_lbl = tk.Label(r2, text="", bg=C["surface"], fg=C["violet"],
                                font=("Consolas", 9))
        self.geo_lbl.pack(side="right")

        r3 = tk.Frame(inner, bg=C["surface"])
        r3.pack(fill="x", pady=(10, 0))
        self._lbl(r3, "Источник", muted=True).pack(side="left", padx=(0, 6))
        cb = ttk.Combobox(r3, textvariable=self.src_var, width=22, state="readonly",
                          values=list(self.SRC_LABELS.values()))
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda ev: self._apply_sources())
        self._btn(r3, "Применить", self._apply_sources, "violet", True).pack(side="left", padx=8)
        self.src_info = tk.Label(r3, text="—", bg=C["surface"], fg=C["muted"],
                                 font=("Consolas", 9))
        self.src_info.pack(side="left", padx=8)
        self.mon_state = tk.Label(r3, text="монитор выключен", bg=C["surface"],
                                  fg=C["dim"], font=("Segoe UI", 9))
        self.mon_state.pack(side="right")
        self.status_lbl = tk.Label(r3, text="Остановлен", bg=C["surface"], fg=C["muted"],
                                   font=("Segoe UI", 10, "bold"))
        self.status_lbl.pack(side="right", padx=14)

        self.sched = ScheduleStrip(self.root)
        self.sched.pack(fill="x", padx=10, pady=(0, 6))

        pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg=C["bg"],
                              sashwidth=6, sashrelief="flat", bd=0)
        pane.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.pane = pane
        left = tk.Frame(pane, bg=C["surface"], highlightthickness=1,
                        highlightbackground=C["border"])
        pane.add(left, width=760, minsize=480, stretch="always")
        lh = tk.Frame(left, bg=C["surface"])
        lh.pack(fill="x", padx=10, pady=(8, 4))
        tk.Label(lh, text="ЖУРНАЛ", bg=C["surface"], fg=C["blue"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self.pulse_lbl = tk.Label(lh, text="меток 0 · 0%", bg=C["surface"], fg=C["muted"],
                                  font=("Consolas", 9))
        self.pulse_lbl.pack(side="left", padx=14)
        self.dut1_lbl = tk.Label(lh, text="DUT1 —", bg=C["surface"], fg=C["violet"],
                                 font=("Consolas", 9))
        self.dut1_lbl.pack(side="left", padx=6)
        self._btn(lh, "Очистить", self._clear_log, "ghost", True).pack(side="right")
        lfr = tk.Frame(left, bg=C["surface"])
        lfr.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.log_text = tk.Text(lfr, bg="#070b11", fg="#c8d6e5", relief="flat",
                                font=("Consolas", 9), state="disabled", wrap="none",
                                insertbackground=C["blue"], padx=10, pady=8,
                                highlightthickness=1, highlightbackground=C["border"])
        sb = ttk.Scrollbar(lfr, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)
        for t, c in (("ok", C["green"]), ("warn", C["amber"]), ("err", C["red"]),
                     ("mark", C["magenta"]), ("info", C["blue"]),
                     ("ntp", C["cyan"]), ("geo", C["violet"]), ("dim", C["dim"])):
            self.log_text.tag_configure(t, foreground=c)
        self.log_text.tag_configure("mark", font=("Consolas", 9, "bold"))

        right = tk.Frame(pane, bg=C["bg"])
        pane.add(right, width=470, minsize=380, stretch="always")
        mf = tk.Frame(right, bg=C["surface"], highlightthickness=1,
                      highlightbackground=C["border"])
        mf.pack(fill="x", pady=(0, 6))
        mi = tk.Frame(mf, bg=C["surface"])
        mi.pack(fill="x", padx=12, pady=10)
        for title, attr, lab in (("ВХОД · активный источник", "vu_in", "in_val"),
                                 ("ВЫХОД · монитор 48 кГц", "vu_out", "out_val")):
            row = tk.Frame(mi, bg=C["surface"])
            row.pack(fill="x", pady=3)
            tk.Label(row, text=title, bg=C["surface"], fg=C["muted"],
                     font=("Segoe UI", 8, "bold"), width=26, anchor="w").pack(side="left")
            bar = LevelBar(row, width=230, height=16)
            bar.pack(side="left", padx=6)
            v = tk.Label(row, text="0%", bg=C["surface"], fg=C["text"],
                         font=("Consolas", 9, "bold"), width=5, anchor="e")
            v.pack(side="left")
            setattr(self, attr, bar)
            setattr(self, lab, v)
        wf = tk.Frame(right, bg=C["surface"], highlightthickness=1,
                      highlightbackground=C["border"])
        wf.pack(fill="both", expand=True)
        wh = tk.Frame(wf, bg=C["surface"])
        wh.pack(fill="x", padx=12, pady=(8, 2))
        tk.Label(wh, text="ВОДОПАД · 0…3000 Гц", bg=C["surface"], fg=C["blue"],
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        self.wf_state = tk.Label(wh, text="—", bg=C["surface"], fg=C["dim"],
                                 font=("Consolas", 8))
        self.wf_state.pack(side="right")
        if MPL and np is not None:
            self.fig = Figure(figsize=(4, 3.4), dpi=90, facecolor=C["surface"])
            self.ax = self.fig.add_subplot(111)
            self.ax.set_facecolor("#05080d")
            self.fig.subplots_adjust(left=0.11, right=0.98, top=0.97, bottom=0.13)
            self.wf_data = np.full((self.wf_lines, self.WF_BINS), -90.0)
            self.wf_img = self.ax.imshow(
                self.wf_data, aspect="auto", origin="lower", cmap="inferno",
                extent=[0, self.WF_MAX, 0, self.wf_lines * self.WF_MS / 1000.0],
                vmin=-90, vmax=-20)
            self.ax.set_xlabel("частота, Гц", color=C["muted"], fontsize=8)
            self.ax.set_ylabel("время, с", color=C["muted"], fontsize=8)
            self.ax.tick_params(colors=C["dim"], labelsize=7)
            for s in self.ax.spines.values():
                s.set_color(C["border"])
            self.tline = self.ax.axvline(float(CONFIG.get("pulse_center_hz", 2000)),
                                         color=C["cyan"], lw=1.0, ls="--", alpha=0.75)
            self.canvas = FigureCanvasTkAgg(self.fig, master=wf)
            w = self.canvas.get_tk_widget()
            w.configure(bg=C["surface"], highlightthickness=0)
            w.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        else:
            self.fig = self.ax = self.wf_img = self.canvas = self.tline = None
            tk.Label(wf, text="matplotlib недоступен", bg=C["surface"],
                     fg=C["red"]).pack(fill="both", expand=True)
        self.root.after(200, self._sash)

    def _sash(self):
        try:
            self.root.update_idletasks()
            w = self.pane.winfo_width()
            if w > 200:
                self.pane.sash_place(0, int(w * 0.60), 1)
        except Exception:
            pass

    def _paint_freq(self, b, f):
        b.configure(bg="#16452c" if abs(f - self.rx_freq) < 1 else C["surface2"],
                    fg="#c6ffdf" if abs(f - self.rx_freq) < 1 else C["text"])

    def _paint_all_freq(self):
        for f, b in self._fbtn.items():
            self._paint_freq(b, f)

    def _usb_label(self):
        t = float(CONFIG.get("pulse_center_hz", 2000))
        self.geo_lbl.configure(
            text=f"USB {int(self.rx_freq-t)} Гц → {t:.0f} Гц · "
                 f"{CONFIG.get('locator','')} · путь {self.decoder.path_ms:.2f} мс"
            if getattr(self, "decoder", None) else "")

    # ---------- геолокация ----------
    def _calc_geo(self, apply_now=False, quiet=False):
        loc = self.loc_var.get().strip().upper()
        ll = maidenhead_to_latlon(loc)
        if ll is None:
            if not quiet:
                messagebox.showerror("Ошибка", "Неверный локатор Maidenhead (напр. LN23HA)")
            return False
        self.geo = ll
        self.dist_km = haversine_km(ll[0], ll[1], RWM_LAT, RWM_LON)
        try:
            h = float(self.iono_h_var.get())
        except Exception:
            h = 300.0
        try:
            hops = max(1, int(float(self.iono_n_var.get())))
        except Exception:
            hops = 1
        self.gnd_ms, self.iono_ms = path_delays_ms(self.dist_km, h, hops)
        CONFIG["locator"] = loc
        CONFIG["iono_height_km"] = h
        CONFIG["iono_hops"] = hops
        CONFIG["path_mode"] = self.path_mode_var.get()
        mode = self.path_mode_var.get()
        if mode == "Ground":
            val = self.gnd_ms
        elif mode == "Iono":
            val = self.iono_ms
        else:
            try:
                val = float(self.path_var.get())
            except Exception:
                val = self.iono_ms
        if mode != "Manual":
            self.path_var.set(f"{val:.3f}")
        CONFIG["path_delay_ms"] = round(val, 4)
        save_config(CONFIG)
        if getattr(self, "decoder", None):
            self.decoder.path_ms = val
        if hasattr(self, "geo_status"):
            self.geo_status.config(
                text=f"{loc} → {ll[0]:.4f}°N {ll[1]:.4f}°E · "
                     f"RWM 56.7333°N 37.6333°E\n"
                     f"расстояние {self.dist_km:.1f} км · наземная {self.gnd_ms:.3f} мс · "
                     f"ионо {hops}-hop h={h:.0f} км → {self.iono_ms:.3f} мс\n"
                     f"применено ({mode}): {val:.3f} мс", fg=C["green"])
        self._usb_label()
        if not quiet:
            self._log(f"[ГЕО] {loc} → {ll[0]:.4f}°N {ll[1]:.4f}°E · "
                      f"до RWM {self.dist_km:.1f} км")
            self._log(f"[ГЕО] наземная {self.gnd_ms:.3f} мс · "
                      f"ионо {hops}-hop h={h:.0f} км {self.iono_ms:.3f} мс · "
                      f"применено {val:.3f} мс ({mode})")
        return True

    # ---------- настройки ----------
    def _settings(self):
        self.sw = tk.Toplevel(self.root)
        self.sw.title("Настройки")
        self.sw.configure(bg=C["bg"])
        self.sw.geometry("880x700")
        self.sw.withdraw()
        self.sw.protocol("WM_DELETE_WINDOW", lambda: self.sw.withdraw())
        nb = ttk.Notebook(self.sw)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        # Гео
        t0 = tk.Frame(nb, bg=C["surface"])
        nb.add(t0, text="🌍 Локатор / путь")
        f0 = tk.Frame(t0, bg=C["surface"])
        f0.pack(fill="both", expand=True, padx=16, pady=16)
        self._lbl(f0, "Локатор Maidenhead", muted=True).grid(row=0, column=0, sticky="e", pady=6)
        self._ent(f0, self.loc_var, 12).grid(row=0, column=1, sticky="w", padx=8)
        self._lbl(f0, "напр. LN23HA", muted=True).grid(row=0, column=2, sticky="w")
        self._lbl(f0, "Высота слоя F2, км", muted=True).grid(row=1, column=0, sticky="e", pady=6)
        self._ent(f0, self.iono_h_var, 12).grid(row=1, column=1, sticky="w", padx=8)
        self._lbl(f0, "обычно 250–350", muted=True).grid(row=1, column=2, sticky="w")
        self._lbl(f0, "Число скачков", muted=True).grid(row=2, column=0, sticky="e", pady=6)
        ttk.Combobox(f0, textvariable=self.iono_n_var, width=10, state="readonly",
                     values=["1", "2", "3"]).grid(row=2, column=1, sticky="w", padx=8)
        self._lbl(f0, "Модель задержки", muted=True).grid(row=3, column=0, sticky="e", pady=6)
        ttk.Combobox(f0, textvariable=self.path_mode_var, width=12, state="readonly",
                     values=["Ground", "Iono", "Manual"]).grid(row=3, column=1, sticky="w", padx=8)
        self._lbl(f0, "Задержка пути, мс", muted=True).grid(row=4, column=0, sticky="e", pady=6)
        self._ent(f0, self.path_var, 12).grid(row=4, column=1, sticky="w", padx=8)
        self._lbl(f0, "Manual — вводится вручную", muted=True).grid(row=4, column=2, sticky="w")
        self._lbl(f0, "Подстройка времени, мс", muted=True).grid(row=5, column=0, sticky="e", pady=6)
        self._ent(f0, self.trim_var, 12).grid(row=5, column=1, sticky="w", padx=8)
        bf0 = tk.Frame(f0, bg=C["surface"])
        bf0.grid(row=6, column=0, columnspan=3, sticky="w", pady=10)
        self._btn(bf0, "Рассчитать и применить", lambda: self._calc_geo(apply_now=True),
                  "violet").pack(side="left", padx=(0, 8))
        self._btn(bf0, "Применить подстройку", self._apply_trim, "blue").pack(side="left")
        self.geo_status = tk.Label(f0, text="—", bg=C["surface2"], fg=C["muted"],
                                   font=("Consolas", 9), padx=10, pady=10,
                                   justify="left", anchor="w")
        self.geo_status.grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)
        self._lbl(f0, "RWM: 56°44′N 37°38′E · 10 кВт · 4996/9996/14996 кГц\n"
                      "Погрешность несущей 5·10⁻¹² (усреднение за сутки).\n"
                      "Наземная волна = d/c. Ионо: n-hop, плоская модель 2·√((d/2n)²+h²).\n"
                      "Остаточное смещение относительно NTP убрать «Подстройкой».",
                  muted=True).grid(row=8, column=0, columnspan=3, sticky="w")

        # Аудио
        t1 = tk.Frame(nb, bg=C["surface"])
        nb.add(t1, text="① Аудио / Монитор")
        f = tk.Frame(t1, bg=C["surface"])
        f.pack(fill="both", expand=True, padx=16, pady=16)
        self._lbl(f, "Входное устройство", muted=True).grid(row=0, column=0, sticky="e", pady=8)
        self.audio_cb = ttk.Combobox(f, textvariable=self.audio_var, width=44, state="readonly")
        self.audio_cb.grid(row=0, column=1, padx=8)
        self._btn(f, "Обновить", self._refresh_audio, "ghost", True).grid(row=0, column=2)
        ttk.Checkbutton(f, text="AGC входа (пиковый)", variable=self.agc_var,
                        command=self._save_agc).grid(row=1, column=0, columnspan=3,
                                                     sticky="w", pady=8)
        tk.Frame(f, bg=C["border"], height=1).grid(row=2, column=0, columnspan=3,
                                                   sticky="ew", pady=12)
        self._lbl(f, "Выход монитора", muted=True).grid(row=3, column=0, sticky="e", pady=8)
        self.mout_cb = ttk.Combobox(f, textvariable=self.mout_var, width=44, state="readonly")
        self.mout_cb.grid(row=3, column=1, padx=8)
        self._btn(f, "Применить", self._apply_mout, "blue", True).grid(row=3, column=2)
        self._lbl(f, "Пусто = по умолчанию. Выход 48000 Гц, sinc ×4.",
                  muted=True).grid(row=4, column=0, columnspan=3, sticky="w", pady=8)

        # SDR
        t2 = tk.Frame(nb, bg=C["surface"])
        nb.add(t2, text="② RTL-SDR · USB")
        f2 = tk.Frame(t2, bg=C["surface"])
        f2.pack(fill="both", expand=True, padx=16, pady=16)
        self._lbl(f2, "Устройство", muted=True).grid(row=0, column=0, sticky="e", pady=5)
        self.sdr_cb = ttk.Combobox(f2, width=42, state="readonly")
        self.sdr_cb.grid(row=0, column=1, padx=8, sticky="w")
        self._btn(f2, "Обновить", self._refresh_sdr, "ghost", True).grid(row=0, column=2)
        self._lbl(f2, "Режим", muted=True).grid(row=1, column=0, sticky="e", pady=5)
        self._lbl(f2, "Direct Q · USB", bold=True).grid(row=1, column=1, sticky="w", padx=8)
        for i, (lab, var, vals, note) in enumerate([
                ("Частота дискретизации", self.rate_var, [str(r) for r in SDRInput.RATES],
                 "240000 / 20 = ровно 12000"),
                ("БПФ несущей (точек)", self.fft_var,
                 ["8192", "16384", "32768", "65536"], "32768 → 7.3 Гц/бин")]):
            self._lbl(f2, lab, muted=True).grid(row=2 + i, column=0, sticky="e", pady=5)
            ttk.Combobox(f2, textvariable=var, width=12, state="readonly",
                         values=vals).grid(row=2 + i, column=1, sticky="w", padx=8)
            self._lbl(f2, note, muted=True).grid(row=2 + i, column=2, sticky="w")
        self._lbl(f2, "Полоса тракта ±Гц", muted=True).grid(row=4, column=0, sticky="e", pady=5)
        self._ent(f2, self.bw_var, 12).grid(row=4, column=1, sticky="w", padx=8)
        self._lbl(f2, "Поиск несущей ±Гц", muted=True).grid(row=5, column=0, sticky="e", pady=5)
        self._ent(f2, self.span_var, 12).grid(row=5, column=1, sticky="w", padx=8)
        ttk.Checkbutton(f2, text="Авто-перезапуск при обрыве", variable=self.ar_var,
                        command=self._save_sdr).grid(row=6, column=0, columnspan=3,
                                                     sticky="w", pady=8)
        bf = tk.Frame(f2, bg=C["surface"])
        bf.grid(row=7, column=0, columnspan=3, sticky="w", pady=6)
        self._btn(bf, "Применить / перезапустить", self._apply_sources, "violet"
                  ).pack(side="left", padx=(0, 8))
        self._btn(bf, "Тест USB", self._test_sdr, "amber").pack(side="left", padx=(0, 8))
        self._btn(bf, "Калибровать PPM", self._cal_ppm, "cyan").pack(side="left")
        self.sdr_status = tk.Label(f2, text="Не запущен", bg=C["surface2"], fg=C["muted"],
                                   font=("Consolas", 9), padx=10, pady=8,
                                   justify="left", anchor="w")
        self.sdr_status.grid(row=8, column=0, columnspan=3, sticky="ew", pady=10)

        # Декодер
        t3 = tk.Frame(nb, bg=C["surface"])
        nb.add(t3, text="🎯 Декодер RWM")
        f3 = tk.Frame(t3, bg=C["surface"])
        f3.pack(fill="both", expand=True, padx=16, pady=14)
        rows = [("Тон / центр фильтра, Гц", self.tone_var),
                ("Полоса фильтра, Гц", self.band_var),
                ("Сглаживание, мс", self.smooth_var),
                ("Фронт: нижняя доля", self.lo_var),
                ("Фронт: верхняя доля", self.hi_var),
                ("Окно грубое ±мс", self.gc_var),
                ("Окно точное ±мс", self.gf_var),
                ("Допуск длительности", self.dtol_var),
                ("Порог коррекции, мс", self.thr_var)]
        for i, (lab, var) in enumerate(rows):
            self._lbl(f3, lab, muted=True).grid(row=i, column=0, sticky="e", pady=3)
            self._ent(f3, var, 12).grid(row=i, column=1, sticky="w", padx=8)
        bf3 = tk.Frame(f3, bg=C["surface"])
        bf3.grid(row=len(rows), column=0, columnspan=2, sticky="w", pady=8)
        self._btn(bf3, "Применить", self._apply_dec, "blue").pack(side="left", padx=(0, 8))
        self._btn(bf3, "Измерить длительности 30 с", self._measure, "amber").pack(side="left")
        ttk.Checkbutton(f3, text="Декодировать DUT1 + dUT1", variable=self.dut1_var,
                        command=self._save_dut1).grid(row=len(rows) + 1, column=0,
                                                      columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(f3, text="Резервная синхронизация по NTP (клиент)",
                        variable=self.ntp_var, command=self._save_ntp
                        ).grid(row=len(rows) + 2, column=0, columnspan=2, sticky="w", pady=4)
        self._lbl(f3,
                  "ITU-R TF.2487-0: A1X 100 мс / минутная 500 мс;\n"
                  "A1N 20 мс / начало секунды 40 мс / начало минуты 500 мс.\n"
                  "Класс метки берётся из положения в сетке UTC, длительность — контроль.\n"
                  "Опорная точка — начало переднего фронта: точки 20 % и 80 %\n"
                  "экстраполируются к нулю, задержки КИХ вычитаются.\n"
                  "Сек 56–59 при minute % 5 == 4 пропускаются (сноска к табл. 23).",
                  muted=True).grid(row=len(rows) + 3, column=0, columnspan=2, sticky="w", pady=6)

        # NTP
        t4 = tk.Frame(nb, bg=C["surface"])
        nb.add(t4, text="🕓 NTP-сервер")
        f4 = tk.Frame(t4, bg=C["surface"])
        f4.pack(fill="both", expand=True, padx=16, pady=16)
        ttk.Checkbutton(f4, text="Раздавать время по NTP (UDP)",
                        variable=self.ne_var).grid(row=0, column=0, columnspan=3,
                                                   sticky="w", pady=6)
        self._lbl(f4, "Порт UDP", muted=True).grid(row=1, column=0, sticky="e", pady=5)
        self._ent(f4, self.np_var, 10).grid(row=1, column=1, sticky="w", padx=8)
        self._lbl(f4, "123 стандартный; 1123 если занят w32time",
                  muted=True).grid(row=1, column=2, sticky="w")
        self._lbl(f4, "Слушать адрес", muted=True).grid(row=2, column=0, sticky="e", pady=5)
        ttk.Combobox(f4, textvariable=self.nb_var, width=18, state="readonly",
                     values=local_ipv4_list()).grid(row=2, column=1, sticky="w", padx=8)
        ttk.Checkbutton(f4, text="Отвечать только при синхронизации (иначе LI=3, stratum 16)",
                        variable=self.nr_var).grid(row=3, column=0, columnspan=3,
                                                   sticky="w", pady=6)
        self._lbl(f4, "Разрешённые подсети", muted=True).grid(row=4, column=0, sticky="e", pady=5)
        self._ent(f4, self.na_var, 30).grid(row=4, column=1, columnspan=2, sticky="w", padx=8)
        self._lbl(f4, "Пусто = всем. Пример: 192.168.0.0/16, 10.0.0.0/8",
                  muted=True).grid(row=5, column=0, columnspan=3, sticky="w")
        bf4 = tk.Frame(f4, bg=C["surface"])
        bf4.grid(row=6, column=0, columnspan=3, sticky="w", pady=8)
        self._btn(bf4, "Применить / перезапустить", self._apply_ntp, "cyan"
                  ).pack(side="left", padx=(0, 8))
        self._btn(bf4, "Тест NTP", self._test_ntp, "amber").pack(side="left")
        self.ntp_status = tk.Label(f4, text="Остановлен", bg=C["surface2"], fg=C["muted"],
                                   font=("Consolas", 9), padx=10, pady=8,
                                   justify="left", anchor="w")
        self.ntp_status.grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)
        self._lbl(f4, "Stratum 1/RWM — захват меток; 2/NTP; 3/LOCL — свободный ход.\n"
                      "Windows: net stop w32time, либо порт 1123. Откройте UDP в брандмауэре.",
                  muted=True).grid(row=8, column=0, columnspan=3, sticky="w")

        # CAT
        t5 = tk.Frame(nb, bg=C["surface"])
        nb.add(t5, text="CAT")
        f5 = tk.Frame(t5, bg=C["surface"])
        f5.pack(fill="both", expand=True, padx=16, pady=16)
        self._lbl(f5, "Режим", muted=True).grid(row=0, column=0, sticky="e", pady=5)
        ttk.Combobox(f5, textvariable=self.cat_mode_var, width=15, state="readonly",
                     values=["Direct CAT", "Sniffer"]).grid(row=0, column=1, sticky="w", padx=8)
        self._lbl(f5, "COM-порт", muted=True).grid(row=1, column=0, sticky="e", pady=5)
        self.cat_cb = ttk.Combobox(f5, textvariable=self.cat_com_var, width=14, state="readonly")
        self.cat_cb.grid(row=1, column=1, sticky="w", padx=8)
        self._btn(f5, "Обновить", self._refresh_cat, "ghost", True).grid(row=1, column=2)
        self._lbl(f5, "Скорость", muted=True).grid(row=2, column=0, sticky="e", pady=5)
        ttk.Combobox(f5, textvariable=self.cat_baud_var, width=12, state="readonly",
                     values=[str(x) for x in (9600, 19200, 38400, 57600, 115200)]
                     ).grid(row=2, column=1, sticky="w", padx=8)
        self._lbl(f5, "Смещение CAT, Гц", muted=True).grid(row=3, column=0, sticky="e", pady=5)
        self._ent(f5, self.cat_off_var, 12).grid(row=3, column=1, sticky="w", padx=8)
        self._btn(f5, "Применить CAT", self._apply_cat, "blue"
                  ).grid(row=4, column=1, sticky="w", padx=8, pady=10)
        self.cat_status = tk.Label(f5, text="Не подключено", bg=C["surface2"], fg=C["muted"],
                                   font=("Consolas", 9), padx=10, pady=8, anchor="w")
        self.cat_status.grid(row=5, column=0, columnspan=3, sticky="ew", pady=6)

        self._refresh_audio()
        self._refresh_cat()
        self._refresh_sdr()

    def open_settings(self):
        self.sw.deiconify()
        self.sw.lift()

    # ---------- монитор ----------
    def _toggle_mon(self):
        s = not self.monitor.enabled
        self.monitor.set_enabled(s)
        CONFIG["monitor_enabled"] = s
        save_config(CONFIG)
        self._mon_btn()
        self._log(f"[МОНИТОР] {'ВКЛЮЧЁН' if s else 'ВЫКЛЮЧЕН'} · "
                  f"громкость {int(self.vol_var.get())}%")

    def _on_vol(self, _=None):
        v = float(self.vol_var.get())
        self.monitor.set_volume(v / 100)
        self.vol_lbl.config(text=f"{int(v)}%")
        CONFIG["monitor_volume"] = v / 100
        save_config(CONFIG)

    def _mon_btn(self):
        if self.monitor.enabled:
            self.mon_button.configure(text="🔊  Звук на ПК", bg="#16452c", fg="#c6ffdf")
            self.mon_button._pal = ("#16452c", "#c6ffdf", "#2b7a4d", C["green"])
        else:
            self.mon_button.configure(text="🔇  Звук на ПК", bg=C["surface2"], fg=C["muted"])
            self.mon_button._pal = (C["surface2"], C["muted"], C["border"], C["blue"])

    def _apply_mout(self):
        n = self.mout_var.get().strip()
        CONFIG["monitor_output"] = n
        save_config(CONFIG)
        self.monitor.set_device(n)
        self._log(f"[МОНИТОР] выход: {n or 'по умолчанию'}")

    # ---------- NTP ----------
    def _toggle_ntp(self):
        self.ne_var.set(not self.ntp.enabled)
        self._apply_ntp()

    def _apply_ntp(self):
        try:
            p = int(self.np_var.get())
            if not 1 <= p <= 65535:
                raise ValueError
        except Exception:
            messagebox.showerror("Ошибка", "Порт 1…65535")
            return
        CONFIG.update({"ntp_server_enabled": bool(self.ne_var.get()),
                       "ntp_server_port": p,
                       "ntp_server_bind": self.nb_var.get().strip() or "0.0.0.0",
                       "ntp_server_require_sync": bool(self.nr_var.get()),
                       "ntp_server_allow": self.na_var.get().strip()})
        save_config(CONFIG)
        self.ntp.configure(enabled=CONFIG["ntp_server_enabled"], port=p,
                           bind=CONFIG["ntp_server_bind"],
                           require_sync=CONFIG["ntp_server_require_sync"],
                           allow=CONFIG["ntp_server_allow"])
        self.ntp.restart()
        self._log(f"[NTP-SRV] {'ВКЛЮЧЁН' if CONFIG['ntp_server_enabled'] else 'ВЫКЛЮЧЕН'} "
                  f"· {CONFIG['ntp_server_bind']}:{p}")
        self.root.after(700, self._ntp_status_upd)

    def _ntp_status_upd(self):
        s = self.ntp.get_stats()
        if s["running"]:
            self.ntp_status.config(
                text=f"Слушаю {s['bind']}:{s['port']} UDP · 2^{s['precision']} c\n"
                     f"запросов {s['requests']} · отвечено {s['served']} · "
                     f"отклонено {s['denied']} · последний {s['last_client'] or '—'}",
                fg=C["green"])
        elif s["enabled"] and s["error"]:
            self.ntp_status.config(text=f"Ошибка: {s['error']}", fg=C["red"])
        else:
            self.ntp_status.config(text="Остановлен", fg=C["muted"])

    def _test_ntp(self):
        if not self.ntp.running:
            self._log("[NTP-SRV] сервер не запущен")
            return
        try:
            r = self.ntp.self_test()
            self._log(f"[NTP-SRV] ✅ stratum {r['stratum']} · {r['refid']} · LI {r['li']} · "
                      f"offset {r['offset_ms']:+.3f} мс · delay {r['delay_ms']:.3f} мс")
        except Exception as e:
            self._log(f"[NTP-SRV] ❌ {e}")

    # ---------- источники ----------
    def _refresh_audio(self):
        if pyaudio is None:
            return
        try:
            p = pyaudio.PyAudio()
            ins, outs = [], []
            for i in range(p.get_device_count()):
                inf = p.get_device_info_by_index(i)
                if inf.get("maxInputChannels", 0) > 0:
                    ins.append(inf.get("name", f"Device {i}"))
                if inf.get("maxOutputChannels", 0) > 0:
                    outs.append(inf.get("name", f"Device {i}"))
            p.terminate()
            self.audio_cb["values"] = ins
            if ins and not self.audio_var.get():
                self.audio_var.set(ins[0])
            self.mout_cb["values"] = [""] + outs
        except Exception:
            pass

    def _refresh_sdr(self):
        items = SDRInput.device_list(log=lambda m: self._log(f"[SDR] {m}"))
        self.sdr_cb["values"] = items
        if items:
            if not self.sdr_cb.get():
                self.sdr_cb.current(min(int(CONFIG.get("sdr_index", 0)), len(items) - 1))
        else:
            self.sdr_cb.set("устройства не найдены")

    def _refresh_cat(self):
        ports = []
        if SERIAL_AVAILABLE:
            try:
                ports = sorted([p.device for p in serial.tools.list_ports.comports()])
            except Exception:
                pass
        self.cat_cb["values"] = ports
        if ports and not self.cat_com_var.get():
            self.cat_com_var.set(ports[0])

    def _save_sdr(self):
        try:
            s = self.sdr_cb.get()
            if s.startswith("#"):
                CONFIG["sdr_index"] = int(s.split()[0].lstrip("#"))
        except Exception:
            pass
        for k, v, cast, d in (("sdr_sample_rate", self.rate_var, int, 240000),
                              ("sdr_if_bw_hz", self.bw_var, float, 1200.0),
                              ("sdr_lock_span_hz", self.span_var, float, 1500.0),
                              ("sdr_carrier_fft", self.fft_var, int, 32768)):
            try:
                CONFIG[k] = cast(v.get())
            except Exception:
                CONFIG[k] = d
        CONFIG["sdr_autorestart"] = bool(self.ar_var.get())
        save_config(CONFIG)

    def _apply_sources(self, startup=False):
        m = self.LBL2MODE.get(self.src_var.get(), "SDR")
        self.sources.set_mode(m)
        CONFIG["audio_input"] = self.audio_var.get().strip()
        self._save_sdr()
        if m in ("Audio", "Auto"):
            if not (self.sources.audio and self.sources.audio.running):
                self.sources.start_audio(CONFIG["audio_input"], self.agc_var.get())
            else:
                self.sources.audio.agc_enabled = self.agc_var.get()
        else:
            self.sources.stop_audio()
        if m in ("SDR", "Auto"):
            if not (self.sources.sdr and self.sources.sdr.running):
                if self.sources.start_sdr():
                    self.sources.sdr.set_frequency(self.rx_freq)
            else:
                self.sources.sdr.set_frequency(self.rx_freq)
        else:
            self.sources.stop_sdr()
        self._sdr_status_upd()
        if getattr(self, "decoder", None):
            self.decoder.reset()

    def _sdr_status_upd(self):
        s = self.sources.sdr
        if s and s.running:
            i = s.get_info()
            self.sdr_status.config(
                text=f"USB {i['center_hz']/1000:.1f} ← RWM {i['rx_hz']/1000:.1f} кГц → "
                     f"{i['tone_hz']:.0f} Гц · ФНЧ ±{i['if_bw']:.0f} Гц\n"
                     f"SR {i['fs_in']:.0f} → {i['fs_dec']:.0f} Гц · блоков {i['blocks']} · "
                     f"ошибок {i['usb_err']}\n"
                     f"Δf {i['offset_hz']:+.1f} Гц · PPM {i['ppm']:+.2f} · "
                     f"захват {'ДА' if i['locked'] else 'НЕТ'} · {i['fft_bin']:.1f} Гц/бин",
                fg=C["green"])
        else:
            self.sdr_status.config(text="Не запущен — смотрите журнал", fg=C["muted"])

    def _cal_ppm(self):
        s = self.sources.sdr
        if not (s and s.running and s.locked):
            self._log("[SDR] калибровка PPM: нужен захват несущей")
            return
        ppm = s.get_info()["ppm"]
        CONFIG["sdr_ppm_cal"] = round(ppm, 3)
        save_config(CONFIG)
        s.ppm_cal = float(CONFIG["sdr_ppm_cal"])
        with s._lock:
            s.carrier_offset_hz = 0.0
        self._log(f"[SDR] ✅ PPM по несущей RWM (5·10⁻¹²): {ppm:+.3f} ppm")

    def _watch(self):
        if not self.running:
            return
        try:
            if (self.sources.mode in ("SDR", "Auto") and CONFIG.get("sdr_autorestart", True)
                    and not (self.sources.sdr and self.sources.sdr.running)
                    and time.time() - self._sdr_restart > 15):
                self._sdr_restart = time.time()
                self._log("[SDR] поток остановлен — авто-перезапуск")
                if self.sources.start_sdr():
                    self.sources.sdr.set_frequency(self.rx_freq)
                self._sdr_status_upd()
        except Exception:
            pass
        self.root.after(5000, self._watch)

    def _test_sdr(self):
        self._log("─" * 60)
        self._save_sdr()
        s = SDRInput(CONFIG, log_func=lambda m: self._log(f"[ТЕСТ] {m}"))
        if s.start():
            self.root.after(1800, lambda: self._test_sdr_done(s))
        else:
            self._log(f"[ТЕСТ] ❌ {s._last_err}")

    def _test_sdr_done(self, s):
        b = s.get_buffer(0.3)
        if b is not None:
            self._log(f"[ТЕСТ] ✅ RMS {float(np.sqrt(np.mean(b.astype(np.float64)**2))):.0f} · "
                      f"Q {s.quality():.2f} · блоков {s.blocks_ok}")
        else:
            self._log("[ТЕСТ] буфер пуст")
        s.close()
        self._log("─" * 60)

    def _save_agc(self):
        CONFIG["agc_enabled"] = self.agc_var.get()
        save_config(CONFIG)
        if self.sources.audio:
            self.sources.audio.agc_enabled = self.agc_var.get()

    def _save_dut1(self):
        CONFIG["decode_dut1"] = self.dut1_var.get()
        save_config(CONFIG)
        self.decoder.decode_dut1 = self.dut1_var.get()

    def _apply_trim(self):
        try:
            v = float(self.trim_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Введите число")
            return
        CONFIG["timing_trim_ms"] = v
        save_config(CONFIG)
        self.decoder.trim_ms = v
        self._log(f"[ВРЕМЯ] подстройка {v:+.3f} мс · полная задержка "
                  f"{self.decoder.total_delay(self.sources.get_active())*1000:.3f} мс")

    def _apply_dec(self):
        try:
            v = {"pulse_center_hz": float(self.tone_var.get()),
                 "pulse_band_hz": float(self.band_var.get()),
                 "pulse_smooth_ms": float(self.smooth_var.get()),
                 "edge_lo_frac": float(self.lo_var.get()),
                 "edge_hi_frac": float(self.hi_var.get()),
                 "gate_coarse_ms": float(self.gc_var.get()),
                 "gate_fine_ms": float(self.gf_var.get()),
                 "dur_tolerance": float(self.dtol_var.get()),
                 "sync_threshold_ms": float(self.thr_var.get())}
        except ValueError:
            messagebox.showerror("Ошибка", "Все поля должны быть числами")
            return
        if not (100 < v["pulse_center_hz"] < 5000) or v["pulse_band_hz"] <= 0:
            messagebox.showerror("Ошибка", "Тон 100…5000 Гц, полоса > 0")
            return
        if not (0.05 <= v["edge_lo_frac"] < v["edge_hi_frac"] <= 0.95):
            messagebox.showerror("Ошибка", "Доли фронта: 0.05 ≤ lo < hi ≤ 0.95")
            return
        CONFIG.update(v)
        save_config(CONFIG)
        d = self.decoder
        d.tone, d.band = v["pulse_center_hz"], v["pulse_band_hz"]
        d.smooth_ms = v["pulse_smooth_ms"]
        d.lo_frac, d.hi_frac = v["edge_lo_frac"], v["edge_hi_frac"]
        d.gate_coarse = v["gate_coarse_ms"] / 1000
        d.gate_fine = v["gate_fine_ms"] / 1000
        d.dur_tol = v["dur_tolerance"]
        d._build()
        d.reset()
        if self.sources.sdr:
            self.sources.sdr.tone_hz = d.tone
            self.sources.sdr.set_frequency(self.rx_freq)
        if self.tline is not None:
            self.tline.set_xdata([d.tone, d.tone])
            self.canvas.draw_idle()
        self._usb_label()
        self._log(f"[ДЕКОДЕР] тон {d.tone:.0f} Гц · полоса {d.band:.0f} Гц · "
                  f"сглаж {d.smooth_ms:.1f} мс · фронт {d.lo_frac:.2f}/{d.hi_frac:.2f} · "
                  f"задержка КИХ {d.fixed_delay_s*1000:.2f} мс")

    def _measure(self):
        self.decoder.measure_until = time.time() + 30.0
        with self.decoder._lock:
            for v in self.decoder._dur.values():
                v.clear()
        self._log("[ИЗМЕР] сбор длительностей 30 с — ожидается 100/500 (A1X) или 20/40/500 (A1N)")

    def _save_ntp(self):
        CONFIG["use_ntp_fallback"] = self.ntp_var.get()
        save_config(CONFIG)
        self.sync.use_ntp = self.ntp_var.get()

    # ---------- обработка ----------
    def _t_decode(self):
        if not self.running:
            return
        self.sources.tick()
        s = self.sources.get_active()
        if s and self.decoder:
            try:
                self.decoder.process(s, RWMSchedule.get_mode())
            except Exception as e:
                self._log(f"[ДЕКОДЕР] ошибка: {e}")
        self.root.after(200, self._t_decode)

    def test_signal(self):
        s = self.sources.get_active()
        if s is None:
            self._log("[ТЕСТ] нет активного источника")
            return
        b = s.get_buffer(2.0)
        if b is None:
            self._log(f"[ТЕСТ:{s.name}] нет данных")
            return
        st = self.decoder.get_status()
        self._log(f"[ТЕСТ:{s.name}] RMS {float(np.sqrt(np.mean(b.astype(np.float64)**2))):.0f} · "
                  f"Q {s.quality():.2f} · SNR {st['last_snr']:.1f} · "
                  f"меток {st['accepted']} · {st['hit_rate']*100:.0f}% · "
                  f"задержка полная {self.decoder.total_delay(s)*1000:.3f} мс "
                  f"(КИХ {st['delay_ms']:.2f} + тракт {s.latency_sec*1000:.2f} + "
                  f"путь {st['path_ms']:.2f} + подр {st['trim_ms']:+.2f})")

    # ---------- журнал ----------
    def _log(self, msg, level=None):
        ts = datetime.now(timezone.utc).strftime("%H:%M:%S.%f")[:-3]
        m = str(msg)
        tag = "dim"
        if "★" in m:
            tag = "mark"
        elif "❌" in m or "ОШИБКА" in m or "ошибка" in m or "Не могу" in m:
            tag = "err"
        elif "⚠" in m or "ПРЕДУПРЕЖД" in m:
            tag = "warn"
        elif "✅" in m or "✓" in m or "апущен" in m or "ВКЛЮЧ" in m:
            tag = "ok"
        elif m.startswith("[NTP"):
            tag = "ntp"
        elif m.startswith(("[ГЕО]", "[DUT1]", "[ВРЕМЯ]")):
            tag = "geo"
        elif m.startswith(("[SDR]", "[ИСТОЧНИК]", "[МОНИТОР]", "[ДЕКОДЕР]", "[ИЗМЕР]")):
            tag = "info"
        self.log_q.put((f"[{ts}] {m}", tag))
        print(f"[{ts}] {m}")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _t_log(self):
        if not self.running:
            return
        try:
            while True:
                line, tag = self.log_q.get_nowait()
                self.log_text.configure(state="normal")
                self.log_text.insert("end", line + "\n", tag)
                if int(self.log_text.index("end-1c").split(".")[0]) > 4000:
                    self.log_text.delete("1.0", "500.0")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(100, self._t_log)

    def _t_diag(self):
        if not self.running:
            return
        parts = []
        for s in (self.sources.audio, self.sources.sdr):
            if s and s.running:
                parts.append(f"{s.name} RMS {s.get_raw_rms():.0f} · g {s.agc_gain:.1f} · "
                             f"{'OK' if s.data_alive() else 'НЕТ ДАННЫХ'}")
        if parts:
            self._log("[ДИАГ] " + " │ ".join(parts))
        s = self.sources.sdr
        if s and s.running:
            i = s.get_info()
            self._log(f"[SDR] Δf {i['offset_hz']:+.2f} Гц · PPM {i['ppm']:+.2f} · "
                      f"{i['level_db']:.1f} dB · захват {'ДА' if i['locked'] else 'НЕТ'} · "
                      f"блоков {i['blocks']}")
            self._sdr_status_upd()
        if self.ntp.enabled:
            self._ntp_status_upd()
        self.root.after(4000, self._t_diag)

    # ---------- статус ----------
    def _t_sched(self):
        if not self.running:
            return
        now = datetime.now(timezone.utc)
        self.utc_lbl.config(text=now.strftime("%H:%M:%S UTC"))
        m = RWMSchedule.get_mode()
        txt = SEG_LABELS.get(m, m)
        if RWMSchedule.seconds_skipped(now.minute) and now.second >= 56:
            txt += " · пропуск 56–59"
        elif not RWMSchedule.is_sync_possible(m):
            s = RWMSchedule.seconds_to_next_sync()
            txt += f" · окно через {s//60}м"
        self.mode_badge.config(text=txt, bg=SEG_COLORS.get(m, C["surface2"]), fg=C["text"])
        self.sched.update_position()
        self.root.after(1000, self._t_sched)

    def _t_status(self):
        if not self.running:
            return
        s = self.sources.get_active()
        title = s.title if s else "нет источника"
        extra = ""
        if s is not None and s.name == "sdr":
            i = self.sources.sdr.get_info()
            extra = f" · Δf {i['offset_hz']:+.1f} Гц"
            if not s.data_alive():
                extra += " · НЕТ ДАННЫХ"
        self.src_info.config(text=f"{title}{extra}",
                             fg=C["green"] if (s and s.data_alive()) else C["red"])
        st = self.decoder.get_status()
        self.pulse_lbl.config(text=f"меток {st['accepted']} · {st['hit_rate']*100:.0f}% · "
                                   f"мин {st['minute_marks']} · проп {st['skipped']}")
        self.dut1_lbl.config(text=f"DUT1 {st['dut1']:+.3f} c" if st["dut1"] is not None
                             else "DUT1 —")
        if self.sync.running:
            good = st["hit_rate"] > 0.3 and st["n"] >= 3
            self.status_lbl.config(text="СИНХРОНИЗИРОВАНО" if good else "поиск меток…",
                                   fg=C["green"] if good else C["amber"])
            self.corr_lbl.config(text=f"коррекций {self.sync.correction_count}")
            o = st["offset_ms"]
            if o is not None:
                col = C["green"] if abs(o) < 10 else (C["amber"] if abs(o) < 100 else C["red"])
                self.off_lbl.config(text=f"Δ {o:+.2f} мс", fg=col)
        else:
            self.status_lbl.config(text="Остановлен", fg=C["muted"])
        self._mon_btn()
        if not self.monitor.enabled:
            self.mon_state.config(text="монитор выключен", fg=C["dim"])
        elif self.monitor.playing:
            self.mon_state.config(text=f"🔊 {self.monitor.out_rate} Гц · {title}", fg=C["green"])
        elif self.monitor.last_error:
            self.mon_state.config(text=f"ошибка: {self.monitor.last_error[:34]}", fg=C["red"])
        else:
            self.mon_state.config(text="ожидание сигнала…", fg=C["amber"])
        ns = self.ntp.get_stats()
        if ns["running"]:
            self.ntp_button.configure(text="🕓  NTP-сервер", bg="#0f3b3a", fg="#c9fff8")
            self.ntp_button._pal = ("#0f3b3a", "#c9fff8", "#237a72", C["cyan"])
            sv = self.sync.server_status()
            if sv["rwm_age"] is not None and sv["rwm_age"] < 180:
                badge, col = f":{ns['port']} · S1 RWM · {ns['served']}", C["green"]
            elif sv["ntp_age"] is not None and sv["ntp_age"] < 900:
                badge, col = f":{ns['port']} · S2 NTP · {ns['served']}", C["cyan"]
            else:
                badge, col = f":{ns['port']} · S3 LOCL · {ns['served']}", C["amber"]
            self.ntp_badge.config(text=badge, fg=col)
        else:
            self.ntp_button.configure(text="🕓  NTP-сервер", bg=C["surface2"], fg=C["muted"])
            self.ntp_button._pal = (C["surface2"], C["muted"], C["border"], C["cyan"])
            self.ntp_badge.config(text="ошибка порта" if ns["error"] else "выкл",
                                  fg=C["red"] if ns["error"] else C["dim"])
        self.root.after(400, self._t_status)

    def _t_vu(self):
        if not self.running:
            return
        s = self.sources.get_active()
        v = s.get_rms() if s else 0.0
        self.vu_in.set(v)
        self.in_val.config(text=f"{int(v*100)}%")
        o = self.monitor.out_rms
        self.vu_out.set(o)
        self.out_val.config(text=f"{int(o*100)}%")
        self.root.after(100, self._t_vu)

    def _t_wf(self):
        if not self.running:
            return
        s = self.sources.get_active()
        alive = bool(s and s.data_alive())
        if self.canvas is not None:
            if not alive and not self._wf_nodata:
                self._wf_nodata = True
                self.wf_state.config(text="НЕТ ДАННЫХ", fg=C["red"])
            elif alive and self._wf_nodata:
                self._wf_nodata = False
                self.wf_state.config(text="приём", fg=C["green"])
        if alive and self.wf_img is not None:
            b = s.get_buffer(0.25)
            if b is not None and len(b) >= self.FFT:
                seg = b[-self.FFT:].astype(np.float64) * np.hanning(self.FFT)
                mag = np.abs(np.fft.rfft(seg))
                fr = np.fft.rfftfreq(self.FFT, 1 / BufferedSource.SAMPLE_RATE)
                mag = mag[:max(2, int(np.searchsorted(fr, self.WF_MAX)))]
                if mag.size:
                    row = resize_1d(20 * np.log10(mag / (self.FFT / 2) + 1e-9), self.WF_BINS)
                    if row is not None and len(row) == self.WF_BINS:
                        self.wf_data = np.roll(self.wf_data, -1, axis=0)
                        self.wf_data[-1, :] = row
                        self.wf_img.set_data(self.wf_data)
                        lo = float(np.percentile(self.wf_data, 10))
                        hi = float(np.percentile(self.wf_data, 99.5))
                        if hi - lo < 6:
                            mid = 0.5 * (hi + lo)
                            lo, hi = mid - 3, mid + 3
                        self.wf_img.set_clim(lo, hi)
                        self.canvas.draw_idle()
        self.root.after(self.WF_MS, self._t_wf)

    # ---------- синхронизация ----------
    def start_sync(self):
        if not self.sources.any_running():
            self._apply_sources()
        if not self.sources.any_running():
            self._log("[ОШИБКА] ни один источник не работает")
            return
        m = RWMSchedule.get_mode()
        if not RWMSchedule.is_sync_possible(m):
            s = RWMSchedule.seconds_to_next_sync()
            self._log(f"[ПРЕДУПРЕЖДЕНИЕ] {m}: меток нет, окно через {s//60} мин {s%60} с")
        self.decoder.reset()
        self.sync.start()
        self._log(f"[СИНХРОНИЗАЦИЯ] запущена · {self.sources.active_title()} · "
                  f"задержка {self.decoder.total_delay(self.sources.get_active())*1000:.3f} мс")

    def stop_sync(self):
        self.sync.stop()

    def set_freq(self, f=None):
        if f is None:
            try:
                f = int(float(self.freq_var.get()))
            except ValueError:
                messagebox.showerror("Ошибка", "Введите частоту в Гц")
                return
        f = int(f)
        if not 100000 <= f <= 30000000:
            messagebox.showerror("Ошибка", "Диапазон 0.1–30 МГц")
            return
        self.rx_freq = f
        CONFIG["rx_freq"] = f
        save_config(CONFIG)
        self.freq_var.set(str(f))
        try:
            off = int(float(self.cat_off_var.get()))
        except Exception:
            off = -2000
        CONFIG["cat_freq_offset"] = off
        save_config(CONFIG)
        self._send_freq(f + off)
        if self.sources.sdr:
            self.sources.sdr.set_frequency(f)
        self.decoder.reset()
        self._paint_all_freq()
        self._usb_label()
        self._log(f"[ЧАСТОТА] RWM {f} │ CAT {f+off} │ "
                  f"SDR USB {f-int(CONFIG['pulse_center_hz'])} Гц")

    def set_rwm(self, f):
        self.set_freq(f)
        self.root.after(120, self._send_ud)

    def _send_freq(self, hz):
        fr = self._civ(hz)
        if self.cat_mode == "Direct CAT":
            self.direct_cat.write(fr)
        elif self.sniffer:
            self.sniffer.write_real(fr)

    def _send_ud(self):
        f = bytes([0xFE, 0xFE, 0x88, 0xE0, 0x26, 0x00, 0x01, 0x01, 0x01, 0xFD])
        ok = (self.direct_cat.write(f) if self.cat_mode == "Direct CAT"
              else (self.sniffer.write_real(f) if self.sniffer else False))
        if ok:
            self._log("[CAT] команда режима отправлена")

    def _civ(self, hz):
        s = f"{int(hz):010d}"
        return (bytes([0xFE, 0xFE, 0x88, 0xE0, 0x05]) +
                bytes(int(s[i:i + 2], 16) for i in range(8, -2, -2)) + bytes([0xFD]))

    def _apply_cat(self, startup=False):
        m = self.cat_mode_var.get().strip()
        p = self.cat_com_var.get().strip()
        try:
            b = int(self.cat_baud_var.get() or 19200)
        except Exception:
            b = 19200
        CONFIG.update({"cat_mode": m, "cat_com_port": p, "cat_baudrate": b})
        save_config(CONFIG)
        self.cat_mode = m
        if self.sniffer:
            self.sniffer.stop()
            self.sniffer = None
        self.direct_cat.close()
        if m == "Direct CAT":
            if not p:
                self.cat_status.config(text="COM не выбран", fg=C["muted"])
                return
            if self.direct_cat.configure(p, b):
                self.cat_status.config(text=f"Direct CAT · {p} · {b}", fg=C["green"])
                self._log(f"[CAT] {p} @ {b}")
            else:
                self.cat_status.config(text=f"Не удалось открыть {p}", fg=C["red"])
                self._log(f"[ОШИБКА] CAT: {p} занят или недоступен")
        else:
            if not p or not SERIAL_AVAILABLE:
                self.cat_status.config(text="COM/pyserial недоступны", fg=C["red"])
                return
            try:
                vp = CONFIG.get("cat_virtual_port", "COM9")
                self.sniffer = SerialSniffer(p, vp, b, self._log)
                self.sniffer.start()
                self.cat_status.config(text=f"Sniffer · {p} ↔ {vp}", fg=C["green"])
            except Exception as e:
                self.cat_status.config(text="Ошибка Sniffer", fg=C["red"])
                self._log(f"[ОШИБКА] Sniffer: {e}")

    def on_close(self):
        self.running = False
        try:
            self.sync.stop()
        except Exception:
            pass
        self.monitor.stop()
        self.ntp.stop()
        self.sources.stop_all()
        self.direct_cat.close()
        if self.sniffer:
            self.sniffer.stop()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = TimeSyncApp(root)
    root.mainloop()