# -*- coding: utf-8 -*-
"""
Радиосинхронизация времени по ГОСТ 8.323-2016 (сигналы RWM)

Полная исправленная версия:
- AGC, VU-метр, водопад, непрерывная обработка аудио.
- Учтены режимы RWM: SECOND_PULSES / TENTH_PULSES.
- Добавлен полосовой фильтр, сглаживание огибающей и гистерезис.
- Водопад: 0...3000 Гц, частота по оси X, время по оси Y.
- Лог: белый текст.
- Кнопки управления вынесены вверх и не перекрываются водопадом.
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
from datetime import datetime, timedelta, timezone


# ------------------------------------------------------------
# Автоустановка зависимостей (опционально)
# ------------------------------------------------------------
def ensure_packages():
    required = {
        "numpy": "numpy",
        "pyaudio": "pyaudio",
        "serial": "pyserial",
        "matplotlib": "matplotlib",
    }

    if sys.platform.startswith("win"):
        required["win32com"] = "pywin32"

    for mod, pkg in required.items():
        try:
            importlib.import_module(mod)
        except Exception:
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", pkg],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
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
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    MATPLOTLIB_AVAILABLE = True
except Exception:
    MATPLOTLIB_AVAILABLE = False
    Figure = None
    FigureCanvasTkAgg = None

try:
    import serial
    import serial.tools.list_ports
    SERIAL_AVAILABLE = True
except Exception:
    SERIAL_AVAILABLE = False


# ------------------------------------------------------------
# EXE-совместимая директория
# ------------------------------------------------------------
def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

CONFIG_FILE = os.path.join(APP_DIR, "config.json")


# ------------------------------------------------------------
# Конфигурация
# ------------------------------------------------------------
DEFAULT_CONFIG = {
    "audio_input": "",
    "audio_output": "",
    "cat_mode": "Direct CAT",
    "cat_com_port": "",
    "cat_virtual_port": "COM9",
    "cat_baudrate": 19200,
    "agc_enabled": True,
    "rx_freq": 4996000,
    "auto_threshold": True,
    "manual_threshold": 0.3,
    "use_ntp_fallback": True,
    "ntp_servers": [
        "0.ru.pool.ntp.org",
        "1.ru.pool.ntp.org"
    ],

    # Настройки фильтра импульсов.
    # Если приём через offset -2000 Гц, обычно звуковой тон около 2000 Гц.
    "pulse_center_hz": 2000,
    "pulse_band_hz": 300,
    "pulse_smooth_ms": 5,
}


def _clean_config_value(v):
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, list):
        return [_clean_config_value(x) for x in v]
    if isinstance(v, dict):
        return {str(k).strip(): _clean_config_value(val) for k, val in v.items()}
    return v


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            if isinstance(data, dict):
                data = {str(k).strip(): _clean_config_value(v) for k, v in data.items()}
                merged = DEFAULT_CONFIG.copy()
                merged.update(data)
                return merged
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


# ------------------------------------------------------------
# Вспомогательная функция для водопада
# ------------------------------------------------------------
def resize_1d(arr, size):
    if np is None:
        return None

    if arr is None or len(arr) == 0:
        return np.zeros(size, dtype=np.float64)

    if len(arr) == size:
        return arr.astype(np.float64)

    if len(arr) == 1:
        return np.full(size, float(arr[0]), dtype=np.float64)

    x_old = np.linspace(0.0, 1.0, len(arr))
    x_new = np.linspace(0.0, 1.0, size)
    return np.interp(x_new, x_old, arr)


# ------------------------------------------------------------
# Расписание работы RWM (ГОСТ 8.323-2016)
# ------------------------------------------------------------
class RWMSchedule:
    MODE_N0N = "N0N"
    MODE_OFF = "OFF"
    MODE_MORSE = "MORSE"
    MODE_SECOND_PULSES = "SECOND_PULSES"
    MODE_TENTH_PULSES = "TENTH_PULSES"

    # Цикл 30 минут, затем повторяется.
    SCHEDULE = [
        (0, 8, MODE_N0N),
        (8, 9, MODE_OFF),
        (9, 10, MODE_MORSE),
        (10, 20, MODE_SECOND_PULSES),
        (20, 30, MODE_TENTH_PULSES),
    ]

    @classmethod
    def get_mode(cls, dt=None):
        if dt is None:
            dt = datetime.now(timezone.utc)

        minute_in_cycle = dt.minute % 30

        for start, end, mode in cls.SCHEDULE:
            if start <= minute_in_cycle < end:
                return mode

        return cls.MODE_N0N

    @classmethod
    def get_mode_description(cls, mode):
        descriptions = {
            cls.MODE_N0N: "Немодулированная несущая (нет импульсов)",
            cls.MODE_OFF: "Передатчик выключен",
            cls.MODE_MORSE: "Позывной RWM азбукой Морзе",
            cls.MODE_SECOND_PULSES: "Секундные импульсы (1 Гц) — пригоден для синхронизации",
            cls.MODE_TENTH_PULSES: "10 импульсов/сек — для высокоточных измерений",
        }
        return descriptions.get(mode, mode)

    @classmethod
    def is_sync_possible(cls, mode):
        return mode in (cls.MODE_SECOND_PULSES, cls.MODE_TENTH_PULSES)


# ------------------------------------------------------------
# CAT / PTT
# ------------------------------------------------------------
class DirectCAT:
    def __init__(self):
        self.ser = None
        self.lock = threading.Lock()

    def configure(self, port, baudrate=19200):
        self.close()

        if not port:
            return True

        if not SERIAL_AVAILABLE:
            return False

        try:
            self.ser = serial.Serial(
                port=port,
                baudrate=baudrate,
                timeout=0.1,
                write_timeout=0.5
            )
            return True
        except Exception:
            return False

    def write(self, data):
        if self.ser is None or not self.ser.is_open:
            return False

        with self.lock:
            try:
                self.ser.write(data)
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
    def __init__(self, real_port, virtual_port, baudrate, log_func):
        self.real_port = real_port
        self.virtual_port = virtual_port
        self.baudrate = baudrate
        self.log_func = log_func
        self.real = None
        self.virtual = None
        self.running = False

    def start(self):
        if not SERIAL_AVAILABLE:
            raise RuntimeError("pyserial недоступен")

        try:
            self.real = serial.Serial(
                port=self.real_port,
                baudrate=self.baudrate,
                timeout=0.05,
                write_timeout=0.5
            )

            self.virtual = serial.Serial(
                port=self.virtual_port,
                baudrate=self.baudrate,
                timeout=0.05,
                write_timeout=0.5
            )
        except Exception:
            self.stop()
            raise

        self.running = True
        threading.Thread(target=self._v2r, daemon=True).start()
        threading.Thread(target=self._r2v, daemon=True).start()

        if self.log_func:
            self.log_func(f"Sniffer: {self.virtual_port} <-> {self.real_port}")

    def write_real(self, data):
        if self.real and self.real.is_open:
            try:
                self.real.write(data)
                return True
            except Exception:
                return False

        return False

    def _v2r(self):
        while self.running:
            try:
                data = self.virtual.read(4096)
                if data:
                    self.write_real(data)
            except Exception:
                time.sleep(0.2)

    def _r2v(self):
        while self.running:
            try:
                data = self.real.read(4096)
                if data and self.virtual and self.virtual.is_open:
                    self.virtual.write(data)
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

        self.real = None
        self.virtual = None


# ------------------------------------------------------------
# Аудио захват с автоматическим усилением
# ------------------------------------------------------------
class AudioInput:
    SAMPLE_RATE = 12000
    CHUNK = 1024
    BUFFER_SECONDS = 5

    def __init__(self, device_keywords="", agc_enabled=True):
        self.device_keywords = device_keywords
        self.agc_enabled = agc_enabled
        self.pya = pyaudio.PyAudio() if AUDIO_AVAILABLE else None
        self.stream = None

        if np is not None:
            self.audio_buffer = np.zeros((0,), dtype=np.int16)
        else:
            self.audio_buffer = None

        self.buffer_lock = threading.Lock()
        self.running = False
        self.agc_gain = 1.0
        self.last_rms = 0.0
        self.device_index = None

        self._find_input_device()

    def _find_input_device(self):
        if self.pya is None:
            self.device_index = None
            return

        keywords = [
            k.strip().lower()
            for k in self.device_keywords.split(",")
            if k.strip()
        ]

        for dev_idx in range(self.pya.get_device_count()):
            try:
                info = self.pya.get_device_info_by_index(dev_idx)
                name = info.get("name", "")

                if info.get("maxInputChannels", 0) > 0:
                    if not keywords or all(kw in name.lower() for kw in keywords):
                        self.device_index = dev_idx
                        print(f"[AudioInput] Found: #{dev_idx} {name}")
                        return
            except Exception:
                continue

        try:
            self.device_index = self.pya.get_default_input_device_info()["index"]
            print(f"[AudioInput] Fallback default #{self.device_index}")
        except Exception:
            self.device_index = None

    def start(self):
        if not AUDIO_AVAILABLE:
            return False

        if self.pya is None or self.device_index is None:
            return False

        self.running = True

        self.stream = self.pya.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=self.SAMPLE_RATE,
            input=True,
            input_device_index=self.device_index,
            frames_per_buffer=self.CHUNK,
            stream_callback=self._callback
        )

        self.stream.start_stream()
        return True

    def _callback(self, in_data, frame_count, time_info, status):
        if not self.running:
            return (None, pyaudio.paAbort)

        arr = np.frombuffer(in_data, dtype=np.int16).astype(np.float64)

        if self.agc_gain != 1.0:
            arr *= self.agc_gain
            arr = np.clip(arr, -32768, 32767).astype(np.int16)
        else:
            arr = arr.astype(np.int16)

        with self.buffer_lock:
            if self.audio_buffer is None:
                self.audio_buffer = arr
            else:
                self.audio_buffer = np.concatenate((self.audio_buffer, arr))

            max_len = self.SAMPLE_RATE * self.BUFFER_SECONDS
            if len(self.audio_buffer) > max_len:
                self.audio_buffer = self.audio_buffer[-max_len:]

        if self.agc_enabled:
            rms = float(np.sqrt(np.mean(np.square(arr.astype(np.float64)))))
            self.last_rms = rms

            # AGC должен не только усиливать, но и ослаблять сигнал.
            if rms < 100 and self.agc_gain < 20.0:
                self.agc_gain *= 1.05
            elif rms > 6000 and self.agc_gain > 0.05:
                self.agc_gain /= 1.10
        else:
            self.last_rms = float(np.sqrt(np.mean(np.square(arr.astype(np.float64)))))

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

    def get_buffer(self, duration_sec):
        if not AUDIO_AVAILABLE or self.audio_buffer is None:
            return None

        need = int(duration_sec * self.SAMPLE_RATE)

        with self.buffer_lock:
            if len(self.audio_buffer) < need:
                return None

            return self.audio_buffer[-need:].copy()

    def get_rms(self):
        if not AUDIO_AVAILABLE or self.audio_buffer is None:
            return 0.0

        with self.buffer_lock:
            if len(self.audio_buffer) == 0:
                return 0.0

            rms = float(np.sqrt(np.mean(np.square(self.audio_buffer.astype(np.float64)))))
            return min(1.0, rms / 32768.0)

    def get_raw_rms(self):
        if not AUDIO_AVAILABLE or self.audio_buffer is None:
            return 0.0

        with self.buffer_lock:
            if len(self.audio_buffer) == 0:
                return 0.0

            return float(np.sqrt(np.mean(np.square(self.audio_buffer.astype(np.float64)))))

    def get_noise_floor(self):
        if not AUDIO_AVAILABLE or self.audio_buffer is None:
            return 0.0

        with self.buffer_lock:
            if len(self.audio_buffer) < 512:
                return 0.0

            samples = self.audio_buffer[-int(self.SAMPLE_RATE * 0.5):]
            abs_vals = np.abs(samples.astype(np.float64))
            return float(np.median(abs_vals) / 32768.0)

    def set_agc_gain(self, gain):
        self.agc_gain = float(gain)

    def close(self):
        self.stop()

        if self.pya:
            try:
                self.pya.terminate()
            except Exception:
                pass

            self.pya = None


# ------------------------------------------------------------
# Декодер сигналов времени RWM
# ------------------------------------------------------------
class RadioTimeDecoder:
    SAMPLE_RATE = AudioInput.SAMPLE_RATE

    def __init__(self, threshold=0.3, log_func=None, auto_threshold=True):
        self.manual_threshold = threshold
        self.auto_threshold = auto_threshold
        self.log_func = log_func

        try:
            self.filter_center_hz = float(CONFIG.get("pulse_center_hz", 2000.0))
            self.filter_band_hz = float(CONFIG.get("pulse_band_hz", 300.0))
            self.smooth_ms = float(CONFIG.get("pulse_smooth_ms", 5.0))
        except Exception:
            self.filter_center_hz = 2000.0
            self.filter_band_hz = 300.0
            self.smooth_ms = 5.0

        self.last_impulse_time = None
        self.impulse_count = 0
        self.minute_start_time = None
        self.is_synced = False
        self.sync_lock = threading.Lock()

        self.current_second = 0
        self.last_adaptive_threshold = 0.3
        self.last_snr = 0.0
        self.last_max_amp = 0.0
        self.last_noise = 0.0
        self.last_result = None

        self.last_diag_log = 0.0

    def _bandpass(self, sig):
        """
        Простой частотный фильтр через FFT.
        Выделяет узкую область вокруг pulse_center_hz.
        """
        if sig is None or len(sig) < 256:
            return sig

        if self.filter_center_hz <= 0:
            return sig

        try:
            n = len(sig)
            spectrum = np.fft.rfft(sig)
            freqs = np.fft.rfftfreq(n, 1.0 / self.SAMPLE_RATE)

            nyq = self.SAMPLE_RATE / 2.0
            center = float(self.filter_center_hz)

            if center <= 50.0 or center >= nyq - 50.0:
                return sig

            sigma = max(float(self.filter_band_hz) / 2.3548, 5.0)
            mask = np.exp(-0.5 * ((freqs - center) / sigma) ** 2)

            # Убрать DC.
            mask[0] = 0.0

            spectrum *= mask
            return np.fft.irfft(spectrum, n)

        except Exception:
            return sig

    def _smooth_envelope(self, env):
        """
        Сглаживание огибающей, чтобы убрать всплески 0.2-1 мс.
        """
        if env is None or len(env) == 0:
            return env

        win = int(self.SAMPLE_RATE * max(self.smooth_ms, 1.0) / 1000.0)
        if win < 3:
            return env

        kernel = np.ones(win, dtype=np.float64) / float(win)
        return np.convolve(env, kernel, mode="same")

    def _detect_pulses_hysteresis(self, env, high, low, min_gap_samples):
        """
        Детекция импульсов с гистерезисом:
        high - порог входа,
        low - порог выхода.
        """
        starts = []
        ends = []

        in_pulse = False
        start = 0

        for i, v in enumerate(env):
            if not in_pulse:
                if v >= high:
                    in_pulse = True
                    start = i
            else:
                if v <= low:
                    in_pulse = False
                    starts.append(start)
                    ends.append(i)

        if in_pulse:
            starts.append(start)
            ends.append(len(env))

        if not starts:
            return np.array([], dtype=int), np.array([], dtype=int)

        # Склейка импульсов, если между ними очень маленький разрыв.
        merged_starts = []
        merged_ends = []

        for s, e in zip(starts, ends):
            if merged_starts and (s - merged_ends[-1]) <= min_gap_samples:
                if e > merged_ends[-1]:
                    merged_ends[-1] = e
            else:
                merged_starts.append(s)
                merged_ends.append(e)

        return np.array(merged_starts, dtype=int), np.array(merged_ends, dtype=int)

    def _dominant_freq(self, sig):
        """
        Диагностика: доминирующая частота в сигнале.
        """
        try:
            if sig is None or len(sig) < 512:
                return 0.0

            window = np.hanning(len(sig))
            spectrum = np.fft.rfft(sig * window)
            freqs = np.fft.rfftfreq(len(sig), 1.0 / self.SAMPLE_RATE)
            mag = np.abs(spectrum)

            # Диапазон диагностики ограничен 0...3000 Гц.
            mask = (freqs > 50.0) & (freqs < min(3000.0, self.SAMPLE_RATE / 2.0))
            if not np.any(mask):
                return 0.0

            freqs = freqs[mask]
            mag = mag[mask]

            if len(mag) == 0:
                return 0.0

            idx = int(np.argmax(mag))
            return float(freqs[idx])

        except Exception:
            return 0.0

    def process_buffer(self, samples, noise_floor=None, mode=None):
        """
        Обрабатывает буфер сэмплов.
        mode: строка, возвращаемая RWMSchedule.get_mode()
        """

        if samples is None or len(samples) < 100:
            return None

        if mode not in (RWMSchedule.MODE_SECOND_PULSES, RWMSchedule.MODE_TENTH_PULSES):
            return {
                "status": "mode_not_sync",
                "mode": mode,
            }

        # Нормализация и удаление постоянной составляющей.
        raw = samples.astype(np.float64) / 32768.0
        raw -= float(np.mean(raw))

        # Полосовой фильтр вокруг ожидаемой частоты импульсов.
        filtered = self._bandpass(raw)

        # Огибающая.
        envelope = np.abs(filtered)
        envelope = self._smooth_envelope(envelope)

        if envelope is None or len(envelope) == 0:
            return None

        max_amp = float(np.max(envelope))
        self.last_max_amp = max_amp

        if max_amp < 0.001:
            return {
                "status": "no_signal",
                "max_amp": max_amp,
            }

        median = float(np.median(envelope))
        self.last_noise = median

        snr_ratio = max_amp / (median + 1e-6)
        self.last_snr = snr_ratio

        # ------------------------------------------------------------
        # Адаптивный порог
        # ------------------------------------------------------------
        if self.auto_threshold:
            if snr_ratio > 10:
                coef = 0.08
            elif snr_ratio > 3:
                coef = 0.15
            else:
                coef = 0.25

            threshold_val = median + coef * (max_amp - median)
        else:
            threshold_val = max_amp * self.manual_threshold

        # Защита от слишком низкого порога.
        threshold_val = float(max(threshold_val, median * 1.5, 0.004))
        self.last_adaptive_threshold = threshold_val

        high = threshold_val
        low = threshold_val * 0.6

        # ------------------------------------------------------------
        # Детекция импульсов
        # ------------------------------------------------------------
        min_gap_samples = int(0.015 * self.SAMPLE_RATE)

        starts, ends = self._detect_pulses_hysteresis(
            envelope,
            high=high,
            low=low,
            min_gap_samples=min_gap_samples
        )

        if len(starts) == 0 or len(ends) == 0:
            return {
                "status": "no_pulses",
                "max_amp": max_amp,
                "threshold": threshold_val,
                "adaptive": self.auto_threshold,
                "snr_ratio": snr_ratio,
                "noise": median,
            }

        # ------------------------------------------------------------
        # Допустимые длительности по режиму RWM
        # ------------------------------------------------------------
        if mode == RWMSchedule.MODE_TENTH_PULSES:
            min_dur_s = 0.008
            max_dur_s = 0.065
            minute_marker_dur_s = 0.450

        elif mode == RWMSchedule.MODE_SECOND_PULSES:
            min_dur_s = 0.020
            max_dur_s = 0.250
            minute_marker_dur_s = 0.450

        else:
            return {
                "status": "mode_not_sync",
                "mode": mode,
            }

        valid_indices = []
        rejected_durations = []

        for i in range(len(starts)):
            start = int(starts[i])
            end = int(ends[i])
            duration_samples = end - start

            if duration_samples < 3:
                continue

            duration_s = duration_samples / self.SAMPLE_RATE

            is_valid_duration = False

            if min_dur_s <= duration_s <= max_dur_s:
                is_valid_duration = True
            elif minute_marker_dur_s is not None and min_dur_s <= duration_s <= minute_marker_dur_s + 0.100:
                is_valid_duration = True

            if not is_valid_duration:
                rejected_durations.append(duration_s)
                continue

            seg = envelope[start:end]
            if len(seg) == 0:
                continue

            # Импульс должен быть заметно выше порога.
            if float(np.max(seg)) < threshold_val * 1.10:
                continue

            valid_indices.append(i)

        # ------------------------------------------------------------
        # Диагностика, если нет подходящих импульсов
        # ------------------------------------------------------------
        if len(valid_indices) == 0:
            now = time.time()

            if self.log_func and rejected_durations and now - self.last_diag_log > 5.0:
                self.last_diag_log = now

                sample = [f"{d * 1000:.1f} мс" for d in rejected_durations[:5]]
                extra = f" и ещё {len(rejected_durations) - 5}" if len(rejected_durations) > 5 else ""

                dom_freq = self._dominant_freq(raw)

                self.log_func(
                    f"[ДЕКОДЕР] Отброшены импульсы: {', '.join(sample)}{extra} "
                    f"(допустимо: {min_dur_s * 1000:.1f}–{max_dur_s * 1000:.1f} мс)"
                )

                if dom_freq > 0:
                    self.log_func(
                        f"[ДЕКОДЕР] Доминирующая частота: {dom_freq:.0f} Гц. "
                        f"Если импульсы RWM видны на другой частоте, измените pulse_center_hz "
                        f"(сейчас: {self.filter_center_hz:.0f} Гц)."
                    )

            return {
                "status": "no_valid_pulses",
                "max_amp": max_amp,
                "threshold": threshold_val,
                "adaptive": self.auto_threshold,
                "snr_ratio": snr_ratio,
                "noise": median,
                "rejected_durations": rejected_durations,
            }

        # Берём самый длинный из валидных.
        lengths = ends[valid_indices] - starts[valid_indices]
        best_local_idx = int(np.argmax(lengths))
        best_idx = valid_indices[best_local_idx]

        start_sample = int(starts[best_idx])
        end_sample = int(ends[best_idx])

        duration_samples = end_sample - start_sample
        duration_s = duration_samples / self.SAMPLE_RATE

        mid_sample = start_sample
        t_offset = mid_sample / self.SAMPLE_RATE

        now = time.time()
        buf_duration = len(samples) / self.SAMPLE_RATE
        impulse_abs_time = now - (buf_duration - t_offset)

        is_minute_marker = False
        if minute_marker_dur_s is not None and duration_s >= minute_marker_dur_s - 0.080:
            is_minute_marker = True

        result = {
            "status": "pulse_detected",
            "impulse_time": impulse_abs_time,
            "duration_s": duration_s,
            "is_minute_marker": is_minute_marker,
            "max_amp": max_amp,
            "threshold": threshold_val,
            "adaptive": self.auto_threshold,
            "snr_ratio": snr_ratio,
            "noise": median,
            "mode": mode,
        }

        # ------------------------------------------------------------
        # Обновление состояния синхронизации
        # ------------------------------------------------------------
        with self.sync_lock:
            if mode == RWMSchedule.MODE_TENTH_PULSES:
                min_delta = 0.04
                max_delta = 0.16
            else:
                min_delta = 0.7
                max_delta = 1.3

            delta = None
            if self.last_impulse_time is not None:
                delta = impulse_abs_time - self.last_impulse_time

            was_synced = self.is_synced

            if is_minute_marker:
                self.minute_start_time = impulse_abs_time
                self.is_synced = True
                self.impulse_count = 0
                self.current_second = 0

                if self.log_func:
                    prefix = (
                        "[ДЕКОДЕР] ★ МИНУТНЫЙ МАРКЕР ★"
                        if was_synced
                        else "[ДЕКОДЕР] ★ ПЕРВЫЙ МИНУТНЫЙ МАРКЕР ★"
                    )

                    self.log_func(
                        f"{prefix} обнаружен! "
                        f"Время: {datetime.fromtimestamp(impulse_abs_time).strftime('%H:%M:%S')}, "
                        f"длит: {duration_s * 1000:.0f} мс, "
                        f"порог: {threshold_val:.4f}, "
                        f"SNR: {snr_ratio:.1f}"
                    )

            elif delta is not None and min_delta < delta < max_delta:
                self.impulse_count += 1

                if self.minute_start_time is not None:
                    self.current_second = int(round((impulse_abs_time - self.minute_start_time) % 60))
                else:
                    self.current_second = 0

                if (
                    self.is_synced and
                    self.minute_start_time is not None and
                    self.impulse_count % 5 == 0 and
                    self.log_func
                ):
                    self.log_func(
                        f"[ДЕКОДЕР] Секунда {self.current_second:02d} подтверждена "
                        f"(всего: {self.impulse_count}, "
                        f"длит: {duration_s * 1000:.0f} мс, "
                        f"порог: {threshold_val:.4f}, "
                        f"SNR: {snr_ratio:.1f})"
                    )

            self.last_impulse_time = impulse_abs_time
            self.last_result = result

        return result

    def get_minute_start(self):
        with self.sync_lock:
            if self.is_synced and self.minute_start_time:
                return self.minute_start_time
            return None

    def get_status(self):
        with self.sync_lock:
            return {
                "is_synced": self.is_synced,
                "impulse_count": self.impulse_count,
                "minute_start": self.minute_start_time,
                "current_second": self.current_second,
                "last_snr": self.last_snr,
                "last_max_amp": self.last_max_amp,
                "last_noise": self.last_noise,
                "last_threshold": self.last_adaptive_threshold,
            }

    def apply_time_shift(self, shift_sec):
        with self.sync_lock:
            if self.minute_start_time is not None:
                self.minute_start_time += shift_sec

            if self.last_impulse_time is not None:
                self.last_impulse_time += shift_sec

    def reset(self):
        with self.sync_lock:
            self.last_impulse_time = None
            self.impulse_count = 0
            self.minute_start_time = None
            self.is_synced = False
            self.current_second = 0


# ------------------------------------------------------------
# Синхронизатор системного времени
# ------------------------------------------------------------
class TimeSyncController:
    def __init__(self, decoder, log_func, use_ntp=True, ntp_servers=None):
        self.decoder = decoder
        self.log_func = log_func
        self.use_ntp = use_ntp
        self.ntp_servers = ntp_servers or ["0.ru.pool.ntp.org"]

        self.running = False
        self.thread = None

        self.correction_count = 0
        self.last_correction_time = 0
        self.sync_quality = 0

        self.last_big_offset_log = 0.0
        self.last_small_offset_log = 0.0

    def start(self):
        if self.running:
            return

        self.running = True
        self.thread = threading.Thread(target=self._sync_loop, daemon=True)
        self.thread.start()

        self.log_func("[СИНХРОНИЗАТОР] Запущен")

    def stop(self):
        self.running = False

        if self.thread:
            self.thread.join(timeout=1)
            self.thread = None

        self.log_func("[СИНХРОНИЗАТОР] Остановлен")

    def _sync_loop(self):
        last_status_log = 0
        minute_start_known = False

        while self.running:
            now = time.time()

            if now - last_status_log > 5:
                status = self.decoder.get_status()
                mode = RWMSchedule.get_mode()
                mode_desc = RWMSchedule.get_mode_description(mode)
                sync_possible = RWMSchedule.is_sync_possible(mode)

                log_msg = (
                    f"[СТАТУС] Режим RWM: {mode} ({mode_desc}) | "
                    f"Синхронизация: {'ДА' if status['is_synced'] else 'НЕТ'} | "
                    f"Импульсов: {status['impulse_count']} | "
                    f"Секунда: {status['current_second']:02d} | "
                    f"Пригоден: {'ДА' if sync_possible else 'НЕТ'} | "
                    f"Сигнал: {status['last_max_amp']:.3f} | "
                    f"Шум: {status['last_noise']:.3f} | "
                    f"Порог: {status['last_threshold']:.4f} | "
                    f"SNR: {status['last_snr']:.1f}"
                )

                self.log_func(log_msg)
                last_status_log = now

            mode = RWMSchedule.get_mode()

            if not RWMSchedule.is_sync_possible(mode):
                time.sleep(1)
                continue

            minute_start = self.decoder.get_minute_start()

            if minute_start is not None:
                if not minute_start_known:
                    minute_start_known = True
                    self.log_func(
                        f"[СИНХРОНИЗАТОР] Начало минуты получено: "
                        f"{datetime.fromtimestamp(minute_start).strftime('%H:%M:%S')}"
                    )

                # Минутный маркер должен быть кратен 60 секундам UTC.
                nearest_minute = math.floor(minute_start / 60.0 + 0.5) * 60.0
                offset = nearest_minute - minute_start

                self.sync_quality = max(0.0, min(100.0, 100.0 - abs(offset) * 100.0))

                if abs(offset) > 30:
                    if now - self.last_big_offset_log > 30:
                        self.last_big_offset_log = now
                        self.log_func(
                            f"[СИНХРОНИЗАТОР] Слишком большое расхождение {offset:.3f} с. "
                            f"Требуется NTP или ручная коррекция."
                        )

                elif abs(offset) > 0.050:
                    self.log_func(
                        f"[СИНХРОНИЗАТОР] Расхождение: {offset * 1000:+.1f} мс | "
                        f"Качество: {self.sync_quality:.0f}% | "
                        f"Коррекция..."
                    )

                    if self._adjust_system_time(offset):
                        self.decoder.apply_time_shift(offset)
                        self.last_correction_time = time.time()
                        self.correction_count += 1

                        self.log_func(
                            f"[СИНХРОНИЗАТОР] ✓ Коррекция выполнена "
                            f"(всего: {self.correction_count})"
                        )

                else:
                    if abs(offset) > 0.001 and now - self.last_small_offset_log > 30:
                        self.last_small_offset_log = now
                        self.log_func(
                            f"[СИНХРОНИЗАТОР] Расхождение: {offset * 1000:+.1f} мс — "
                            f"в пределах допуска, коррекция не требуется"
                        )

            else:
                if minute_start_known:
                    minute_start_known = False
                    self.log_func("[СИНХРОНИЗАТОР] Потеря сигнала начала минуты")

            if self.use_ntp and (time.time() - self.last_correction_time) > 300:
                ntp_offset = self._query_ntp()

                if ntp_offset is not None:
                    self.log_func(f"[NTP] Расхождение с NTP: {ntp_offset * 1000:+.1f} мс")

                    if abs(ntp_offset) > 0.1:
                        if self._adjust_system_time(ntp_offset):
                            self.decoder.apply_time_shift(ntp_offset)
                            self.last_correction_time = time.time()
                            self.correction_count += 1

                            self.log_func(
                                f"[NTP] ✓ Коррекция по NTP выполнена "
                                f"(всего: {self.correction_count})"
                            )

            time.sleep(0.5)

    def _adjust_system_time(self, offset_sec):
        try:
            new_time = datetime.now(timezone.utc) + timedelta(seconds=offset_sec)

            # Сначала пробуем win32api, если доступен.
            try:
                import win32api

                win32api.SetSystemTime(
                    new_time.year,
                    new_time.month,
                    new_time.isoweekday() % 7,
                    new_time.day,
                    new_time.hour,
                    new_time.minute,
                    new_time.second,
                    new_time.microsecond // 1000
                )

                self.log_func(
                    f"[СИСТЕМА] Время установлено: "
                    f"{new_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} UTC"
                )

                return True

            except Exception:
                pass

            # Запасной вариант через ctypes.
            import ctypes

            if not hasattr(ctypes, "windll"):
                raise OSError("Установка системного времени доступна только в Windows")

            class SYSTEMTIME(ctypes.Structure):
                _fields_ = [
                    ("wYear", ctypes.c_ushort),
                    ("wMonth", ctypes.c_ushort),
                    ("wDayOfWeek", ctypes.c_ushort),
                    ("wDay", ctypes.c_ushort),
                    ("wHour", ctypes.c_ushort),
                    ("wMinute", ctypes.c_ushort),
                    ("wSecond", ctypes.c_ushort),
                    ("wMilliseconds", ctypes.c_ushort),
                ]

            st = SYSTEMTIME()

            st.wYear = new_time.year
            st.wMonth = new_time.month
            st.wDayOfWeek = new_time.isoweekday() % 7
            st.wDay = new_time.day
            st.wHour = new_time.hour
            st.wMinute = new_time.minute
            st.wSecond = new_time.second
            st.wMilliseconds = new_time.microsecond // 1000

            if ctypes.windll.kernel32.SetSystemTime(ctypes.byref(st)) == 0:
                raise ctypes.WinError()

            self.log_func(
                f"[СИСТЕМА] Время установлено: "
                f"{new_time.strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]} UTC"
            )

            return True

        except Exception as e:
            self.log_func(f"[СИСТЕМА] Ошибка установки времени: {e}")
            return False

    def _query_ntp(self):
        NTP_PORT = 123
        NTP_EPOCH = 2208988800

        packet = bytearray(48)
        packet[0] = 0b00100011

        for server in self.ntp_servers:
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            sock.settimeout(2)

            data = None

            try:
                sock.sendto(packet, (server, NTP_PORT))
                data, _ = sock.recvfrom(1024)
            except Exception:
                continue
            finally:
                sock.close()

            if data is None or len(data) < 48:
                continue

            try:
                unpacked = struct.unpack("!12I", data[:48])
                t = float(unpacked[10]) + float(unpacked[11]) / (2 ** 32)
                ntp_time = t - NTP_EPOCH
                return ntp_time - time.time()
            except Exception:
                continue

        return None


# ------------------------------------------------------------
# GUI (главное окно)
# ------------------------------------------------------------
class TimeSyncApp:
    RWM_FREQS = {
        "4996": 4996000,
        "9996": 9996000,
        "14996": 14996000,
    }

    WATERFALL_UPDATE_MS = 150
    WATERFALL_FREQ_MAX = 3000
    WATERFALL_BINS = 128
    FFT_SIZE = 1024

    def __init__(self, root):
        self.root = root
        self.root.title("Радиосинхронизация времени (ГОСТ 8.323-2016) — RWM")

        try:
            self.root.state("zoomed")
        except Exception:
            pass

        try:
            self.root.minsize(1080, 640)
        except Exception:
            pass

        self.log_queue = queue.Queue()
        self.running = True

        self.cat_mode = CONFIG.get("cat_mode", "Direct CAT")
        self.audio_in_device = CONFIG.get("audio_input", "")
        self.rx_freq = CONFIG.get("rx_freq", 4996000)

        self.audio_input = None
        self.direct_cat = DirectCAT()
        self.sniffer = None

        self.decoder = None
        self.sync_controller = None

        self.main_pane = None

        self.waterfall_data = None
        self.waterfall_idx = 0
        self.waterfall_max_lines = 220
        self.last_decoder_status = ""

        self._create_vars()
        self._build_menu()
        self._build_main_ui()
        self._build_settings_ui()

        self._apply_audio_settings(startup=True)
        self._apply_cat_settings(startup=True)

        self.decoder = RadioTimeDecoder(
            threshold=CONFIG.get("manual_threshold", 0.3),
            log_func=self._log,
            auto_threshold=CONFIG.get("auto_threshold", True)
        )

        self.sync_controller = TimeSyncController(
            self.decoder,
            self._log,
            use_ntp=CONFIG.get("use_ntp_fallback", True),
            ntp_servers=CONFIG.get("ntp_servers", ["0.ru.pool.ntp.org"])
        )

        self.root.after(100, self._process_log_queue)
        self.root.after(150, self._update_vu)
        self.root.after(self.WATERFALL_UPDATE_MS, self._update_waterfall)
        self.root.after(500, self._update_status)
        self.root.after(2000, self._log_rms)
        self.root.after(500, self._process_audio)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self._log("=" * 60)
        self._log("ПРОГРАММА ЗАПУЩЕНА")
        self._log(f"Частота приёма: {self.rx_freq} Гц")
        self._log("Расписание RWM: цикл 30 минут, секундные импульсы в 10-20 и 40-50 минутах")
        self._log("Нажмите 'Старт' для начала синхронизации")
        self._log("=" * 60)

        if not AUDIO_AVAILABLE:
            self._log("[ОШИБКА] numpy/pyaudio недоступны — аудиообработка отключена", "ERROR")

        if not MATPLOTLIB_AVAILABLE:
            self._log("[ВНИИМАНИЕ] matplotlib недоступен — водопад отключен", "WARNING")

    # --------------------------------------------------------
    # Непрерывная обработка аудио
    # --------------------------------------------------------
    def _process_audio(self):
        if not self.running:
            return

        if (
            AUDIO_AVAILABLE and
            self.audio_input and
            self.audio_input.running and
            self.decoder
        ):
            buf = self.audio_input.get_buffer(1.0)

            if buf is not None and len(buf) > 100:
                noise = self.audio_input.get_noise_floor()

                if noise <= 0:
                    noise = None

                mode = RWMSchedule.get_mode()
                result = self.decoder.process_buffer(buf, noise_floor=noise, mode=mode)

                if result and result.get("status") not in ("no_signal", "no_pulses", "no_valid_pulses"):
                    status = result.get("status")

                    if status != self.last_decoder_status:
                        self.last_decoder_status = status

                        if status == "mode_not_sync":
                            self._log(f"[ДЕКОДЕР] Режим {result.get('mode')} не содержит импульсов")

        self.root.after(500, self._process_audio)

    # --------------------------------------------------------
    # GUI построение
    # --------------------------------------------------------
    def _create_vars(self):
        self.cat_mode_var = tk.StringVar(value=CONFIG.get("cat_mode", "Direct CAT"))
        self.cat_com_var = tk.StringVar(value=CONFIG.get("cat_com_port", ""))
        self.cat_baud_var = tk.StringVar(value=str(CONFIG.get("cat_baudrate", 19200)))

        self.audio_input_var = tk.StringVar(value=CONFIG.get("audio_input", ""))
        self.agc_enabled = tk.BooleanVar(value=CONFIG.get("agc_enabled", True))

        self.freq_var = tk.StringVar(value=str(CONFIG.get("rx_freq", 4996000)))

        self.auto_threshold_var = tk.BooleanVar(value=CONFIG.get("auto_threshold", True))
        self.manual_threshold_var = tk.DoubleVar(value=CONFIG.get("manual_threshold", 0.3))

        self.pulse_center_var = tk.StringVar(value=str(CONFIG.get("pulse_center_hz", 2000)))
        self.pulse_band_var = tk.StringVar(value=str(CONFIG.get("pulse_band_hz", 300)))

    def _build_menu(self):
        menubar = tk.Menu(self.root)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Настройки", command=self.open_settings)

        menubar.add_cascade(label="Настройки", menu=settings_menu)
        self.root.config(menu=menubar)

    def _waterfall_time_extent(self):
        return max(1.0, self.waterfall_max_lines * self.WATERFALL_UPDATE_MS / 1000.0)

    def _build_main_ui(self):
        # ------------------------------------------------------------
        # Верхняя панель управления (кнопки всегда видны)
        # ------------------------------------------------------------
        ctrl = tk.LabelFrame(self.root, text="Управление", padx=8, pady=6)
        ctrl.pack(fill="x", padx=6, pady=(6, 3))

        ctrl.grid_columnconfigure(7, weight=1)

        self.start_btn = tk.Button(
            ctrl,
            text="Старт",
            command=self.start_sync,
            bg="green",
            fg="white",
            width=10
        )
        self.start_btn.grid(row=0, column=0, padx=4, pady=4, sticky="w")

        self.stop_btn = tk.Button(
            ctrl,
            text="Стоп",
            command=self.stop_sync,
            bg="red",
            fg="white",
            width=10
        )
        self.stop_btn.grid(row=0, column=1, padx=4, sticky="w")

        tk.Button(
            ctrl,
            text="Тест аудио",
            command=self.test_audio,
            bg="orange"
        ).grid(row=0, column=2, padx=10, sticky="w")

        self.status_label = tk.Label(ctrl, text="Статус: Остановлен", fg="gray")
        self.status_label.grid(row=0, column=7, padx=8, sticky="e")

        self.correction_label = tk.Label(ctrl, text="Коррекций: 0")
        self.correction_label.grid(row=0, column=8, padx=8, sticky="e")

        tk.Label(ctrl, text="Частота (Гц):").grid(row=1, column=0, sticky="e", padx=(4, 2), pady=(0, 6))

        self.freq_entry = tk.Entry(ctrl, textvariable=self.freq_var, width=10)
        self.freq_entry.grid(row=1, column=1, sticky="w", padx=2)

        tk.Button(ctrl, text="Установить", command=self.set_frequency).grid(row=1, column=2, sticky="w", padx=2)

        tk.Label(ctrl, text="RWM:").grid(row=1, column=3, sticky="e", padx=(12, 2))

        col = 4
        for label, freq_hz in self.RWM_FREQS.items():
            btn = tk.Button(
                ctrl,
                text=f"{label} кГц",
                width=8,
                command=lambda f=freq_hz: self.set_rwm_frequency(f)
            )
            btn.grid(row=1, column=col, padx=2, sticky="w")
            col += 1

        # ------------------------------------------------------------
        # Нижняя раздвигаемая панель: лог слева, водопад справа
        # ------------------------------------------------------------
        self.main_pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill="both", expand=True, padx=6, pady=3)

        left_frame = tk.Frame(self.main_pane)
        self.main_pane.add(left_frame, width=650, minsize=430, stretch="always")

        log_frame = tk.LabelFrame(left_frame, text="Лог", padx=6, pady=6)
        log_frame.pack(fill="both", expand=True, padx=0, pady=3)

        self.log_text = tk.Text(
            log_frame,
            height=14,
            state="disabled",
            font=("Consolas", 9),
            bg="#000",
            fg="white",
            insertbackground="white"
        )
        self.log_text.pack(fill="both", expand=True)

        right_frame = tk.Frame(self.main_pane)
        self.main_pane.add(right_frame, width=430, minsize=330, stretch="always")

        meter_frame = tk.LabelFrame(right_frame, text="Входной уровень", padx=4, pady=4)
        meter_frame.pack(fill="x", pady=3)

        self.vu_progress = ttk.Progressbar(
            meter_frame,
            orient=tk.HORIZONTAL,
            length=200,
            mode="determinate",
            maximum=100
        )
        self.vu_progress.pack(pady=5, padx=5)

        self.vu_label = tk.Label(meter_frame, text="0 %")
        self.vu_label.pack()

        waterfall_frame = tk.LabelFrame(right_frame, text="Водопад (спектрограмма)", padx=4, pady=4)
        waterfall_frame.pack(fill="both", expand=True, pady=3)

        if MATPLOTLIB_AVAILABLE and np is not None:
            self.fig = Figure(figsize=(4, 3), dpi=80)
            self.ax = self.fig.add_subplot(111)

            # Исправленные подписи: X = частота, Y = время.
            self.ax.set_title("Спектр")
            self.ax.set_xlabel("Частота (Гц)")
            self.ax.set_ylabel("Время (с)")

            self.waterfall_data = np.zeros((self.waterfall_max_lines, self.WATERFALL_BINS))

            extent = [
                0,
                self.WATERFALL_FREQ_MAX,
                0,
                self._waterfall_time_extent()
            ]

            self.waterfall_img = self.ax.imshow(
                self.waterfall_data,
                aspect="auto",
                origin="lower",
                cmap="viridis",
                extent=extent
            )

            self.ax.set_xlim(0, self.WATERFALL_FREQ_MAX)
            self.ax.set_ylim(0, self._waterfall_time_extent())

            self.canvas = FigureCanvasTkAgg(self.fig, master=waterfall_frame)
            self.canvas.get_tk_widget().pack(fill="both", expand=True)

        else:
            self.fig = None
            self.ax = None
            self.waterfall_img = None
            self.canvas = None
            self.waterfall_data = None

            tk.Label(
                waterfall_frame,
                text="matplotlib/numpy недоступны — водопад отключен",
                fg="red",
                wraplength=300,
                justify="center"
            ).pack(fill="both", expand=True)

        # Стартовая позиция разделителя, чтобы лог и водопад были распределены нормально.
        self.root.after(150, self._set_initial_sash)

    def _set_initial_sash(self):
        try:
            self.root.update_idletasks()

            if self.main_pane is None:
                return

            width = self.main_pane.winfo_width()

            if width < 100:
                width = self.root.winfo_screenwidth() - 100

            pos = int(width * 0.62)
            self.main_pane.sash_place(0, pos, 1)

        except Exception:
            pass

    # --------------------------------------------------------
    # Настройки (окно)
    # --------------------------------------------------------
    def _build_settings_ui(self):
        self.settings_window = tk.Toplevel(self.root)
        self.settings_window.title("Настройки")
        self.settings_window.withdraw()
        self.settings_window.protocol("WM_DELETE_WINDOW", self._hide_settings)

        notebook = ttk.Notebook(self.settings_window)
        notebook.pack(fill="both", expand=True, padx=6, pady=6)

        audio_tab = ttk.Frame(notebook)
        notebook.add(audio_tab, text="Аудио")

        tk.Label(audio_tab, text="Входное устройство:").grid(row=0, column=0, sticky="e", padx=6, pady=6)

        self.audio_input_cb = ttk.Combobox(
            audio_tab,
            textvariable=self.audio_input_var,
            width=40,
            state="readonly"
        )
        self.audio_input_cb.grid(row=0, column=1, padx=2)

        tk.Button(audio_tab, text="Обновить", command=self._refresh_audio).grid(row=0, column=2, padx=4)
        tk.Button(audio_tab, text="Применить", command=self._apply_audio_settings).grid(row=0, column=3, padx=4)

        tk.Checkbutton(
            audio_tab,
            text="AGC (автоусиление)",
            variable=self.agc_enabled,
            command=self._save_agc
        ).grid(row=1, column=0, columnspan=4, sticky="w", padx=6, pady=6)

        cat_tab = ttk.Frame(notebook)
        notebook.add(cat_tab, text="CAT")

        tk.Label(cat_tab, text="Режим:").grid(row=0, column=0, sticky="e", padx=6, pady=6)

        ttk.Combobox(
            cat_tab,
            textvariable=self.cat_mode_var,
            values=["Direct CAT", "Sniffer"],
            width=15,
            state="readonly"
        ).grid(row=0, column=1, padx=2)

        tk.Label(cat_tab, text="COM:").grid(row=0, column=2, sticky="e", padx=6)

        self.cat_com_cb = ttk.Combobox(
            cat_tab,
            textvariable=self.cat_com_var,
            width=10,
            state="readonly"
        )
        self.cat_com_cb.grid(row=0, column=3, padx=2)

        tk.Label(cat_tab, text="Baud:").grid(row=0, column=4, sticky="e", padx=6)

        ttk.Combobox(
            cat_tab,
            textvariable=self.cat_baud_var,
            values=[str(x) for x in (9600, 19200, 38400, 57600, 115200)],
            width=8,
            state="readonly"
        ).grid(row=0, column=5, padx=2)

        tk.Button(cat_tab, text="Обновить", command=self._refresh_cat_ports).grid(row=0, column=6, padx=4)
        tk.Button(cat_tab, text="Применить", command=self._apply_cat_settings).grid(row=0, column=7, padx=4)

        self.cat_status_label = tk.Label(cat_tab, text="Не подключено", fg="gray")
        self.cat_status_label.grid(row=1, column=0, columnspan=8, sticky="w", padx=6, pady=6)

        signal_tab = ttk.Frame(notebook)
        notebook.add(signal_tab, text="Сигнал")

        tk.Label(signal_tab, text="Порог детекции (0..1):").grid(row=0, column=0, sticky="e", padx=6, pady=6)

        self.threshold_scale = tk.Scale(
            signal_tab,
            from_=0.05,
            to=0.8,
            resolution=0.05,
            orient="horizontal",
            variable=self.manual_threshold_var,
            length=200
        )
        self.threshold_scale.grid(row=0, column=1, padx=2)

        tk.Button(signal_tab, text="Применить", command=self._apply_threshold).grid(row=0, column=2, padx=4)

        tk.Checkbutton(
            signal_tab,
            text="Автопорог",
            variable=self.auto_threshold_var,
            command=self._toggle_auto_threshold
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=6)

        self.ntp_var = tk.BooleanVar(value=CONFIG.get("use_ntp_fallback", True))

        tk.Checkbutton(
            signal_tab,
            text="Резерв NTP",
            variable=self.ntp_var,
            command=self._save_ntp_option
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=6)

        tk.Label(signal_tab, text="Центр фильтра (Гц):").grid(row=3, column=0, sticky="e", padx=6, pady=6)
        tk.Entry(signal_tab, textvariable=self.pulse_center_var, width=10).grid(row=3, column=1, sticky="w", padx=2)

        tk.Label(signal_tab, text="Полоса фильтра (Гц):").grid(row=4, column=0, sticky="e", padx=6, pady=6)
        tk.Entry(signal_tab, textvariable=self.pulse_band_var, width=10).grid(row=4, column=1, sticky="w", padx=2)

        tk.Button(
            signal_tab,
            text="Применить фильтр",
            command=self._apply_filter_settings
        ).grid(row=5, column=1, sticky="w", padx=2, pady=6)

        self._refresh_audio()
        self._refresh_cat_ports()
        self._update_threshold_ui()

    def _update_threshold_ui(self):
        auto = self.auto_threshold_var.get()
        self.threshold_scale.config(state="disabled" if auto else "normal")

    def _toggle_auto_threshold(self):
        auto = self.auto_threshold_var.get()
        self._update_threshold_ui()

        CONFIG["auto_threshold"] = auto
        save_config(CONFIG)

        if self.decoder:
            self.decoder.auto_threshold = auto

        self._log(f"[ДЕКОДЕР] Автопорог {'включён' if auto else 'выключен'}")

    # --------------------------------------------------------
    # Логирование
    # --------------------------------------------------------
    def _log(self, msg, level="INFO"):
        ts = datetime.now(timezone.utc).strftime("%H:%M:%S.%f")[:-3]
        line = f"[{ts}] {msg}"

        self.log_queue.put(line)
        print(line)

    def _log_rms(self):
        if not self.running:
            return

        if self.audio_input and self.audio_input.running:
            rms = self.audio_input.get_raw_rms()
            gain = self.audio_input.agc_gain
            self._log(f"[АУДИО] RMS: {rms:.0f}  Усиление: {gain:.1f}x")

        self.root.after(3000, self._log_rms)

    def _process_log_queue(self):
        if not self.running:
            return

        try:
            while True:
                line = self.log_queue.get_nowait()

                self.log_text.configure(state="normal")
                self.log_text.insert("end", line + "\n")
                self.log_text.see("end")

                try:
                    if int(self.log_text.index("end-1c").split(".")[0]) > 5000:
                        self.log_text.delete("1.0", "100.0")
                except Exception:
                    pass

                self.log_text.configure(state="disabled")

        except queue.Empty:
            pass

        self.root.after(100, self._process_log_queue)

    def _update_status(self):
        if not self.running:
            return

        mode = RWMSchedule.get_mode()

        if self.sync_controller and self.sync_controller.running:
            status = self.decoder.get_status()
            sync_str = "СИНХР" if status["is_synced"] else "поиск..."

            rms_val = 0.0
            if self.audio_input and self.audio_input.running:
                rms_val = self.audio_input.get_raw_rms()

            self.status_label.config(
                text=(
                    f"Статус: {sync_str} | Режим: {mode} | "
                    f"Имп: {status['impulse_count']} | SNR: {status['last_snr']:.1f} | "
                    f"RMS: {rms_val:.0f}"
                ),
                fg="green" if status["is_synced"] else "orange"
            )

            self.correction_label.config(text=f"Коррекций: {self.sync_controller.correction_count}")

        else:
            self.status_label.config(text=f"Статус: Остановлен | Режим: {mode}", fg="gray")

        self.root.after(500, self._update_status)

    # --------------------------------------------------------
    # Тест аудио
    # --------------------------------------------------------
    def test_audio(self):
        if not AUDIO_AVAILABLE:
            self._log("[ТЕСТ] numpy/pyaudio недоступны", "ERROR")
            return

        if self.audio_input is None or not self.audio_input.running:
            self._log("[ТЕСТ] Аудио не запущено", "WARNING")
            return

        buf = self.audio_input.get_buffer(2.0)

        if buf is None or len(buf) < 100:
            self._log("[ТЕСТ] Нет данных — проверьте устройство ввода", "ERROR")
            return

        rms = float(np.sqrt(np.mean(np.square(buf.astype(np.float64)))))
        max_val = int(np.max(np.abs(buf)))

        self._log(f"[ТЕСТ] RMS: {rms:.0f},  Макс: {max_val},  Длина: {len(buf)}")

        if rms < 10:
            self._log("[ТЕСТ] ⚠️ Очень слабый сигнал — проверьте уровень громкости", "WARNING")
        else:
            self._log("[ТЕСТ] ✅ Сигнал есть, уровень достаточный", "INFO")

    # --------------------------------------------------------
    # VU-метр
    # --------------------------------------------------------
    def _update_vu(self):
        if not self.running:
            return

        level = 0.0

        if self.audio_input and self.audio_input.running:
            level = self.audio_input.get_rms()

        percent = int(level * 100)

        self.vu_progress["value"] = percent
        self.vu_label.config(text=f"{percent} %")

        if percent > 80:
            self.vu_progress["style"] = "red.Horizontal.TProgressbar"
        elif percent > 50:
            self.vu_progress["style"] = "yellow.Horizontal.TProgressbar"
        else:
            self.vu_progress["style"] = "green.Horizontal.TProgressbar"

        self.root.after(100, self._update_vu)

    # --------------------------------------------------------
    # Водопад
    # --------------------------------------------------------
    def _update_waterfall(self):
        if not self.running:
            return

        if (
            MATPLOTLIB_AVAILABLE and
            np is not None and
            self.audio_input and
            self.audio_input.running and
            self.waterfall_data is not None and
            self.waterfall_img is not None
        ):
            buf = self.audio_input.get_buffer(0.25)

            if buf is not None and len(buf) >= self.FFT_SIZE:
                segment = buf[-self.FFT_SIZE:].astype(np.float64)
                window = np.hanning(self.FFT_SIZE)
                segment *= window

                fft = np.fft.rfft(segment)
                mag = np.abs(fft)

                freqs = np.fft.rfftfreq(self.FFT_SIZE, 1 / AudioInput.SAMPLE_RATE)

                idx = int(np.searchsorted(freqs, self.WATERFALL_FREQ_MAX))
                idx = max(2, idx)

                mag = mag[:idx]

                if mag.size > 0:
                    mag_db = 20 * np.log10(mag + 1e-6)
                    new_mag = resize_1d(mag_db, self.WATERFALL_BINS)

                    if new_mag is not None and len(new_mag) == self.WATERFALL_BINS:
                        self.waterfall_data = np.roll(self.waterfall_data, -1, axis=0)
                        self.waterfall_data[-1, :] = new_mag

                        self.waterfall_img.set_data(self.waterfall_data)

                        vmin = float(np.percentile(self.waterfall_data, 5))
                        vmax = float(np.percentile(self.waterfall_data, 95))

                        if vmax <= vmin:
                            vmax = vmin + 1e-6

                        self.waterfall_img.set_clim(vmin=vmin, vmax=vmax)
                        self.canvas.draw_idle()

        self.root.after(self.WATERFALL_UPDATE_MS, self._update_waterfall)

    # --------------------------------------------------------
    # Управление синхронизацией
    # --------------------------------------------------------
    def start_sync(self):
        if not AUDIO_AVAILABLE:
            self._log("[ОШИБКА] PyAudio/numpy не доступны", "ERROR")
            return

        if self.audio_input is None or not self.audio_input.running:
            self._log("[АУДИО] Не запущено, попытка запуска...")

            if not self._apply_audio_settings():
                self._log("[ОШИБКА] Не удалось запустить аудио", "ERROR")
                return

        mode = RWMSchedule.get_mode()

        if not RWMSchedule.is_sync_possible(mode):
            self._log(
                f"[ПРЕДУПРЕЖДЕНИЕ] Текущий режим RWM: {mode} — "
                f"синхронизация невозможна. Ожидайте перехода в SECOND_PULSES "
                f"(10-20 или 40-50 минуты)"
            )

        self.decoder.reset()
        self.sync_controller.start()

        self._log("[СИНХРОНИЗАЦИЯ] Запущена")

    def stop_sync(self):
        if self.sync_controller:
            self.sync_controller.stop()

        self._log("[СИНХРОНИЗАЦИЯ] Остановлена")

    # --------------------------------------------------------
    # Управление частотой через CAT
    # --------------------------------------------------------
    def set_frequency(self, freq_hz=None):
        if freq_hz is None:
            try:
                freq_hz = int(self.freq_var.get())
            except ValueError:
                messagebox.showerror("Ошибка", "Введите частоту в Гц")
                return

        self.rx_freq = freq_hz

        CONFIG["rx_freq"] = freq_hz
        save_config(CONFIG)

        self.freq_var.set(str(freq_hz))
        self._send_freq(freq_hz)

        self._log(f"[ЧАСТОТА] Установлена: {freq_hz} Гц")

    def set_rwm_frequency(self, rwm_freq_hz):
        target_freq = rwm_freq_hz - 2000
        self.set_frequency(target_freq)

        self.root.after(100, self._send_u_d)

        self._log(
            f"[RWM] Переключение на {rwm_freq_hz // 1000} кГц "
            f"(частота приёма: {target_freq} Гц)"
        )

    def _send_freq(self, freq_hz):
        frame = self._freq_to_civ(freq_hz)

        if self.cat_mode == "Direct CAT":
            self.direct_cat.write(frame)
        elif self.cat_mode == "Sniffer" and self.sniffer:
            self.sniffer.write_real(frame)

    def _send_u_d(self):
        u_d_frame = bytes([0xFE, 0xFE, 0x88, 0xE0, 0x26, 0x00, 0x01, 0x01, 0x01, 0xFD])

        sent = False

        if self.cat_mode == "Direct CAT":
            sent = self.direct_cat.write(u_d_frame)
            if sent:
                self._log("[CAT] U-D отправлен (Direct CAT)")

        elif self.cat_mode == "Sniffer" and self.sniffer:
            sent = self.sniffer.write_real(u_d_frame)
            if sent:
                self._log("[CAT] U-D отправлен (Sniffer)")

        if not sent:
            self._log("[CAT] CAT недоступен", "WARNING")

    def _freq_to_civ(self, hz):
        s = f"{int(hz):010d}"
        bcd = bytes(int(s[i:i + 2], 16) for i in range(8, -2, -2))
        return bytes([0xFE, 0xFE, 0x88, 0xE0, 0x05]) + bcd + bytes([0xFD])

    # --------------------------------------------------------
    # Настройки (реализация)
    # --------------------------------------------------------
    def open_settings(self):
        self.settings_window.deiconify()
        self.settings_window.lift()

    def _hide_settings(self):
        self.settings_window.withdraw()

    def _refresh_audio(self):
        if pyaudio is None:
            return

        try:
            p = pyaudio.PyAudio()
            inputs = []

            for i in range(p.get_device_count()):
                info = p.get_device_info_by_index(i)

                if info.get("maxInputChannels", 0) > 0:
                    inputs.append(info.get("name", f"Device {i}"))

            p.terminate()

            self.audio_input_cb["values"] = inputs

            if inputs and not self.audio_input_var.get():
                self.audio_input_var.set(inputs[0])

        except Exception:
            pass

    def _apply_audio_settings(self, startup=False):
        dev = self.audio_input_var.get().strip()

        CONFIG["audio_input"] = dev
        save_config(CONFIG)

        if self.audio_input:
            self.audio_input.close()
            self.audio_input = None

        if not dev:
            self._log("[АУДИО] Устройство не выбрано", "WARNING")
            return False

        if not AUDIO_AVAILABLE:
            self._log("[ОШИБКА] numpy/pyaudio недоступны", "ERROR")
            return False

        self.audio_input = AudioInput(dev, agc_enabled=self.agc_enabled.get())

        if self.audio_input.start():
            self._log(f"[АУДИО] Запущено: {dev}")
            return True

        self._log(f"[ОШИБКА] Не удалось запустить аудио на {dev}", "ERROR")
        self.audio_input = None
        return False

    def _save_agc(self):
        CONFIG["agc_enabled"] = self.agc_enabled.get()
        save_config(CONFIG)

        if self.audio_input:
            self.audio_input.agc_enabled = self.agc_enabled.get()

        self._log(f"[AGC] {'Включён' if self.agc_enabled.get() else 'Выключен'}")

    def _refresh_cat_ports(self):
        ports = []

        if SERIAL_AVAILABLE:
            try:
                ports = sorted([p.device for p in serial.tools.list_ports.comports()])
            except Exception:
                pass

        self.cat_com_cb["values"] = ports

        if ports and not self.cat_com_var.get():
            self.cat_com_var.set(ports[0])

    def _apply_cat_settings(self, startup=False):
        mode = self.cat_mode_var.get().strip()
        port = self.cat_com_var.get().strip()
        baud = int(self.cat_baud_var.get() or 19200)

        CONFIG["cat_mode"] = mode
        CONFIG["cat_com_port"] = port
        CONFIG["cat_baudrate"] = baud

        save_config(CONFIG)

        self.cat_mode = mode

        if self.sniffer:
            self.sniffer.stop()
            self.sniffer = None

        self.direct_cat.close()

        if mode == "Direct CAT":
            if not port:
                self.cat_status_label.config(text="Не выбран COM", fg="gray")
                self._log("[CAT] Direct CAT выбран, но COM пуст", "WARNING")
                return

            if self.direct_cat.configure(port, baud):
                self.cat_status_label.config(text=f"Direct {port}", fg="green")
                self._log(f"[CAT] Direct на {port} {baud} бод")
            else:
                self.cat_status_label.config(text="Ошибка открытия", fg="red")
                self._log(f"[ОШИБКА] CAT: не удалось открыть {port}", "ERROR")

        elif mode == "Sniffer":
            if not port:
                self.cat_status_label.config(text="Не выбран COM", fg="gray")
                self._log("[CAT] Sniffer выбран, но COM пуст", "WARNING")
                return

            if not SERIAL_AVAILABLE:
                self.cat_status_label.config(text="pyserial недоступен", fg="red")
                self._log("[ОШИБКА] pyserial недоступен", "ERROR")
                return

            try:
                virtual_port = CONFIG.get("cat_virtual_port", "COM9").strip()

                self.sniffer = SerialSniffer(port, virtual_port, baud, self._log)
                self.sniffer.start()

                self.cat_status_label.config(text=f"Sniffer {port}", fg="green")
                self._log(f"[CAT] Sniffer на {port} -> {virtual_port}, {baud} бод")

            except Exception as e:
                self.cat_status_label.config(text="Ошибка", fg="red")
                self._log(f"[ОШИБКА] CAT Sniffer: {e}", "ERROR")

    def _apply_threshold(self):
        manual = float(self.manual_threshold_var.get())

        CONFIG["manual_threshold"] = manual
        save_config(CONFIG)

        if self.decoder:
            self.decoder.manual_threshold = manual

        self._log(f"[ДЕКОДЕР] Ручной порог: {manual:.2f}")

    def _apply_filter_settings(self):
        try:
            center = float(self.pulse_center_var.get())
            band = float(self.pulse_band_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Введите числовые значения фильтра")
            return

        if center <= 0 or band <= 0:
            messagebox.showerror("Ошибка", "Частота и полоса должны быть больше нуля")
            return

        CONFIG["pulse_center_hz"] = center
        CONFIG["pulse_band_hz"] = band
        save_config(CONFIG)

        if self.decoder:
            self.decoder.filter_center_hz = center
            self.decoder.filter_band_hz = band

        self._log(f"[ФИЛЬТР] Центр: {center:.0f} Гц, полоса: {band:.0f} Гц")

    def _save_ntp_option(self):
        val = self.ntp_var.get()

        CONFIG["use_ntp_fallback"] = val
        save_config(CONFIG)

        if self.sync_controller:
            self.sync_controller.use_ntp = val

        self._log(f"[NTP] Резерв {'включён' if val else 'выключен'}")

    # --------------------------------------------------------
    # Закрытие
    # --------------------------------------------------------
    def on_closing(self):
        self.running = False

        if self.sync_controller:
            self.sync_controller.stop()

        if self.audio_input:
            self.audio_input.close()

        if self.direct_cat:
            self.direct_cat.close()

        if self.sniffer:
            self.sniffer.stop()

        self.root.destroy()


# ------------------------------------------------------------
# Запуск
# ------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()

    style = ttk.Style()

    try:
        style.theme_use("default")
        style.configure("green.Horizontal.TProgressbar", background="green")
        style.configure("yellow.Horizontal.TProgressbar", background="orange")
        style.configure("red.Horizontal.TProgressbar", background="red")
    except Exception:
        pass

    app = TimeSyncApp(root)
    root.mainloop()