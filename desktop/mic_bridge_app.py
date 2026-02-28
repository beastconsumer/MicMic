import audioop
import math
import os
import struct
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
import shutil
import socket
import subprocess
import threading
import tkinter as tk

import customtkinter as ctk

try:
    import sounddevice as sd
except ImportError:  # pragma: no cover
    sd = None

try:
    import comtypes
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, EDataFlow, ERole, IAudioMeterInformation, IMMEndpoint
except ImportError:  # pragma: no cover
    comtypes = None
    CLSCTX_ALL = None
    AudioUtilities = None
    IMMEndpoint = None
    EDataFlow = None
    ERole = None
    IAudioMeterInformation = None

try:
    import winreg
except ImportError:  # pragma: no cover
    winreg = None

APP_TITLE = "MICMIC"
APP_PACKAGE = "com.micmic.mobilemic"
ADB_PORT = 28282
SAMPLE_RATE = 48000
CHANNELS = 2
BLOCK_SIZE = 2048

ROOT_DIR = Path(__file__).resolve().parents[1]
APK_PATH = ROOT_DIR / "android" / "app" / "build" / "outputs" / "apk" / "debug" / "app-debug.apk"

TARGET_MIC_NAME = "MICMIC"
TARGET_MIC_LONG_NAME = "MICMIC Virtual Mic"

PREFERRED_OUTPUT_HINTS = (
    "Virtual Speakers for AudioRelay",
    "CABLE Input",
    "VB-Audio",
)
PREFERRED_CAPTURE_HINTS = (
    TARGET_MIC_NAME,
    "Virtual Mic for AudioRelay",
    "CABLE Output",
    "VB-Audio",
)

COLOR_BG = "#05060A"
COLOR_SURFACE = "#0E111A"
COLOR_CARD = "#131827"
COLOR_CARD_ALT = "#171F31"
COLOR_BORDER = "#283247"
COLOR_TEXT = "#F4F7FF"
COLOR_MUTED = "#8FA0BC"
COLOR_ACCENT = "#00B7A8"
COLOR_ACCENT_HOVER = "#00A093"
COLOR_OK = "#22C55E"
COLOR_WARN = "#F59E0B"
COLOR_ERR = "#EF4444"
COLOR_INFO = "#58D6FF"


class MicBridgeError(RuntimeError):
    pass


@dataclass
class OutputDevice:
    index: int
    name: str


@dataclass
class CaptureDevice:
    device_id: str
    name: str


@dataclass
class AdbDevice:
    serial: str
    state: str
    model: str


@contextmanager
def com_scope():
    if comtypes is None:
        yield
        return
    comtypes.CoInitialize()
    try:
        yield
    finally:
        comtypes.CoUninitialize()


def run_command(command, check=True, timeout=45):
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=timeout,
        shell=False,
    )
    if check and completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip() or "sem detalhes"
        raise MicBridgeError(f"Falha em {' '.join(command)}: {detail}")
    return completed


def resolve_adb_executable() -> str:
    env_candidate = os.environ.get("ADB_PATH", "").strip()
    if env_candidate:
        candidate = Path(env_candidate)
        if candidate.exists():
            return str(candidate)

    from_path = shutil.which("adb")
    if from_path:
        return from_path

    local_app_data = os.environ.get("LOCALAPPDATA", "")
    fallback = Path(local_app_data) / "Android" / "Sdk" / "platform-tools" / "adb.exe"
    if fallback.exists():
        return str(fallback)

    raise MicBridgeError("ADB nao encontrado. Instale Android Platform Tools.")


def parse_adb_devices(output: str) -> list[AdbDevice]:
    devices = []
    for line in output.splitlines()[1:]:
        line = line.strip()
        if not line:
            continue
        parts = line.split()
        serial = parts[0]
        state = parts[1] if len(parts) > 1 else "unknown"
        model = "-"
        for part in parts[2:]:
            if part.startswith("model:"):
                model = part.split(":", 1)[1]
        devices.append(AdbDevice(serial=serial, state=state, model=model))
    return devices


class AdbClient:
    def __init__(self):
        self.exe = resolve_adb_executable()

    def run(self, args, check=True, timeout=45):
        return run_command([self.exe, *args], check=check, timeout=timeout)

    def get_connected_device(self) -> AdbDevice:
        completed = self.run(["devices", "-l"], check=False)
        devices = parse_adb_devices(completed.stdout)
        for dev in devices:
            if dev.state == "device":
                return dev
        if any(dev.state == "unauthorized" for dev in devices):
            raise MicBridgeError("Celular nao autorizado. Aceite a chave RSA no celular.")
        if any(dev.state == "offline" for dev in devices):
            raise MicBridgeError("Celular offline no ADB. Reconecte o cabo USB.")
        raise MicBridgeError("Nenhum celular conectado no ADB.")

    def is_package_installed(self, package_name: str) -> bool:
        completed = self.run(["shell", "pm", "path", package_name], check=False)
        return completed.returncode == 0 and "package:" in completed.stdout

    def install_apk(self, apk_path: Path):
        if not apk_path.exists():
            raise MicBridgeError(f"APK nao encontrado em {apk_path}")
        self.run(["install", "-r", str(apk_path)], timeout=180)


def list_output_devices() -> list[OutputDevice]:
    if sd is None:
        raise MicBridgeError("Dependencia de audio ausente. Rode: python -m pip install -r requirements.txt")
    items = []
    for index, dev in enumerate(sd.query_devices()):
        if int(dev.get("max_output_channels", 0)) > 0:
            items.append(OutputDevice(index=index, name=str(dev.get("name", f"Output {index}"))))
    return items


def find_input_device_index_by_name(name_hint: str) -> int | None:
    if sd is None:
        return None
    hint = (name_hint or "").strip().lower()
    if not hint:
        return None

    best_index = None
    best_score = -10_000
    for index, dev in enumerate(sd.query_devices()):
        if int(dev.get("max_input_channels", 0)) <= 0:
            continue
        name = str(dev.get("name", ""))
        name_lower = name.lower().strip()
        score = 0
        if name_lower == hint:
            score += 200
        if hint in name_lower:
            score += 140
        if name_lower in hint:
            score += 40
        if "micmic" in hint and "micmic" in name_lower:
            score += 120
        if "cable output" in hint and "cable output" in name_lower:
            score += 90
        if "point" in name_lower:
            score -= 40
        if score > best_score:
            best_score = score
            best_index = index
    return best_index if best_score > 0 else None


def list_capture_devices() -> list[CaptureDevice]:
    if AudioUtilities is None or IMMEndpoint is None or EDataFlow is None:
        raise MicBridgeError("Dependencia pycaw/comtypes ausente.")

    captures = {}
    with com_scope():
        for dev in AudioUtilities.GetAllDevices():
            try:
                flow = dev._dev.QueryInterface(IMMEndpoint).GetDataFlow()
            except Exception:
                continue
            if flow != EDataFlow.eCapture.value:
                continue
            if not dev.FriendlyName:
                continue
            captures[dev.id] = CaptureDevice(device_id=dev.id, name=dev.FriendlyName)
    return sorted(captures.values(), key=lambda d: d.name.lower())


def choose_preferred(devices, hints):
    if not devices:
        return None
    for hint in hints:
        hint_lower = hint.lower()
        for dev in devices:
            if hint_lower in dev.name.lower():
                return dev
    return devices[0]


def set_default_capture_device(device: CaptureDevice):
    if AudioUtilities is None or ERole is None:
        raise MicBridgeError("pycaw indisponivel para definir microfone padrao.")
    with com_scope():
        AudioUtilities.SetDefaultDevice(
            device.device_id,
            roles=[ERole.eConsole, ERole.eMultimedia, ERole.eCommunications],
        )


def extract_endpoint_guid(device_id: str) -> str | None:
    match = re.search(r"\.\{([0-9a-fA-F-]{36})\}$", device_id)
    return match.group(1) if match else None


def try_rename_capture_device_to_micmic(device: CaptureDevice) -> tuple[bool, str]:
    if winreg is None:
        return False, "winreg indisponivel."

    guid = extract_endpoint_guid(device.device_id)
    if not guid:
        return False, "ID do endpoint invalido para renomeacao."

    reg_path = (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture"
        + f"\\{{{guid}}}\\Properties"
    )
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, reg_path, 0, winreg.KEY_SET_VALUE
        ) as key:
            winreg.SetValueEx(
                key,
                "{a45c254e-df1c-4efd-8020-67d146a850e0},2",
                0,
                winreg.REG_SZ,
                TARGET_MIC_NAME,
            )
            winreg.SetValueEx(
                key,
                "{b3f8fa53-0004-438e-9003-51a46e139bfc},6",
                0,
                winreg.REG_SZ,
                TARGET_MIC_LONG_NAME,
            )
        return True, "Renomeado para MICMIC no registro."
    except PermissionError:
        return False, "Sem permissao de administrador para renomear dispositivo."
    except OSError as exc:
        return False, f"Falha ao renomear no registro: {exc}"


def _find_pycaw_capture_endpoint_by_name(capture_name: str):
    if AudioUtilities is None or IMMEndpoint is None or EDataFlow is None:
        return None
    with com_scope():
        for dev in AudioUtilities.GetAllDevices():
            try:
                flow = dev._dev.QueryInterface(IMMEndpoint).GetDataFlow()
            except Exception:
                continue
            if flow != EDataFlow.eCapture.value:
                continue
            if (dev.FriendlyName or "").strip().lower() == capture_name.strip().lower():
                return dev
    return None


def measure_virtual_route_peak(output_device: OutputDevice, capture_device: CaptureDevice) -> float:
    if sd is None or IAudioMeterInformation is None or CLSCTX_ALL is None:
        return 0.0

    endpoint = _find_pycaw_capture_endpoint_by_name(capture_device.name)
    if endpoint is None:
        return 0.0

    with com_scope():
        iface = endpoint._dev.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)
        meter = iface.QueryInterface(IAudioMeterInformation)

        sample_rate = int(float(sd.query_devices(output_device.index).get("default_samplerate", 48000)))
        if sample_rate <= 0:
            sample_rate = 48000

        freq = 440.0
        step = (2.0 * math.pi * freq) / sample_rate
        phase = 0.0
        peak = 0.0

        with sd.RawOutputStream(
            samplerate=sample_rate,
            channels=2,
            dtype="int16",
            device=output_device.index,
            blocksize=1024,
        ) as out_stream:
            start = time.time()
            while time.time() - start < 1.5:
                buffer = bytearray()
                for _ in range(1024):
                    value = int(9000 * math.sin(phase))
                    phase += step
                    if phase > 2.0 * math.pi:
                        phase -= 2.0 * math.pi
                    buffer += struct.pack("<hh", value, value)
                out_stream.write(bytes(buffer))
                current_peak = float(meter.GetPeakValue())
                if current_peak > peak:
                    peak = current_peak
        return peak


def read_capture_peak(capture_device: CaptureDevice | None) -> float:
    if capture_device is None:
        return 0.0
    if IAudioMeterInformation is None or CLSCTX_ALL is None:
        return 0.0

    endpoint = _find_pycaw_capture_endpoint_by_name(capture_device.name)
    if endpoint is None:
        return 0.0

    with com_scope():
        iface = endpoint._dev.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)
        meter = iface.QueryInterface(IAudioMeterInformation)
        peak = float(meter.GetPeakValue())
    return max(0.0, min(1.0, peak))


class AudioRelay:
    def __init__(self, output_device_index: int, output_device_name: str, status_cb):
        self.output_device_index = output_device_index
        self.output_device_name = output_device_name
        self.status_cb = status_cb
        self._stop = threading.Event()
        self._thread = None
        self._server = None
        self._client = None
        self._stream = None

    def start(self):
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True, name="audio-relay")
        self._thread.start()

    def stop(self):
        self._stop.set()
        for sock in (self._client, self._server):
            if sock:
                try:
                    sock.shutdown(socket.SHUT_RDWR)
                except OSError:
                    pass
                try:
                    sock.close()
                except OSError:
                    pass
        if self._thread:
            self._thread.join(timeout=3)
            self._thread = None

    def _run(self):
        try:
            self._server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self._server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self._server.bind(("127.0.0.1", ADB_PORT))
            self._server.listen(1)
            self._server.settimeout(1.0)
            self.status_cb("Aguardando audio do celular...", "info")

            while not self._stop.is_set():
                try:
                    self._client, _addr = self._server.accept()
                    break
                except socket.timeout:
                    continue
            if self._stop.is_set() or not self._client:
                return

            self._client.settimeout(1.0)
            self._stream = sd.RawOutputStream(
                samplerate=SAMPLE_RATE,
                channels=CHANNELS,
                dtype="int16",
                device=self.output_device_index,
                blocksize=BLOCK_SIZE,
            )
            self._stream.start()
            self.status_cb(f"Stream ativo em {self.output_device_name}", "ok")

            while not self._stop.is_set():
                try:
                    chunk = self._client.recv(4096)
                except socket.timeout:
                    continue
                if not chunk:
                    break
                if len(chunk) % 2:
                    chunk = chunk[:-1]
                if chunk:
                    self._stream.write(audioop.tostereo(chunk, 2, 1, 1))
        except Exception as exc:  # pragma: no cover
            self.status_cb(f"Erro no relay: {exc}", "error")
        finally:
            if self._stream:
                try:
                    self._stream.stop()
                    self._stream.close()
                except Exception:
                    pass
            if self._client:
                try:
                    self._client.close()
                except OSError:
                    pass
            if self._server:
                try:
                    self._server.close()
                except OSError:
                    pass


class MicMicApp(ctk.CTk):
    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        super().__init__()

        self.title(APP_TITLE)
        self.geometry("1020x700")
        self.minsize(900, 620)
        self.configure(fg_color=COLOR_BG)

        self._adb = None
        self._relay = None
        self._running = False
        self._busy = False
        self._lock = threading.Lock()
        self.output_device = None
        self.capture_device = None

        self.status_phone = tk.StringVar(value="Aguardando...")
        self.status_mic = tk.StringVar(value="Aguardando...")
        self.status_stream = tk.StringVar(value="Parado")
        self.status_meter_percent = tk.StringVar(value="0%")
        self.status_meter_state = tk.StringVar(value="Silencio")
        self.status_meter_db = tk.StringVar(value="-60 dB")
        self.status_meter_source = tk.StringVar(value="Entrada: Aguardando...")
        self.status_header = tk.StringVar(value="PRONTO")
        self._meter_job = None
        self._meter_value = 0.0
        self._meter_stream = None
        self._meter_stream_name = None
        self._meter_live_peak = 0.0
        self._meter_lock = threading.Lock()

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(200, self.refresh_async)
        self.after(350, self._poll_meter)

    def _build_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=18, pady=18)
        main.grid_rowconfigure(1, weight=1)
        main.grid_columnconfigure(0, weight=1)

        hero = ctk.CTkFrame(
            main,
            fg_color=COLOR_SURFACE,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=20,
        )
        hero.grid(row=0, column=0, sticky="ew")
        hero.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            hero,
            text="MM",
            width=62,
            height=62,
            corner_radius=31,
            fg_color=COLOR_ACCENT,
            text_color="#03120F",
            font=ctk.CTkFont(family="Bahnschrift", size=25, weight="bold"),
        ).grid(row=0, column=0, rowspan=2, padx=(16, 12), pady=16)

        ctk.CTkLabel(
            hero,
            text="MICMIC Desktop",
            font=ctk.CTkFont(family="Bahnschrift", size=31, weight="bold"),
            text_color=COLOR_TEXT,
        ).grid(row=0, column=1, sticky="w", padx=(0, 10), pady=(14, 2))

        ctk.CTkLabel(
            hero,
            text="Controle profissional de stream USB para Discord e apps de voz.",
            font=ctk.CTkFont(size=13),
            text_color=COLOR_MUTED,
        ).grid(row=1, column=1, sticky="w", padx=(0, 10), pady=(0, 14))

        self.header_badge = ctk.CTkLabel(
            hero,
            textvariable=self.status_header,
            width=110,
            height=36,
            corner_radius=18,
            fg_color="#1A2638",
            text_color=COLOR_TEXT,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.header_badge.grid(row=0, column=2, rowspan=2, padx=(10, 16), pady=16, sticky="e")

        content = ctk.CTkFrame(main, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        content.grid_columnconfigure(0, weight=2)
        content.grid_columnconfigure(1, weight=3)
        content.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(content, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.grid_rowconfigure(3, weight=1)

        status_card = ctk.CTkFrame(
            left,
            fg_color=COLOR_CARD,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=16,
        )
        status_card.grid(row=0, column=0, sticky="ew")
        self._status_line(status_card, "Celular USB", self.status_phone).pack(fill="x", padx=12, pady=(12, 6))
        self._status_line(status_card, "Microfone", self.status_mic).pack(fill="x", padx=12, pady=6)
        self._status_line(status_card, "Stream", self.status_stream).pack(fill="x", padx=12, pady=(6, 8))

        self.discord_hint = ctk.CTkLabel(
            status_card,
            text="Discord: escolha MICMIC em Entrada de voz.",
            text_color=COLOR_INFO,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.discord_hint.pack(anchor="w", padx=14, pady=(0, 12))

        actions = ctk.CTkFrame(
            left,
            fg_color=COLOR_CARD,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=16,
        )
        actions.grid(row=1, column=0, sticky="ew", pady=(10, 10))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)

        self.btn_start = ctk.CTkButton(
            actions,
            text="START",
            command=self.start_async,
            height=50,
            fg_color=COLOR_OK,
            hover_color="#16A34A",
            text_color="#04140D",
            font=ctk.CTkFont(size=16, weight="bold"),
            corner_radius=12,
        )
        self.btn_start.grid(row=0, column=0, columnspan=2, sticky="ew", padx=12, pady=(12, 8))

        self.btn_stop = ctk.CTkButton(
            actions,
            text="STOP",
            command=self.stop_async,
            height=46,
            fg_color=COLOR_ERR,
            hover_color="#DC2626",
            text_color="#F8FAFC",
            font=ctk.CTkFont(size=15, weight="bold"),
            corner_radius=12,
            state="disabled",
        )
        self.btn_stop.grid(row=1, column=0, sticky="ew", padx=(12, 6), pady=(0, 12))

        self.btn_refresh = ctk.CTkButton(
            actions,
            text="Atualizar",
            command=self.refresh_async,
            height=46,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="#03120F",
            font=ctk.CTkFont(size=15, weight="bold"),
            corner_radius=12,
        )
        self.btn_refresh.grid(row=1, column=1, sticky="ew", padx=(6, 12), pady=(0, 12))

        guide_card = ctk.CTkFrame(
            left,
            fg_color=COLOR_CARD_ALT,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=16,
        )
        guide_card.grid(row=2, column=0, sticky="ew")
        ctk.CTkLabel(
            guide_card,
            text="Fluxo Rapido",
            text_color=COLOR_TEXT,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", padx=14, pady=(12, 4))
        ctk.CTkLabel(
            guide_card,
            text="1. Conecte USB + depuracao\n2. Clique START no PC\n3. Selecione MICMIC no Discord",
            justify="left",
            text_color=COLOR_MUTED,
            font=ctk.CTkFont(size=12),
        ).pack(anchor="w", padx=14, pady=(0, 12))

        right = ctk.CTkFrame(content, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        right.grid_rowconfigure(1, weight=1)

        meter_card = ctk.CTkFrame(
            right,
            fg_color=COLOR_CARD,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=16,
        )
        meter_card.grid(row=0, column=0, sticky="ew")

        ctk.CTkLabel(
            meter_card,
            text="Monitor de Voz",
            text_color=COLOR_TEXT,
            font=ctk.CTkFont(size=17, weight="bold"),
        ).pack(anchor="w", padx=14, pady=(12, 2))
        ctk.CTkLabel(
            meter_card,
            textvariable=self.status_meter_source,
            text_color=COLOR_MUTED,
            font=ctk.CTkFont(size=12),
        ).pack(anchor="w", padx=14, pady=(0, 10))

        self.voice_meter_bar = ctk.CTkProgressBar(
            meter_card,
            height=26,
            corner_radius=13,
            fg_color="#1A2538",
            progress_color=COLOR_OK,
            border_width=1,
            border_color="#31435F",
        )
        self.voice_meter_bar.pack(fill="x", padx=14)
        self.voice_meter_bar.set(0.0)

        meter_stats = ctk.CTkFrame(meter_card, fg_color="transparent")
        meter_stats.pack(fill="x", padx=14, pady=(11, 13))
        meter_stats.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            meter_stats,
            textvariable=self.status_meter_percent,
            text_color=COLOR_TEXT,
            font=ctk.CTkFont(size=27, weight="bold"),
        ).grid(row=0, column=0, sticky="w")
        self.voice_meter_state_label = ctk.CTkLabel(
            meter_stats,
            textvariable=self.status_meter_state,
            text_color=COLOR_OK,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.voice_meter_state_label.grid(row=0, column=1, sticky="e")
        ctk.CTkLabel(
            meter_stats,
            textvariable=self.status_meter_db,
            text_color=COLOR_MUTED,
            font=ctk.CTkFont(size=12),
        ).grid(row=1, column=0, columnspan=2, sticky="w")

        log_card = ctk.CTkFrame(
            right,
            fg_color=COLOR_SURFACE,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=16,
        )
        log_card.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        ctk.CTkLabel(
            log_card,
            text="Console de Eventos",
            text_color=COLOR_TEXT,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", padx=14, pady=(10, 6))

        self.log_box = ctk.CTkTextbox(
            log_card,
            height=230,
            fg_color="#0A1220",
            text_color=COLOR_TEXT,
            border_width=1,
            border_color="#263853",
            corner_radius=10,
        )
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.log_box.configure(state="disabled")

    def _status_line(self, parent, title: str, var: tk.StringVar):
        row = ctk.CTkFrame(
            parent,
            fg_color=COLOR_CARD_ALT,
            border_width=1,
            border_color=COLOR_BORDER,
            corner_radius=10,
        )
        ctk.CTkLabel(
            row,
            text=title,
            width=120,
            anchor="w",
            text_color=COLOR_MUTED,
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(side="left", padx=(10, 8), pady=9)
        ctk.CTkLabel(
            row,
            textvariable=var,
            anchor="w",
            text_color=COLOR_TEXT,
            font=ctk.CTkFont(size=13),
        ).pack(side="left", fill="x", expand=True, padx=(0, 10), pady=9)
        return row

    def _set_header_badge(self, text: str, color: str):
        def _apply():
            self.status_header.set(text)
            self.header_badge.configure(fg_color=color)
        self._ui(_apply)

    def _close_meter_stream(self):
        if self._meter_stream:
            try:
                self._meter_stream.stop()
            except Exception:
                pass
            try:
                self._meter_stream.close()
            except Exception:
                pass
        self._meter_stream = None
        self._meter_stream_name = None
        with self._meter_lock:
            self._meter_live_peak = 0.0

    def _meter_callback(self, indata, _frames, _time_info, _status):
        try:
            if not indata:
                return
            # Use the max absolute sample from the current block for fast response.
            peak = float(audioop.max(indata, 2)) / 32768.0
        except Exception:
            peak = 0.0
        peak = max(0.0, min(1.0, peak))
        with self._meter_lock:
            if peak > self._meter_live_peak:
                self._meter_live_peak = peak

    def _ensure_meter_stream(self):
        if sd is None or not self.capture_device:
            self._close_meter_stream()
            return

        capture_name = (self.capture_device.name or "").strip()
        if not capture_name:
            self._close_meter_stream()
            return

        if self._meter_stream and self._meter_stream_name == capture_name:
            return

        self._close_meter_stream()
        index = find_input_device_index_by_name(capture_name)
        if index is None:
            return

        try:
            dev_info = sd.query_devices(index)
            channels = int(dev_info.get("max_input_channels", 1))
            channels = 1 if channels < 1 else min(2, channels)
            sample_rate = int(float(dev_info.get("default_samplerate", SAMPLE_RATE)))
            if sample_rate <= 0:
                sample_rate = SAMPLE_RATE

            self._meter_stream = sd.RawInputStream(
                samplerate=sample_rate,
                channels=channels,
                dtype="int16",
                device=index,
                blocksize=1024,
                callback=self._meter_callback,
            )
            self._meter_stream.start()
            self._meter_stream_name = capture_name
        except Exception:
            self._close_meter_stream()

    def _apply_meter_level(self, peak: float):
        peak = max(0.0, min(1.0, float(peak)))
        if peak >= self._meter_value:
            self._meter_value = peak
        else:
            self._meter_value = (self._meter_value * 0.82) + (peak * 0.18)
        level = max(0.0, min(1.0, self._meter_value))

        if level >= 0.65:
            meter_color = COLOR_ERR
            meter_state = "Forte"
        elif level >= 0.22:
            meter_color = COLOR_WARN
            meter_state = "Falando"
        else:
            meter_color = COLOR_OK
            meter_state = "Silencio"

        percent = int(level * 100)
        db_value = int(20 * math.log10(max(level, 0.001)))

        self.status_meter_percent.set(f"{percent}%")
        self.status_meter_state.set(meter_state)
        self.status_meter_db.set(f"{db_value} dB")
        self.voice_meter_bar.configure(progress_color=meter_color)
        self.voice_meter_bar.set(level)
        self.voice_meter_state_label.configure(text_color=meter_color)

    def _poll_meter(self):
        self._ensure_meter_stream()
        with self._meter_lock:
            peak = self._meter_live_peak
            self._meter_live_peak *= 0.45
        self._apply_meter_level(peak)
        self._meter_job = self.after(95, self._poll_meter)

    def _ui(self, fn):
        try:
            self.after(0, fn)
        except (RuntimeError, tk.TclError):
            pass

    def log(self, text: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        def _write():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", f"[{stamp}] {text}\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self._ui(_write)

    def _set_busy(self, busy: bool):
        self._busy = busy
        self._ui(lambda: self.btn_refresh.configure(state="disabled" if busy else "normal"))
        self._ui(lambda: self.btn_start.configure(state="disabled" if busy or self._running else "normal"))
        self._ui(lambda: self.btn_stop.configure(state="normal" if self._running and not busy else "disabled"))

    def _set_status(self, phone=None, mic=None, stream=None):
        if phone is not None:
            self._ui(lambda: self.status_phone.set(phone))
        if mic is not None:
            self._ui(lambda: self.status_mic.set(mic))
            self._ui(lambda: self.status_meter_source.set(f"Entrada: {mic}"))
        if stream is not None:
            self._ui(lambda: self.status_stream.set(stream))
            stream_lower = stream.lower()
            if "transmit" in stream_lower:
                self._set_header_badge("AO VIVO", "#14532D")
            elif "erro" in stream_lower:
                self._set_header_badge("ERRO", "#7F1D1D")
            else:
                self._set_header_badge("PRONTO", "#1E293B")

    def refresh_async(self):
        if self._busy:
            return
        threading.Thread(target=self._refresh_worker, daemon=True).start()

    def _refresh_worker(self):
        self._set_busy(True)
        try:
            outputs = list_output_devices()
            captures = list_capture_devices()
            self.output_device = choose_preferred(outputs, PREFERRED_OUTPUT_HINTS)
            self.capture_device = choose_preferred(captures, PREFERRED_CAPTURE_HINTS)
            if not self.output_device:
                raise MicBridgeError("Nao achei saida virtual (CABLE Input / Virtual Speakers).")
            if not self.capture_device:
                raise MicBridgeError("Nao achei microfone virtual (MICMIC / Virtual Mic / CABLE Output).")

            self._adb = AdbClient()
            phone = self._adb.get_connected_device()
            self._set_status(phone=f"{phone.model} ({phone.serial})")

            mic_label = self.capture_device.name
            self._set_status(mic=mic_label, stream=("Transmitindo" if self._running else "Parado"))
            self.log(f"Pronto. Mic detectado: {mic_label}")
        except Exception as exc:
            self._set_status(phone=str(exc))
            self.log(f"Diagnostico: {exc}")
        finally:
            self._set_busy(False)

    def start_async(self):
        if self._busy:
            return
        threading.Thread(target=self._start_worker, daemon=True).start()

    def _start_worker(self):
        with self._lock:
            if self._running:
                return
            self._running = True
        self._set_busy(True)
        try:
            self._refresh_worker()
            if not self._adb or not self.output_device or not self.capture_device:
                raise MicBridgeError("Ambiente incompleto para iniciar.")

            phone = self._adb.get_connected_device()
            if not self._adb.is_package_installed(APP_PACKAGE):
                self.log("APK nao encontrado no celular. Instalando...")
                self._adb.install_apk(APK_PATH)
                self.log("APK instalado com sucesso.")

            if TARGET_MIC_NAME.lower() not in self.capture_device.name.lower():
                ok, msg = try_rename_capture_device_to_micmic(self.capture_device)
                self.log(msg)
                if ok:
                    captures = list_capture_devices()
                    selected = choose_preferred(captures, PREFERRED_CAPTURE_HINTS)
                    if selected:
                        self.capture_device = selected

            set_default_capture_device(self.capture_device)
            self.log(f"Microfone padrao do Windows: {self.capture_device.name}")

            route_peak = measure_virtual_route_peak(self.output_device, self.capture_device)
            self.log(f"Teste da rota virtual (peak): {route_peak:.4f}")
            if route_peak < 0.005:
                raise MicBridgeError(
                    "Sem sinal no mic virtual. Driver virtual sem roteamento. "
                    "Use VB-CABLE (CABLE Input/Output) ou reinstale o driver MICMIC/AudioRelay."
                )

            self._relay = AudioRelay(self.output_device.index, self.output_device.name, self._relay_status)
            self._relay.start()

            self._adb.run(["reverse", f"tcp:{ADB_PORT}", f"tcp:{ADB_PORT}"])
            self._adb.run(
                ["shell", "am", "start", "-n", f"{APP_PACKAGE}/.MainActivity", "--es", "command", "start"]
            )
            self._set_status(stream="Transmitindo")
            self.log(f"Stream iniciado com {phone.model}.")

            if TARGET_MIC_NAME.lower() in self.capture_device.name.lower():
                hint = "Discord: selecione MICMIC."
            else:
                hint = f"Discord: selecione '{self.capture_device.name}' ou 'Default'."
            self._ui(lambda: self.discord_hint.configure(text=hint))
        except Exception as exc:
            self.log(f"Falha no start: {exc}")
            self._safe_stop()
            with self._lock:
                self._running = False
            self._set_status(stream="Erro")
        finally:
            self._set_busy(False)

    def stop_async(self):
        if self._busy:
            return
        threading.Thread(target=self._stop_worker, daemon=True).start()

    def _stop_worker(self):
        self._set_busy(True)
        try:
            self._safe_stop()
            with self._lock:
                self._running = False
            self._set_status(stream="Parado")
            self.log("Stream parado.")
        finally:
            self._set_busy(False)

    def _relay_status(self, message: str, _level: str):
        self.log(message)

    def _safe_stop(self):
        try:
            adb = self._adb or AdbClient()
            adb.run(
                ["shell", "am", "start", "-n", f"{APP_PACKAGE}/.MainActivity", "--es", "command", "stop"],
                check=False,
            )
            adb.run(["reverse", "--remove", f"tcp:{ADB_PORT}"], check=False)
            # Hard stop to avoid service lingering in edge cases.
            adb.run(["shell", "am", "force-stop", APP_PACKAGE], check=False)
        except Exception:
            pass
        if self._relay:
            self._relay.stop()
            self._relay = None

    def on_close(self):
        if self._meter_job is not None:
            try:
                self.after_cancel(self._meter_job)
            except Exception:
                pass
            self._meter_job = None
        self._close_meter_stream()
        if self._running:
            self._safe_stop()
        self.destroy()


def main():
    app = MicMicApp()
    app.mainloop()


if __name__ == "__main__":
    main()
