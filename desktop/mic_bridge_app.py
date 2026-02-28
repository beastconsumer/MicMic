import os
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
    from pycaw.pycaw import AudioUtilities, EDataFlow, ERole, IMMEndpoint
except ImportError:  # pragma: no cover
    comtypes = None
    AudioUtilities = None
    IMMEndpoint = None
    EDataFlow = None
    ERole = None

try:
    import winreg
except ImportError:  # pragma: no cover
    winreg = None

APP_TITLE = "MICMIC"
APP_PACKAGE = "com.micmic.mobilemic"
ADB_PORT = 28282
SAMPLE_RATE = 48000
CHANNELS = 1
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

COLOR_BG = "#080B13"
COLOR_CARD = "#111827"
COLOR_TEXT = "#E5E7EB"
COLOR_OK = "#22C55E"
COLOR_WARN = "#F59E0B"
COLOR_ERR = "#EF4444"
COLOR_INFO = "#94A3B8"


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
                    self._stream.write(chunk)
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
        self.geometry("620x520")
        self.minsize(560, 470)
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

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(200, self.refresh_async)

    def _build_ui(self):
        ctk.CTkLabel(self, text="MICMIC", font=ctk.CTkFont(size=34, weight="bold"), text_color=COLOR_TEXT).pack(
            pady=(16, 2)
        )
        ctk.CTkLabel(
            self,
            text="Clique em START e use no Discord.",
            text_color="#94A3B8",
            font=ctk.CTkFont(size=14),
        ).pack(pady=(0, 10))

        card = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=14)
        card.pack(fill="x", padx=16, pady=(0, 10))

        self._status_line(card, "Celular USB", self.status_phone).pack(fill="x", padx=12, pady=(10, 4))
        self._status_line(card, "Microfone", self.status_mic).pack(fill="x", padx=12, pady=4)
        self._status_line(card, "Stream", self.status_stream).pack(fill="x", padx=12, pady=(4, 10))

        self.discord_hint = ctk.CTkLabel(
            self,
            text="Discord: se MICMIC nao aparecer, use 'Default'.",
            text_color="#93C5FD",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.discord_hint.pack(pady=(2, 8))

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.pack(fill="x", padx=16, pady=(0, 10))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)
        actions.grid_columnconfigure(2, weight=1)

        self.btn_refresh = ctk.CTkButton(actions, text="Atualizar", command=self.refresh_async, height=45)
        self.btn_refresh.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.btn_start = ctk.CTkButton(
            actions,
            text="START",
            command=self.start_async,
            height=45,
            fg_color="#22C55E",
            hover_color="#16A34A",
            text_color="#05270f",
            font=ctk.CTkFont(size=16, weight="bold"),
        )
        self.btn_start.grid(row=0, column=1, sticky="ew", padx=6)

        self.btn_stop = ctk.CTkButton(
            actions,
            text="STOP",
            command=self.stop_async,
            height=45,
            fg_color="#EF4444",
            hover_color="#DC2626",
            font=ctk.CTkFont(size=16, weight="bold"),
            state="disabled",
        )
        self.btn_stop.grid(row=0, column=2, sticky="ew", padx=(6, 0))

        self.log_box = ctk.CTkTextbox(self, height=230, fg_color="#020617", text_color=COLOR_TEXT)
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.log_box.configure(state="disabled")

    def _status_line(self, parent, title: str, var: tk.StringVar):
        row = ctk.CTkFrame(parent, fg_color="#0B1220", corner_radius=10)
        ctk.CTkLabel(row, text=title, width=120, anchor="w", text_color="#CBD5E1").pack(side="left", padx=(10, 6), pady=8)
        ctk.CTkLabel(row, textvariable=var, anchor="w", text_color=COLOR_TEXT).pack(side="left", fill="x", expand=True, padx=(0, 10), pady=8)
        return row

    def _ui(self, fn):
        self.after(0, fn)

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
        if stream is not None:
            self._ui(lambda: self.status_stream.set(stream))

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
        if self._running:
            self._safe_stop()
        self.destroy()


def main():
    app = MicMicApp()
    app.mainloop()


if __name__ == "__main__":
    main()
