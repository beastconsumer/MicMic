
import json
import os
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import shutil
import socket
import subprocess
import threading
import tkinter as tk
import webbrowser

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

APP_TITLE = "MICMIC Studio"
APP_PACKAGE = "com.micmic.mobilemic"
ADB_PORT = 28282
SAMPLE_RATE = 48000
CHANNELS = 1
BLOCK_SIZE = 2048

ROOT_DIR = Path(__file__).resolve().parents[1]
APK_PATH = ROOT_DIR / "android" / "app" / "build" / "outputs" / "apk" / "debug" / "app-debug.apk"
CONFIG_PATH = Path(__file__).resolve().with_name("mic_bridge_config.json")

PRIMARY_COLOR = "#22C55E"
SECONDARY_COLOR = "#38BDF8"
ERROR_COLOR = "#FB7185"
WARN_COLOR = "#FBBF24"
TEXT_DIM = "#9CA3AF"

PREFERRED_OUTPUT_HINTS = [
    "Virtual Speakers for AudioRelay",
    "CABLE Input",
    "VB-Audio",
]
PREFERRED_CAPTURE_HINTS = [
    "Virtual Mic for AudioRelay",
    "CABLE Output",
    "VB-Audio",
    "WO Mic",
]

VIRTUAL_DRIVER_GUIDE = "https://vb-audio.com/Cable/"


class MicBridgeError(RuntimeError):
    pass


@dataclass
class OutputDevice:
    index: int
    name: str

    @property
    def label(self) -> str:
        return f"{self.name}  (#{self.index})"


@dataclass
class CaptureDevice:
    device_id: str
    name: str

    @property
    def label(self) -> str:
        return self.name


@dataclass
class AdbDevice:
    serial: str
    state: str
    model: str
    product: str


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


def load_config() -> dict:
    default = {
        "output_label": "",
        "capture_label": "",
        "auto_set_default": True,
    }
    if not CONFIG_PATH.exists():
        return default
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        default.update(data)
    except Exception:
        return default
    return default


def save_config(data: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")


def resolve_adb_executable() -> str:
    env_candidate = os.environ.get("ADB_PATH", "").strip()
    if env_candidate:
        path = Path(env_candidate)
        if path.exists():
            return str(path)

    in_path = shutil.which("adb")
    if in_path:
        return in_path

    local_app_data = os.environ.get("LOCALAPPDATA", "")
    fallback_candidates = [
        Path(local_app_data) / "Android" / "Sdk" / "platform-tools" / "adb.exe",
        Path.home() / "AppData" / "Local" / "Android" / "Sdk" / "platform-tools" / "adb.exe",
    ]
    for candidate in fallback_candidates:
        if candidate.exists():
            return str(candidate)

    raise MicBridgeError(
        "ADB nao encontrado. Instale Android Platform Tools ou defina ADB_PATH com o caminho do adb.exe."
    )


def run_process(command, check=True, timeout=40):
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=timeout,
        shell=False,
    )
    if check and completed.returncode != 0:
        stderr = completed.stderr.strip() or completed.stdout.strip() or "sem detalhes"
        raise MicBridgeError(f"Falha em {' '.join(command)}: {stderr}")
    return completed


def parse_adb_devices(stdout: str) -> list[AdbDevice]:
    devices = []
    lines = stdout.splitlines()
    for line in lines[1:]:
        line = line.strip()
        if not line:
            continue
        parts = line.split()
        serial = parts[0]
        state = parts[1] if len(parts) > 1 else "unknown"
        metadata = {}
        for part in parts[2:]:
            if ":" in part:
                key, value = part.split(":", 1)
                metadata[key] = value
        devices.append(
            AdbDevice(
                serial=serial,
                state=state,
                model=metadata.get("model", "-"),
                product=metadata.get("product", "-"),
            )
        )
    return devices


class AdbClient:
    def __init__(self):
        self.adb_executable = resolve_adb_executable()

    def run(self, args, check=True, timeout=40):
        return run_process([self.adb_executable, *args], check=check, timeout=timeout)

    def list_devices(self) -> list[AdbDevice]:
        completed = self.run(["devices", "-l"], check=False)
        return parse_adb_devices(completed.stdout)

    def get_connected_device(self) -> AdbDevice:
        devices = self.list_devices()
        active = [d for d in devices if d.state == "device"]
        if active:
            return active[0]

        if any(d.state == "unauthorized" for d in devices):
            raise MicBridgeError("Celular encontrado, mas nao autorizado. Aceite a chave RSA no celular.")
        if any(d.state == "offline" for d in devices):
            raise MicBridgeError("Celular offline no ADB. Reconecte o cabo USB e tente novamente.")
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
        raise MicBridgeError("Dependencia de audio nao instalada. Rode: python -m pip install -r requirements.txt")

    devices = []
    for index, dev in enumerate(sd.query_devices()):
        if int(dev.get("max_output_channels", 0)) <= 0:
            continue
        name = str(dev.get("name", f"Output {index}"))
        devices.append(OutputDevice(index=index, name=name))
    return devices


def list_capture_devices() -> list[CaptureDevice]:
    if AudioUtilities is None or IMMEndpoint is None or EDataFlow is None:
        raise MicBridgeError("Dependencia Windows audio nao instalada. Rode: python -m pip install -r requirements.txt")

    capture_devices = {}
    with com_scope():
        for dev in AudioUtilities.GetAllDevices():
            try:
                endpoint = dev._dev.QueryInterface(IMMEndpoint)
                flow = endpoint.GetDataFlow()
            except Exception:
                continue

            if flow != EDataFlow.eCapture.value:
                continue
            if not dev.FriendlyName:
                continue
            capture_devices[dev.id] = CaptureDevice(device_id=dev.id, name=dev.FriendlyName)

    return sorted(capture_devices.values(), key=lambda item: item.name.lower())


def pick_preferred(items, hints):
    lowered = [h.lower() for h in hints]
    for hint in lowered:
        for item in items:
            text = (item.name if hasattr(item, "name") else str(item)).lower()
            if hint in text:
                return item
    return items[0] if items else None


def set_default_capture_device(device: CaptureDevice):
    if AudioUtilities is None or ERole is None:
        raise MicBridgeError("Nao foi possivel carregar pycaw para configurar microfone padrao.")

    with com_scope():
        AudioUtilities.SetDefaultDevice(
            device.device_id,
            roles=[ERole.eConsole, ERole.eMultimedia, ERole.eCommunications],
        )


class AudioRelay:
    def __init__(self, output_device_index: int, output_label: str, status_callback):
        self.output_device_index = output_device_index
        self.output_label = output_label
        self.status_callback = status_callback

        self._stop_event = threading.Event()
        self._thread = None
        self._server_socket = None
        self._client_socket = None
        self._stream = None

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True, name="audio-relay")
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        for sock in (self._client_socket, self._server_socket):
            if sock:
                try:
                    sock.shutdown(socket.SHUT_RDWR)
                except OSError:
                    pass
                try:
                    sock.close()
                except OSError:
                    pass
        self._client_socket = None
        self._server_socket = None
        if self._thread:
            self._thread.join(timeout=3)
            self._thread = None

    def _run(self):
        try:
            self._server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self._server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self._server_socket.bind(("127.0.0.1", ADB_PORT))
            self._server_socket.listen(1)
            self._server_socket.settimeout(1.0)
            self.status_callback("Aguardando audio do celular...", False)

            while not self._stop_event.is_set():
                try:
                    self._client_socket, addr = self._server_socket.accept()
                    self.status_callback(f"Celular conectado ({addr[0]}:{addr[1]})", False)
                    break
                except socket.timeout:
                    continue

            if self._stop_event.is_set() or not self._client_socket:
                return

            self._client_socket.settimeout(1.0)
            self._stream = sd.RawOutputStream(
                samplerate=SAMPLE_RATE,
                channels=CHANNELS,
                dtype="int16",
                device=self.output_device_index,
                blocksize=BLOCK_SIZE,
            )
            self._stream.start()
            self.status_callback(f"Stream ativo em: {self.output_label}", False)

            while not self._stop_event.is_set():
                try:
                    chunk = self._client_socket.recv(4096)
                except socket.timeout:
                    continue

                if not chunk:
                    self.status_callback("Conexao de audio encerrada pelo celular.", True)
                    break

                if len(chunk) % 2 != 0:
                    chunk = chunk[:-1]
                if chunk:
                    self._stream.write(chunk)

        except Exception as exc:  # pragma: no cover
            self.status_callback(f"Erro no relay de audio: {exc}", True)
        finally:
            if self._stream:
                try:
                    self._stream.stop()
                    self._stream.close()
                except Exception:
                    pass
                self._stream = None
            if self._client_socket:
                try:
                    self._client_socket.close()
                except OSError:
                    pass
                self._client_socket = None
            if self._server_socket:
                try:
                    self._server_socket.close()
                except OSError:
                    pass
                self._server_socket = None

class MicMicStudioApp(ctk.CTk):
    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1180x760")
        self.minsize(1000, 680)

        self._relay = None
        self._running = False
        self._busy = False
        self._lock = threading.Lock()

        self._config = load_config()
        self._adb = None
        self._output_devices = []
        self._capture_devices = []

        self.output_lookup = {}
        self.capture_lookup = {}

        self.output_var = tk.StringVar(value="")
        self.capture_var = tk.StringVar(value="")
        self.auto_default_var = tk.BooleanVar(value=bool(self._config.get("auto_set_default", True)))

        self._build_menu()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.after(250, self.refresh_diagnostics_async)

    def _build_menu(self):
        menu_bar = tk.Menu(self)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Atualizar diagnostico", command=self.refresh_diagnostics_async)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.on_close)
        menu_bar.add_cascade(label="Arquivo", menu=file_menu)

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Guia driver virtual", command=self.open_virtual_driver_guide)
        menu_bar.add_cascade(label="Ajuda", menu=help_menu)

        self.config(menu=menu_bar)

    def _build_ui(self):
        root = ctk.CTkFrame(self, fg_color="#090E16", corner_radius=0)
        root.pack(fill="both", expand=True)

        header = ctk.CTkFrame(root, fg_color="#0F172A", corner_radius=16)
        header.pack(fill="x", padx=18, pady=(16, 10))

        ctk.CTkLabel(
            header,
            text="MICMIC Studio",
            font=ctk.CTkFont(size=30, weight="bold"),
            text_color="#E5F3FF",
        ).pack(anchor="w", padx=20, pady=(18, 2))

        ctk.CTkLabel(
            header,
            text="Conecte o microfone do celular por USB e use no Discord com um clique.",
            font=ctk.CTkFont(size=14),
            text_color="#93C5FD",
        ).pack(anchor="w", padx=20, pady=(0, 16))

        body = ctk.CTkFrame(root, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)

        left_panel = ctk.CTkFrame(body, fg_color="#111827", corner_radius=16)
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        right_panel = ctk.CTkFrame(body, fg_color="#111827", corner_radius=16)
        right_panel.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(
            left_panel,
            text="Diagnostico Rapido",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#F8FAFC",
        ).pack(anchor="w", padx=16, pady=(16, 14))

        self.phone_status = self._create_status_row(left_panel, "Celular USB")
        self.apk_status = self._create_status_row(left_panel, "App Android")
        self.virtual_status = self._create_status_row(left_panel, "Microfone virtual")
        self.stream_status = self._create_status_row(left_panel, "Stream")

        ctk.CTkButton(
            left_panel,
            text="Atualizar diagnostico",
            command=self.refresh_diagnostics_async,
            fg_color="#1D4ED8",
            hover_color="#1E40AF",
            height=40,
        ).pack(fill="x", padx=16, pady=(18, 8))

        ctk.CTkButton(
            left_panel,
            text="Instalar/Reinstalar APK",
            command=self.install_apk_async,
            fg_color="#0EA5E9",
            hover_color="#0369A1",
            height=40,
        ).pack(fill="x", padx=16, pady=8)

        ctk.CTkButton(
            left_panel,
            text="Abrir App no celular",
            command=self.open_mobile_app_async,
            fg_color="#334155",
            hover_color="#1E293B",
            height=40,
        ).pack(fill="x", padx=16, pady=8)

        ctk.CTkButton(
            left_panel,
            text="Definir mic padrao Windows",
            command=self.set_default_mic_async,
            fg_color="#15803D",
            hover_color="#166534",
            height=40,
        ).pack(fill="x", padx=16, pady=8)

        ctk.CTkButton(
            left_panel,
            text="Guia do driver virtual",
            command=self.open_virtual_driver_guide,
            fg_color="#7C3AED",
            hover_color="#6D28D9",
            height=38,
        ).pack(fill="x", padx=16, pady=(8, 16))

        ctk.CTkLabel(
            right_panel,
            text="Configuracao do audio",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#F8FAFC",
        ).pack(anchor="w", padx=18, pady=(16, 8))

        fields = ctk.CTkFrame(right_panel, fg_color="#0B1220", corner_radius=12)
        fields.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(fields, text="Saida para stream (render)", text_color="#BFDBFE").pack(
            anchor="w", padx=14, pady=(12, 6)
        )
        self.output_combo = ctk.CTkComboBox(
            fields,
            values=[],
            variable=self.output_var,
            command=self.on_output_changed,
            state="readonly",
        )
        self.output_combo.pack(fill="x", padx=14, pady=(0, 10))

        ctk.CTkLabel(fields, text="Microfone no Windows/Discord (capture)", text_color="#BFDBFE").pack(
            anchor="w", padx=14, pady=(0, 6)
        )
        self.capture_combo = ctk.CTkComboBox(
            fields,
            values=[],
            variable=self.capture_var,
            command=self.on_capture_changed,
            state="readonly",
        )
        self.capture_combo.pack(fill="x", padx=14, pady=(0, 12))

        self.auto_default_switch = ctk.CTkSwitch(
            fields,
            text="Ao iniciar, definir microfone padrao automaticamente",
            variable=self.auto_default_var,
            onvalue=True,
            offvalue=False,
            command=self.on_auto_default_changed,
            text_color="#C7D2FE",
        )
        self.auto_default_switch.pack(anchor="w", padx=14, pady=(0, 12))

        self.alias_label = ctk.CTkLabel(
            right_panel,
            text="Alias do mic no app: MICMIC Virtual Mic",
            text_color="#A5F3FC",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.alias_label.pack(anchor="w", padx=18, pady=(2, 10))

        actions = ctk.CTkFrame(right_panel, fg_color="transparent")
        actions.pack(fill="x", padx=16, pady=(0, 10))
        actions.grid_columnconfigure(0, weight=1)
        actions.grid_columnconfigure(1, weight=1)

        self.start_button = ctk.CTkButton(
            actions,
            text="START STREAM",
            command=self.start_stream_async,
            fg_color=PRIMARY_COLOR,
            hover_color="#16A34A",
            text_color="#06250E",
            font=ctk.CTkFont(size=18, weight="bold"),
            height=54,
        )
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.stop_button = ctk.CTkButton(
            actions,
            text="STOP",
            command=self.stop_stream_async,
            fg_color="#EF4444",
            hover_color="#DC2626",
            font=ctk.CTkFont(size=18, weight="bold"),
            height=54,
            state="disabled",
        )
        self.stop_button.grid(row=0, column=1, sticky="ew", padx=(8, 0))

        ctk.CTkLabel(
            right_panel,
            text="Logs em tempo real",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#E2E8F0",
        ).pack(anchor="w", padx=18, pady=(2, 6))

        self.log_box = ctk.CTkTextbox(
            right_panel,
            height=260,
            fg_color="#020617",
            text_color="#E2E8F0",
            corner_radius=10,
            border_width=1,
            border_color="#1E293B",
        )
        self.log_box.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.log_box.configure(state="disabled")

    def _create_status_row(self, parent, title):
        row = ctk.CTkFrame(parent, fg_color="#0B1220", corner_radius=10)
        row.pack(fill="x", padx=14, pady=6)

        ctk.CTkLabel(row, text=title, text_color="#CBD5E1", font=ctk.CTkFont(size=14)).pack(
            anchor="w", padx=12, pady=(8, 2)
        )

        value = ctk.CTkLabel(
            row,
            text="Aguardando...",
            text_color=TEXT_DIM,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        value.pack(anchor="w", padx=12, pady=(0, 10))
        return value

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")

        def _update():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", f"[{timestamp}] {message}\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

        self.after(0, _update)

    def _set_status_label(self, label_widget, text: str, kind: str = "info"):
        color_map = {
            "ok": PRIMARY_COLOR,
            "warn": WARN_COLOR,
            "error": ERROR_COLOR,
            "info": TEXT_DIM,
            "active": SECONDARY_COLOR,
        }
        label_widget.configure(text=text, text_color=color_map.get(kind, TEXT_DIM))

    def _run_async(self, target):
        thread = threading.Thread(target=target, daemon=True)
        thread.start()

    def on_output_changed(self, value):
        self._config["output_label"] = value
        save_config(self._config)

    def on_capture_changed(self, value):
        self._config["capture_label"] = value
        self._refresh_alias_label(value)
        save_config(self._config)

    def on_auto_default_changed(self):
        self._config["auto_set_default"] = bool(self.auto_default_var.get())
        save_config(self._config)

    def _refresh_alias_label(self, selected_capture_label: str):
        if selected_capture_label:
            self.alias_label.configure(
                text=f"Alias do mic no app: MICMIC Virtual Mic -> {selected_capture_label}"
            )
        else:
            self.alias_label.configure(text="Alias do mic no app: MICMIC Virtual Mic")

    def set_busy(self, busy: bool):
        self._busy = busy
        self.after(
            0,
            lambda: self.start_button.configure(
                state="disabled" if busy or self._running else "normal"
            ),
        )
        self.after(
            0,
            lambda: self.stop_button.configure(
                state="normal" if self._running and not busy else "disabled"
            ),
        )
        self.after(0, lambda: self.output_combo.configure(state="readonly" if not busy else "disabled"))
        self.after(0, lambda: self.capture_combo.configure(state="readonly" if not busy else "disabled"))

    def refresh_diagnostics_async(self):
        if self._busy:
            return
        self._run_async(self._refresh_diagnostics_worker)

    def _refresh_diagnostics_worker(self):
        self.set_busy(True)
        try:
            self.log("Atualizando diagnostico completo...")

            self._output_devices = list_output_devices()
            self._capture_devices = list_capture_devices()

            self.output_lookup = {d.label: d for d in self._output_devices}
            self.capture_lookup = {d.label: d for d in self._capture_devices}

            output_values = [d.label for d in self._output_devices]
            capture_values = [d.label for d in self._capture_devices]

            self.after(0, lambda: self.output_combo.configure(values=output_values))
            self.after(0, lambda: self.capture_combo.configure(values=capture_values))

            selected_output = self._config.get("output_label", "")
            if selected_output not in self.output_lookup:
                preferred_output = pick_preferred(self._output_devices, PREFERRED_OUTPUT_HINTS)
                selected_output = preferred_output.label if preferred_output else ""

            selected_capture = self._config.get("capture_label", "")
            if selected_capture not in self.capture_lookup:
                preferred_capture = pick_preferred(self._capture_devices, PREFERRED_CAPTURE_HINTS)
                selected_capture = preferred_capture.label if preferred_capture else ""

            self.after(0, lambda: self.output_var.set(selected_output))
            self.after(0, lambda: self.capture_var.set(selected_capture))
            self.after(0, lambda: self._refresh_alias_label(selected_capture))

            self._config["output_label"] = selected_output
            self._config["capture_label"] = selected_capture
            save_config(self._config)

            if selected_capture:
                self.after(
                    0,
                    lambda: self._set_status_label(
                        self.virtual_status,
                        f"Pronto: {selected_capture}",
                        "ok",
                    ),
                )
            else:
                self.after(
                    0,
                    lambda: self._set_status_label(
                        self.virtual_status,
                        "Sem mic virtual selecionado. Instale VB-CABLE ou use Virtual Mic.",
                        "warn",
                    ),
                )

            self._adb = AdbClient()
            devices = self._adb.list_devices()
            active = [d for d in devices if d.state == "device"]

            if active:
                device = active[0]
                phone_line = f"Conectado: {device.model} ({device.serial})"
                self.after(0, lambda: self._set_status_label(self.phone_status, phone_line, "ok"))

                if self._adb.is_package_installed(APP_PACKAGE):
                    self.after(
                        0,
                        lambda: self._set_status_label(self.apk_status, "APK instalado no celular", "ok"),
                    )
                else:
                    self.after(
                        0,
                        lambda: self._set_status_label(self.apk_status, "APK nao instalado", "warn"),
                    )
            else:
                warning = "Nenhum celular conectado"
                if any(d.state == "unauthorized" for d in devices):
                    warning = "Celular nao autorizado (aceite chave RSA)"
                elif any(d.state == "offline" for d in devices):
                    warning = "Celular offline no ADB"

                self.after(0, lambda: self._set_status_label(self.phone_status, warning, "error"))
                self.after(
                    0,
                    lambda: self._set_status_label(self.apk_status, "Aguardando celular", "info"),
                )

            if not self._running:
                self.after(
                    0,
                    lambda: self._set_status_label(self.stream_status, "Parado", "info"),
                )
            self.log("Diagnostico atualizado.")

        except Exception as exc:
            self.log(f"Erro no diagnostico: {exc}")
            self.after(0, lambda: self._set_status_label(self.phone_status, str(exc), "error"))
        finally:
            self.set_busy(False)

    def install_apk_async(self):
        if self._busy:
            return
        self._run_async(self._install_apk_worker)

    def _install_apk_worker(self):
        self.set_busy(True)
        try:
            adb = self._adb or AdbClient()
            device = adb.get_connected_device()
            self.log(f"Instalando APK no celular {device.model}...")
            adb.install_apk(APK_PATH)
            self.log("APK instalado com sucesso.")
            self.after(
                0,
                lambda: self._set_status_label(self.apk_status, "APK instalado no celular", "ok"),
            )
        except Exception as exc:
            self.log(f"Falha ao instalar APK: {exc}")
            self.after(0, lambda: self._set_status_label(self.apk_status, str(exc), "error"))
        finally:
            self.set_busy(False)

    def open_mobile_app_async(self):
        if self._busy:
            return
        self._run_async(self._open_mobile_app_worker)

    def _open_mobile_app_worker(self):
        self.set_busy(True)
        try:
            adb = self._adb or AdbClient()
            adb.get_connected_device()
            adb.run(["shell", "am", "start", "-n", f"{APP_PACKAGE}/.MainActivity"])
            self.log("App do celular aberto.")
        except Exception as exc:
            self.log(f"Nao foi possivel abrir app no celular: {exc}")
        finally:
            self.set_busy(False)

    def set_default_mic_async(self):
        if self._busy:
            return
        self._run_async(self._set_default_mic_worker)

    def _set_default_mic_worker(self):
        self.set_busy(True)
        try:
            capture_label = self.capture_var.get().strip()
            capture_device = self.capture_lookup.get(capture_label)
            if not capture_device:
                raise MicBridgeError("Selecione um microfone virtual valido para o Windows.")

            set_default_capture_device(capture_device)
            self.log(f"Microfone padrao definido: {capture_device.name}")
            self.after(
                0,
                lambda: self._set_status_label(
                    self.virtual_status,
                    f"Padrao Windows: {capture_device.name}",
                    "ok",
                ),
            )
        except Exception as exc:
            self.log(f"Falha ao definir mic padrao: {exc}")
            self.after(0, lambda: self._set_status_label(self.virtual_status, str(exc), "error"))
        finally:
            self.set_busy(False)

    def start_stream_async(self):
        if self._busy:
            return
        self._run_async(self._start_stream_worker)

    def _start_stream_worker(self):
        with self._lock:
            if self._running:
                return
            self._running = True

        self.set_busy(True)

        try:
            output_label = self.output_var.get().strip()
            capture_label = self.capture_var.get().strip()

            output_device = self.output_lookup.get(output_label)
            capture_device = self.capture_lookup.get(capture_label)

            if not output_device:
                raise MicBridgeError("Selecione uma saida de stream valida.")
            if not capture_device:
                raise MicBridgeError(
                    "Selecione um microfone virtual valido. Dica: Virtual Mic ou CABLE Output."
                )

            adb = self._adb or AdbClient()
            phone = adb.get_connected_device()

            if not adb.is_package_installed(APP_PACKAGE):
                self.log("APK nao encontrado no celular. Instalando automaticamente...")
                adb.install_apk(APK_PATH)

            if self.auto_default_var.get():
                set_default_capture_device(capture_device)
                self.log(f"Microfone padrao definido automaticamente: {capture_device.name}")

            self._relay = AudioRelay(
                output_device_index=output_device.index,
                output_label=output_device.name,
                status_callback=self._relay_status_callback,
            )
            self._relay.start()

            adb.run(["reverse", f"tcp:{ADB_PORT}", f"tcp:{ADB_PORT}"])
            adb.run(
                [
                    "shell",
                    "am",
                    "start",
                    "-n",
                    f"{APP_PACKAGE}/.MainActivity",
                    "--es",
                    "command",
                    "start",
                ]
            )

            self.log(f"Stream iniciado com {phone.model}.")
            self.after(
                0,
                lambda: self._set_status_label(
                    self.stream_status,
                    f"Transmitindo | Mic alvo: {capture_device.name}",
                    "active",
                ),
            )
            self.after(0, lambda: self.start_button.configure(state="disabled"))
            self.after(0, lambda: self.stop_button.configure(state="normal"))

        except Exception as exc:
            self.log(f"Falha ao iniciar stream: {exc}")
            self._safe_stop_backend()
            with self._lock:
                self._running = False
            self.after(0, lambda: self._set_status_label(self.stream_status, str(exc), "error"))
            self.after(0, lambda: self.start_button.configure(state="normal"))
            self.after(0, lambda: self.stop_button.configure(state="disabled"))
        finally:
            self.set_busy(False)

    def stop_stream_async(self):
        if self._busy:
            return
        self._run_async(self._stop_stream_worker)

    def _stop_stream_worker(self):
        self.set_busy(True)
        try:
            self._safe_stop_backend()
            with self._lock:
                self._running = False

            self.log("Stream parado.")
            self.after(0, lambda: self._set_status_label(self.stream_status, "Parado", "info"))
            self.after(0, lambda: self.start_button.configure(state="normal"))
            self.after(0, lambda: self.stop_button.configure(state="disabled"))
        finally:
            self.set_busy(False)

    def _safe_stop_backend(self):
        try:
            adb = self._adb or AdbClient()
            adb.run(
                [
                    "shell",
                    "am",
                    "start",
                    "-n",
                    f"{APP_PACKAGE}/.MainActivity",
                    "--es",
                    "command",
                    "stop",
                ],
                check=False,
            )
            adb.run(["reverse", "--remove", f"tcp:{ADB_PORT}"], check=False)
        except Exception:
            pass

        if self._relay:
            self._relay.stop()
            self._relay = None

    def _relay_status_callback(self, message: str, is_error: bool):
        self.log(message)
        if is_error:
            self.after(0, lambda: self._set_status_label(self.stream_status, message, "warn"))

    def open_virtual_driver_guide(self):
        self.log("Abrindo guia do driver virtual no navegador...")
        webbrowser.open(VIRTUAL_DRIVER_GUIDE)

    def on_close(self):
        if self._running:
            self._safe_stop_backend()
        self.destroy()


def main():
    app = MicMicStudioApp()
    app.mainloop()


if __name__ == "__main__":
    main()
