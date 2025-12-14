from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import threading
import time
import traceback
from typing import Any, Optional

if sys.platform != "win32":
    print("Sonos-Volume-Sync is only supported on Windows.")
    sys.exit(1)

try:
    from soco import SoCo  # type: ignore[import]
except ImportError:
    print("Missing dependency: soco\nInstall with: pip install soco")
    sys.exit(1)

try:
    import pystray  # type: ignore[import]
    from PIL import Image  # type: ignore[import]
except ImportError:
    print("Missing dependencies: pystray, pillow\nInstall with: pip install pystray pillow")
    sys.exit(1)

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # type: ignore[import]
    from comtypes import CLSCTX_ALL  # type: ignore[import]
    import comtypes  # type: ignore[import]
except ImportError:
    print("Missing dependencies: pycaw, comtypes\nInstall with: pip install pycaw")
    sys.exit(1)

try:
    import ctypes
    from ctypes import wintypes
except ImportError:
    ctypes = None  # type: ignore[assignment]

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "sonos_volume_sync.config.json")
LOG_FILE = os.path.join(BASE_DIR, "sonos_volume_sync.log")
MAX_LOG_SIZE = 1024 * 1024  # 1 MB

if ctypes is None:
    print("ctypes is required on Windows to register global hotkeys.")
    sys.exit(1)

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

WM_HOTKEY = 0x0312  # kept for completeness
WM_QUIT = 0x0012
WM_KEYDOWN = 0x0100
WM_SYSKEYDOWN = 0x0104
WH_KEYBOARD_LL = 13
HC_ACTION = 0
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF


class POINT(ctypes.Structure):
    _fields_ = [
        ("x", wintypes.LONG),
        ("y", wintypes.LONG),
    ]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", POINT),
    ]


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


# Hook state
LRESULT = ctypes.c_long
LowLevelKeyboardProc = ctypes.WINFUNCTYPE(
    LRESULT, wintypes.INT, wintypes.WPARAM, wintypes.LPARAM
)
_keyboard_proc: Optional[LowLevelKeyboardProc] = None
_hook_handle: Optional[int] = None
_hook_thread_id: Optional[int] = None


# Win32 API signatures (for correct 32/64-bit behaviour)
user32.SetWindowsHookExW.restype = ctypes.c_void_p
user32.SetWindowsHookExW.argtypes = [
    wintypes.INT,
    LowLevelKeyboardProc,
    wintypes.HINSTANCE,
    wintypes.DWORD,
]

user32.CallNextHookEx.restype = LRESULT
user32.CallNextHookEx.argtypes = [
    ctypes.c_void_p,
    wintypes.INT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]

user32.GetMessageW.restype = wintypes.BOOL
user32.GetMessageW.argtypes = [
    ctypes.POINTER(MSG),
    wintypes.HWND,
    wintypes.UINT,
    wintypes.UINT,
]

user32.TranslateMessage.restype = wintypes.BOOL
user32.TranslateMessage.argtypes = [ctypes.POINTER(MSG)]

user32.DispatchMessageW.restype = LRESULT
user32.DispatchMessageW.argtypes = [ctypes.POINTER(MSG)]

user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]

kernel32.GetModuleHandleW.restype = wintypes.HMODULE
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]

kernel32.GetCurrentThreadId.restype = wintypes.DWORD
kernel32.GetCurrentThreadId.argtypes = []


def load_config() -> dict[str, Any]:
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as exc:
        print(f"Warning: Failed to read config file '{CONFIG_FILE}': {exc}")
        return {}


_CONFIG = load_config()

# Configuration (defaults closely match original AHK script,
# but debug logging is enabled by default to aid troubleshooting)
try:
    PINNED_VOLUME: int = int(_CONFIG.get("pinned_volume", 33))
except (TypeError, ValueError):
    PINNED_VOLUME = 33
PINNED_VOLUME = max(0, min(100, PINNED_VOLUME))
NOTIFICATIONS_ENABLED: bool = bool(_CONFIG.get("notifications_enabled", False))
DEBUG_LOGGING: bool = bool(_CONFIG.get("debug_logging", True))

# Volume Control Configuration
VOLUME_STEP: float = float(_CONFIG.get("volume_step", 1.0))
USE_EXPONENTIAL: bool = bool(_CONFIG.get("use_exponential", True))
EXPONENTIAL_FACTOR: float = float(_CONFIG.get("exponential_factor", 10.0))

# Sonos / SoCo configuration
SONOS_IP: str = str(_CONFIG.get("sonos_ip", os.getenv("SONOS_IP", "")))
SONOS_NAME: str = str(_CONFIG.get("sonos_name", os.getenv("SONOS_NAME", "Sonos Five")))
WINDOWS_DEVICE_NAME: str = str(_CONFIG.get("windows_device_name", SONOS_NAME))


# Global state
g_device_name: str = ""
g_last_action: str = ""

current_sonos_volume: float = 0.1
_last_volume_value: float = 0.0

_volume_steps: float = 0.0
_volume_steps_lock = threading.Lock()
_volume_timer: Optional[threading.Timer] = None
_sonos_device_lock = threading.Lock()
_sonos_device: Optional[Any] = None
_shutdown_event = threading.Event()
_tray_icon: Optional[Any] = None
_volume_monitor_thread_started = False
_last_mute_state: Optional[bool] = None

_last_device_check_ms: float = 0.0
_cached_device_name: str = ""

_last_key_press_ms: float = 0.0
DEBOUNCE_TIME_MS: float = 0.0
AGGREGATION_DELAY_SEC: float = 0.05
DEVICE_CACHE_TIME_MS: float = 200.0
WINDOWS_VOLUME_POLL_INTERVAL_SEC: float = 0.1
SONOS_STATE_POLL_INTERVAL_SEC: float = 0.5


def debug_log(message: str) -> None:
    if not DEBUG_LOGGING:
        return

    timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    log_entry = f"{timestamp} - {message}\n"

    try:
        if os.path.exists(LOG_FILE) and os.path.getsize(LOG_FILE) > MAX_LOG_SIZE:
            os.remove(LOG_FILE)
            with open(LOG_FILE, "w", encoding="utf-8") as f:
                f.write("Log file cleared due to size limit\n")

        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception:
        # Swallow logging errors to avoid interfering with main behaviour
        pass


def create_tray_icon() -> None:
    """
    Create a system tray icon with an Exit menu item.
    Runs the tray loop in a background thread so the main
    logic (keyboard hook) can continue uninterrupted.
    """
    global _tray_icon

    debug_log("Initializing tray icon.")

    icon_image = None

    # Try to locate the icon both in the frozen bundle (PyInstaller)
    # and next to the script/EXE.
    icon_candidates = []
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        icon_candidates.append(os.path.join(meipass, "sonos_volume_sync.ico"))
    icon_candidates.append(os.path.join(BASE_DIR, "sonos_volume_sync.ico"))

    try:
        for icon_path in icon_candidates:
            if os.path.exists(icon_path):
                icon_image = Image.open(icon_path)
                debug_log(f"Loaded tray icon from: {icon_path}")
                break
    except Exception as exc:
        debug_log(f"Failed to load tray icon image: {exc}")

    if icon_image is None:
        try:
            # Fallback: simple solid icon if .ico is missing
            icon_image = Image.new("RGB", (16, 16), "black")
            debug_log("Using fallback tray icon image.")
        except Exception as exc:
            debug_log(f"Failed to create fallback tray icon image: {exc}")
            return

    def on_exit(icon: Any, item: Any) -> None:
        debug_log("Tray icon exit clicked; shutting down.")
        _shutdown_event.set()
        if _hook_thread_id is not None:
            try:
                user32.PostThreadMessageW(_hook_thread_id, WM_QUIT, 0, 0)
            except Exception as exc_inner:
                debug_log(f"Failed to post WM_QUIT to hook thread: {exc_inner}")
        icon.stop()

    menu = pystray.Menu(pystray.MenuItem("Exit", on_exit))
    _tray_icon = pystray.Icon(
        "Sonos-Volume-Sync",
        icon_image,
        "Sonos-Volume-Sync",
        menu,
    )

    def _run_icon() -> None:
        debug_log("Starting tray icon event loop.")
        try:
            _tray_icon.run()  # type: ignore[union-attr]
        except Exception as exc:
            debug_log(f"Tray icon loop exited with error: {exc}")

    thread = threading.Thread(target=_run_icon, daemon=True)
    thread.start()


def install_keyboard_hook() -> None:
    global _keyboard_proc, _hook_handle, _hook_thread_id

    def low_level_keyboard_proc(nCode: int, wParam: int, lParam: int) -> int:
        if nCode == HC_ACTION and wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
            kb = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
            if kb.vkCode == VK_VOLUME_UP:
                debug_log("Low-level hook: Volume Up")
                handled = handle_volume_key("up")
                if handled:
                    return 1
            elif kb.vkCode == VK_VOLUME_DOWN:
                debug_log("Low-level hook: Volume Down")
                handled = handle_volume_key("down")
                if handled:
                    return 1
        return user32.CallNextHookEx(_hook_handle, nCode, wParam, lParam)

    _keyboard_proc = LowLevelKeyboardProc(low_level_keyboard_proc)
    h_instance = kernel32.GetModuleHandleW(None)
    _hook_handle = user32.SetWindowsHookExW(WH_KEYBOARD_LL, _keyboard_proc, h_instance, 0)
    if not _hook_handle:
        err = ctypes.get_last_error()
        debug_log(f"Failed to install low-level keyboard hook. GetLastError={err}")
        raise RuntimeError(f"Failed to install low-level keyboard hook. GetLastError={err}")

    # Remember the thread id of the hook thread for WM_QUIT
    _hook_thread_id = kernel32.GetCurrentThreadId()


def remove_keyboard_hook() -> None:
    global _hook_handle
    if _hook_handle:
        user32.UnhookWindowsHookEx(_hook_handle)
        _hook_handle = None


def show_message_box(text: str, title: str) -> None:
    if ctypes is None:
        print(f"{title}: {text}")
        return
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x00000010)
    except Exception:
        print(f"{title}: {text}")


def show_notification(message: str) -> None:
    if NOTIFICATIONS_ENABLED:
        print(f"Sonos-Volume-Sync: {message}")


def print_status() -> None:
    print("Sonos-Volume-Sync status")
    print(f"Config file: {CONFIG_FILE} ({'found' if os.path.exists(CONFIG_FILE) else 'missing'})")
    print(f"Sonos name: {SONOS_NAME}")
    print(f"Windows device name match: {WINDOWS_DEVICE_NAME}")
    print(f"Sonos IP: {SONOS_IP if SONOS_IP else '(discovery)'}")
    print(f"Pinned Windows volume: {PINNED_VOLUME}%")
    print(f"Volume step: {VOLUME_STEP} (exponential={USE_EXPONENTIAL}, factor={EXPONENTIAL_FACTOR})")

    try:
        comtypes.CoInitialize()  # type: ignore[attr-defined]
    except Exception:
        pass

    try:
        device = get_default_playback_device() or "(unknown)"
        print(f"Default playback device: {device}")
        print(f"Matches configured device: {is_sonos_five_active()}")

        win_vol = get_windows_volume_percent()
        if win_vol is None:
            print("Windows volume: (unavailable)")
        else:
            print(f"Windows volume: {win_vol}%")

        sonos = get_sonos_device()
        if sonos is None:
            print("Sonos reachable: no")
        else:
            player_name = getattr(sonos, "player_name", "(unknown)")
            ip = getattr(sonos, "ip_address", "(unknown)")
            print(f"Sonos reachable: yes ({player_name} @ {ip})")
            try:
                print(f"Sonos volume: {int(getattr(sonos, 'volume'))}%")
            except Exception:
                pass
    finally:
        try:
            comtypes.CoUninitialize()  # type: ignore[attr-defined]
        except Exception:
            pass




def get_default_playback_device() -> str:
    global _last_device_check_ms, _cached_device_name

    now_ms = time.monotonic() * 1000.0
    if now_ms - _last_device_check_ms < DEVICE_CACHE_TIME_MS and _cached_device_name:
        return _cached_device_name

    try:
        speakers = AudioUtilities.GetSpeakers()
        # FriendlyName is usually available on the interface
        device = speakers.FriendlyName
        if device:
            _cached_device_name = device
            _last_device_check_ms = now_ms
            return device
    except Exception as exc:
        debug_log(f"Error getting audio device via pycaw: {exc}\n{traceback.format_exc()}")
        if not _cached_device_name:
            show_notification(f"Error getting audio device: {exc}")
        return _cached_device_name


def set_volume(volume: int) -> bool:
    try:
        speakers = AudioUtilities.GetSpeakers()
        endpoint = speakers.EndpointVolume
        # volume is 0-100, Scalar is 0.0-1.0
        endpoint.SetMasterVolumeLevelScalar(float(volume) / 100.0, None)
        return True
    except Exception as exc:
        debug_log(f"Failed to set volume via pycaw: {exc}\n{traceback.format_exc()}")
        show_notification(f"Failed to set volume: {exc}")
        return False


def set_playback_mute(mute: bool) -> bool:
    """
    Mute or unmute the default playback device using Core Audio (pycaw).
    This affects the Windows "Sonos Five" device when it is the default.
    """
    try:
        speakers = AudioUtilities.GetSpeakers()
        endpoint = speakers.EndpointVolume  # pycaw AudioDevice.EndpointVolume
        endpoint.SetMute(bool(mute), None)
        debug_log(f"Set playback mute via Core Audio to {mute}.")
        return True
    except Exception as exc:
        debug_log(f"Failed to set playback mute via Core Audio: {exc}\n{traceback.format_exc()}")
        show_notification(f"Failed to set playback mute: {exc}")
        return False


def get_windows_volume_percent() -> Optional[int]:
    """
    Read the current Windows master playback volume (0-100) using Core Audio.
    Returns None if the volume cannot be read.
    """
    try:
        # Use default speakers / playback device
        speakers = AudioUtilities.GetSpeakers()
        endpoint = speakers.EndpointVolume  # pycaw AudioDevice.EndpointVolume
        scalar = float(endpoint.GetMasterVolumeLevelScalar())
        percent = int(round(scalar * 100.0))
        return max(0, min(100, percent))
    except Exception as exc:
        debug_log(f"Failed to read Windows volume via Core Audio: {exc}")
        return None


def is_sonos_five_active() -> bool:
    global g_device_name
    device = get_default_playback_device()
    g_device_name = device
    return WINDOWS_DEVICE_NAME.lower() in device.lower()


def handle_external_volume_change(delta_percent: int) -> None:
    """
    Handle a change in Windows master volume (in percent points) that did not
    come from our intercepted hardware volume keys.
    Positive delta means volume up, negative means volume down.
    """
    if delta_percent == 0:
        return

    direction = "up" if delta_percent > 0 else "down"
    debug_log(f"External Windows volume change detected: {delta_percent:+d}% ({direction})")

    # Map the Windows volume delta to a relative Sonos volume change.
    try:
        current_volume = get_sonos_volume()
        delta_fraction = float(delta_percent) / 100.0
        new_volume = max(0.0, min(1.0, current_volume + delta_fraction))

        device = get_sonos_device()
        if device is None:
            debug_log("Cannot handle external volume change: no reachable Sonos device.")
            return

        target_percent = int(round(new_volume * 100.0))
        device.volume = target_percent

        global current_sonos_volume, _last_volume_value
        current_sonos_volume = new_volume
        _last_volume_value = new_volume

        debug_log(
            f"Applied external Windows volume delta {delta_percent:+d}% "
            f"-> Sonos volume set to {target_percent}%"
        )
        show_notification(f"Lautstärke: {target_percent}%")
    except Exception as exc:
        debug_log(f"Error handling external volume change: {exc}\n{traceback.format_exc()}")
        show_notification(f"Fehler bei externer Lautstärkeänderung: {exc}")
    finally:
        # Always reset Windows volume back to pinned level (guard band)
        set_volume(PINNED_VOLUME)


def volume_monitor_thread() -> None:
    """
    Background thread that polls Windows master volume and forwards any changes
    (when Sonos Five is active) to the Sonos speaker.
    """
    debug_log("Starting Windows volume monitor thread.")

    # Initialize COM on this background thread for Core Audio access.
    try:
        comtypes.CoInitialize()  # type: ignore[attr-defined]
    except Exception as exc:
        debug_log(f"Failed to CoInitialize COM in volume monitor thread: {exc}")
        return

    try:
        last_sonos_poll: float = 0.0
        while not _shutdown_event.is_set():
            try:
                global _last_mute_state

                if not is_sonos_five_active():
                    # If Sonos Five is no longer the active device but we
                    # previously muted, unmute to restore normal behaviour.
                    if _last_mute_state:
                        if set_playback_mute(False):
                            _last_mute_state = False
                    time.sleep(WINDOWS_VOLUME_POLL_INTERVAL_SEC)
                    continue

                # 1) Monitor Windows volume and treat deviations as external changes.
                current_percent = get_windows_volume_percent()
                if current_percent is None:
                    time.sleep(WINDOWS_VOLUME_POLL_INTERVAL_SEC)
                    continue

                if current_percent != PINNED_VOLUME:
                    delta = current_percent - PINNED_VOLUME
                    handle_external_volume_change(delta)

                # 2) Periodically inspect Sonos playback state to decide mute.
                now = time.monotonic()
                if now - last_sonos_poll >= SONOS_STATE_POLL_INTERVAL_SEC:
                    last_sonos_poll = now
                    device = get_sonos_device()
                    if device is not None:
                        try:
                            transport_info = device.get_current_transport_info()
                            state = transport_info.get("current_transport_state", "")

                            media_info = device.get_current_media_info()
                            channel = media_info.get("channel", "")

                            # Mute when Sonos is actively PLAYING and not in Line-In mode,
                            # otherwise unmute.
                            should_mute = state == "PLAYING" and channel != "Line-In"

                            if _last_mute_state is None or _last_mute_state != should_mute:
                                if set_playback_mute(should_mute):
                                    _last_mute_state = should_mute
                        except Exception as exc_sonos:
                            debug_log(
                                f"Error reading Sonos playback state in volume monitor: "
                                f"{exc_sonos}\n{traceback.format_exc()}"
                            )
            except Exception as exc:
                debug_log(f"Error in volume monitor thread: {exc}\n{traceback.format_exc()}")

            time.sleep(WINDOWS_VOLUME_POLL_INTERVAL_SEC)
    finally:
        try:
            comtypes.CoUninitialize()  # type: ignore[attr-defined]
        except Exception as exc:
            debug_log(f"Failed to CoUninitialize COM in volume monitor thread: {exc}")


def get_sonos_device() -> Optional[Any]:
    global _sonos_device

    with _sonos_device_lock:
        if _sonos_device is not None:
            return _sonos_device

        # 1. Try IP if configured
        if SONOS_IP:
            try:
                device = SoCo(SONOS_IP)
                # Probe the device to ensure it is reachable
                _ = device.volume
                _sonos_device = device
                debug_log(f"Connected to Sonos speaker at {SONOS_IP}")
                return device
            except Exception as exc:
                debug_log(f"Failed to connect to Sonos at {SONOS_IP}: {exc}")
                return None

        # 2. Try Name if configured (Discovery)
        if SONOS_NAME:
            debug_log(f"Discovering Sonos devices (looking for name='{SONOS_NAME}')...")
            import soco
            try:
                # Discover devices (timeout=5 seconds default, can be adjusted)
                devices = soco.discover()
                if not devices:
                    debug_log("Discovery returned no devices.")
                    return None
                
                matched_devices = []
                for dev in devices:
                    if dev.player_name.lower() == SONOS_NAME.lower():
                        matched_devices.append(dev)
                
                if not matched_devices:
                    debug_log(f"No devices found with name '{SONOS_NAME}'. Found: {[d.player_name for d in devices]}")
                    return None

                # Find the best target:
                # If we have a coordinator in the matched list, prefer it.
                # If not, take the first match and find its coordinator.
                
                target_device = None
                
                # First pass: check for coordinator in matches
                for dev in matched_devices:
                    if dev.is_coordinator:
                        target_device = dev
                        break
                
                # Second pass: if no coordinator matched directly, pick the first match and use its group coordinator
                if target_device is None:
                    first_match = matched_devices[0]
                    debug_log(f"Found matched device '{first_match.player_name}' at {first_match.ip_address} (not coordinator). Finding group coordinator...")
                    try:
                        target_device = first_match.group.coordinator
                    except Exception as e:
                        debug_log(f"Failed to resolve group coordinator: {e}")
                        target_device = first_match # Fallback

                if target_device:
                    _sonos_device = target_device
                    debug_log(f"Connected to discovered Sonos speaker: {target_device.player_name} at {target_device.ip_address} (Coordinator: {target_device.is_coordinator})")
                    return target_device
                    
            except Exception as exc:
                debug_log(f"Error during Sonos discovery: {exc}")
                return None
        
        debug_log("No Sonos IP or Name configured, or discovery failed.")
        return None


def get_sonos_volume() -> float:
    global current_sonos_volume, _last_volume_value

    device = None
    try:
        device = get_sonos_device()
    except Exception as exc:
        debug_log(f"Error obtaining Sonos device: {exc}")

    if device is None:
        debug_log("Could not obtain Sonos device; using last known volume value.")
        return _last_volume_value

    try:
        volume_percent = float(device.volume)
        volume = max(0.0, min(1.0, volume_percent / 100.0))

        _last_volume_value = volume
        current_sonos_volume = volume
        debug_log(f"Successfully retrieved volume via SoCo: {volume * 100:.0f}%")
        return volume
    except Exception as exc:
        debug_log(f"Error in get_sonos_volume via SoCo: {exc}")
        return _last_volume_value


def send_http_request(service: str, data: str) -> None:
    debug_log(
        f"send_http_request called for service='{service}', "
        "but Home Assistant integration has been removed."
    )
    return


def send_sonos_command(service: str, steps: float = 1.0) -> bool:
    global current_sonos_volume, _last_volume_value

    device = None
    try:
        device = get_sonos_device()
    except Exception as exc:
        debug_log(f"Error obtaining Sonos device: {exc}")

    if device is None:
        debug_log("Cannot send Sonos command: no reachable Sonos device.")
        return False

    try:
        step_size = max(0.1, float(steps))
        current_percent = float(device.volume)

        if service == "volume_up":
            new_percent = min(100.0, current_percent + step_size)
        elif service == "volume_down":
            new_percent = max(0.0, current_percent - step_size)
        else:
            debug_log(f"Unknown Sonos service: {service}")
            return False

        device.volume = int(round(new_percent))

        volume = new_percent / 100.0
        current_sonos_volume = volume
        _last_volume_value = volume

        debug_log(
            f"Set Sonos volume via SoCo to {new_percent:.0f}% "
            f"(service={service}, step_size={step_size:.2f})"
        )
        show_notification(f"Lautstärke: {round(current_sonos_volume * 100)}%")
        return True
    except Exception as exc:
        debug_log(f"Fehler beim Vorbereiten des Befehls (SoCo): {exc}\n{traceback.format_exc()}")
        show_notification(f"Fehler beim Vorbereiten des Befehls: {exc}")
        return False


def process_volume_steps() -> None:
    global _volume_steps, _volume_timer

    with _volume_steps_lock:
        steps = _volume_steps
        _volume_steps = 0.0
        _volume_timer = None

    if steps == 0:
        return

    try:
        direction = 1 if steps > 0 else -1
        step_size = min(abs(steps), 10.0)
        final_step = step_size * direction

        debug_log(f"Verarbeite Lautstärkeänderung: {final_step}%")

        service = "volume_up" if direction > 0 else "volume_down"
        send_sonos_command(service, step_size)
    except Exception as exc:
        debug_log(f"Fehler in process_volume_steps: {exc}")


def handle_volume_key(action: str) -> bool:
    global _last_key_press_ms, _volume_steps, _volume_timer, g_last_action

    if not is_sonos_five_active():
        debug_log("Volume key pressed but Sonos Five is not active; letting OS handle it.")
        return False

    now_ms = time.monotonic() * 1000.0
    if now_ms - _last_key_press_ms < DEBOUNCE_TIME_MS:
        # Debounced, but Sonos is active; swallow the key
        return True
    _last_key_press_ms = now_ms

    g_last_action = action

    current_volume = get_sonos_volume()

    if USE_EXPONENTIAL:
        volume_multiplier = 1.0 + (current_volume * (EXPONENTIAL_FACTOR - 1.0))
        step_multiplier = volume_multiplier
        step = max(VOLUME_STEP, VOLUME_STEP * step_multiplier)

        debug_log(
            "Exp. Steuerung - Aktuell: "
            f"{current_volume * 100:.0f}%, "
            f"Faktor: {step_multiplier:.2f}, "
            f"Schritt: {step:.2f}%"
        )
    else:
        step = VOLUME_STEP
        debug_log(f"Lineare Steuerung - Schritt: {step}%")

    if action == "down":
        step = -step

    with _volume_steps_lock:
        _volume_steps += step

        if _volume_timer is not None:
            _volume_timer.cancel()

        _volume_timer = threading.Timer(AGGREGATION_DELAY_SEC, process_volume_steps)
        _volume_timer.daemon = True
        _volume_timer.start()

    # Reset local volume back to pinned level (guard band)
    set_volume(PINNED_VOLUME)
    return True


def main() -> None:
    debug_log("Starting Sonos-Volume-Sync (Python)")

    parser = argparse.ArgumentParser(prog="sonos_volume_sync")
    parser.add_argument(
        "--status",
        action="store_true",
        help="Print audio device + Sonos connectivity status and exit.",
    )
    args = parser.parse_args()

    if args.status:
        print_status()
        return

    print("Sonos-Volume-Sync (Python) is running. Press Ctrl+C to exit.")

    def hook_thread() -> None:
        try:
            comtypes.CoInitialize()
        except Exception as exc:
            debug_log(f"Failed to CoInitialize in hook_thread: {exc}")

        try:
            install_keyboard_hook()
        except Exception as exc:
            debug_log(f"Failed to install keyboard hook: {exc}\n{traceback.format_exc()}")
            print(f"Failed to install keyboard hook: {exc}")
            return

        msg = MSG()
        while True:
            result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if result == 0:  # WM_QUIT
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

        remove_keyboard_hook()
        try:
            comtypes.CoUninitialize()
        except:
            pass

    thread = threading.Thread(target=hook_thread, daemon=True)
    thread.start()

    global _volume_monitor_thread_started
    if not _volume_monitor_thread_started:
        monitor_thread = threading.Thread(target=volume_monitor_thread, daemon=True)
        monitor_thread.start()
        _volume_monitor_thread_started = True

    # Start tray icon (runs in its own background thread)
    create_tray_icon()

    try:
        while not _shutdown_event.is_set():
            time.sleep(1.0)
    except KeyboardInterrupt:
        print("Sonos-Volume-Sync stopped.")
        # Ask hook thread to exit
        if _hook_thread_id is not None:
            user32.PostThreadMessageW(_hook_thread_id, WM_QUIT, 0, 0)
        _shutdown_event.set()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        # Catch any last-resort errors and ensure they are logged
        debug_log(f"Unhandled error in main: {exc}\n{traceback.format_exc()}")
        raise
