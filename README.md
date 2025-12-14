# 🎵 Sonos-Volume-Sync (Windows + SoCo)

Small Windows helper that intercepts the hardware volume keys when a **Sonos Five** is the default playback device and forwards them directly to the Sonos speaker (via [SoCo](https://github.com/SoCo/SoCo)) instead of changing the local Windows volume.

The goal is to keep the PC volume fixed (e.g. at 33%) while adjusting the Sonos speaker volume via the Sonos HTTP/UPnP API (through SoCo).

## 📂 Project Layout

- `sonos_volume_sync.py` – Main Python script that owns all runtime behaviour.
- `sonos_volume_sync.config.example.json` – Example configuration file. Copy this to `sonos_volume_sync.config.json` and edit it.
- `requirements.txt` – Python dependencies (`soco`, `pycaw`, `pystray`, `Pillow`).
- `sonos_volume_sync.ico` – Tray / application icon (for EXE packaging).
- `build_exe.bat` – Helper script to build a standalone `.exe` using PyInstaller.


## 🚀 What the App Does

- Intercepts hardware volume keys when a **Sonos** speaker (e.g., "Sonos Five") is the default playback device.
- Keeps the Windows playback volume locked at a configured level (`pinned_volume`).
- Forwards volume changes directly to the Sonos speaker via [SoCo](https://github.com/SoCo/SoCo) (using its discovery or IP).
- Supports linear or exponential volume step calculation and short-term aggregation of rapid key presses.
- Also watches Windows master volume changes (e.g. via the volume slider, mouse wheel tools like Volume2, etc.) and forwards those deltas to the Sonos speaker.
- When another device is the default, volume keys work normally on Windows.

In other words there are two “inputs”:

- **Volume key presses** (global keyboard hook): volume up/down is captured while the Sonos device is active and translated into Sonos volume steps.
- **Volume changes** (Windows volume monitor): Windows volume is polled; any deviation from `pinned_volume` is treated as a relative change and also forwarded to Sonos. After forwarding, Windows volume is reset back to `pinned_volume`.

Important: don’t set `pinned_volume` to `100`, otherwise Windows can’t go “higher” and the monitor can’t observe volume-up deltas.

## 📋 Requirements

- Windows 10/11 (desktop, with hardware volume keys).
- Python 3.11+ (installed and on `PATH`).
- Sonos speaker reachable on your local network (discovery by name or fixed IP).

## ⚙️ Configuration (`sonos_volume_sync.config.json`)

The Python script reads `sonos_volume_sync.config.json` at startup. An example configuration (`sonos_volume_sync.config.example.json`) with common settings is provided. Copy this file to `sonos_volume_sync.config.json` and adjust the values as needed.

```jsonc
{
  "sonos_name": "Sonos Five",
  "windows_device_name": "Sonos Five",
  "pinned_volume": 33,
  "volume_step": 1.0,
  "use_exponential": true,
  "exponential_factor": 10.0,
  "notifications_enabled": false,
  "debug_logging": false
}
```

Notes:

- `sonos_name`: The exact name of your Sonos speaker in the Sonos network (e.g., "Sonos Five"). The script will attempt to discover it.
- `windows_device_name`: The name of the playback device in Windows (e.g. "Sonos Five"). This is used to detect if the Sonos speaker is currently the default audio device.
- `sonos_ip`: (Optional) If discovery by `sonos_name` is unreliable, you can specify the IP address of your Sonos speaker directly. Using a static IP or DHCP reservation is recommended.
- `pinned_volume`: Controls the Windows volume "guard band" (0-100). Windows volume is kept near this level while Sonos is active.
- `pinned_volume` in practice: for my setup (Sonos Five connected via 3.5mm line-in), values above ~33% start to clip/overdrive the input signal, so I keep it around 33%.
- `debug_logging`: Set to `true` to enable `sonos_volume_sync.log` for troubleshooting.

## 🧩 Compatibility

This project was developed for my Windows PC and is **Windows-only as-is** (it uses Windows-specific APIs for global key capture and audio device/volume access). Since the core logic is Python, it should be possible to port to Linux/macOS by replacing those platform-specific parts, but this has **not** been tested.

## 🤖 Background (Why / How)

I always missed a small tool like this in my setup, but never took the time to properly build it. For a long time I ran a very simple **AutoHotkey** script instead.

This time I used the opportunity to implement it with **OpenAI Codex** and **Google Gemini**. The code in this repository is therefore **fully AI-generated**.

## 🐍 Installing Python Dependencies

From the project root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

If you prefer not to activate the venv:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## ▶️ Running the Python Version

From the project directory:

```powershell
.\.venv\Scripts\python.exe sonos_volume_sync.py
```

Quick status check (no keyboard hook / tray icon):

```powershell
.\.venv\Scripts\python.exe sonos_volume_sync.py --status
```

Behaviour to verify:

- With your **Sonos speaker** as the default playback device:
  - Pressing volume up/down or changing the system volume (mouse wheel/slider) should leave Windows volume at (about) the pinned level (`pinned_volume`).
  - Sonos volume should change directly on the speaker.
- With another device as default playback device:
  - Volume keys should work normally on Windows.

If `debug_logging` is enabled, inspect `sonos_volume_sync.log` for details about device detection and SoCo interactions.

## 📦 Building the Standalone EXE

To create a standalone `sonos_volume_sync.exe` using PyInstaller:

```powershell
.\build_exe.bat
```

This will generate the executable in the `dist/` folder.

## 💡 Personal Usage Tip (Deployment)

I personally use the tool by compiling the standalone EXE and placing it together with the `sonos_volume_sync.config.json` in a permanent folder. 

I then use **Windows Task Scheduler** to automatically start `sonos_volume_sync.exe` at system login. This ensures it runs silently in the background every time I start my PC. If I ever need to stop it manually, I can simply right-click the small tray icon and select "Exit".
