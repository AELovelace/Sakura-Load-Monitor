# Sakura Load Monitor

Desktop telemetry monitor built with PySide6 for AI workstation load checks.

## Features

- Per-GPU (NVIDIA/NVML + LibreHardwareMonitor when available):
  - GPU memory usage bar with used/total memory and percentage
  - Shared memory usage bar with used/total memory and percentage when exposed by the backend
  - GPU utilization bar
  - Core clock (MHz)
  - Power draw (Watts) when the driver exposes it
- CPU + System RAM:
  - CPU utilization bar
  - System RAM usage bar with used/total memory and percentage
  - CPU/RAM power from LibreHardwareMonitor when available (`N/A` otherwise)
- Runtime diagnostics card:
  - Active Python version/executable
  - Expected runtime status (.venv311)
  - NVML GPU count and LibreHardwareMonitor GPU count
  - LibreHardwareMonitor backend status
- Light Sakura visual style with pink progress bars and maroon accents

## Requirements

- Python 3.10+
- NVIDIA driver + NVML support for GPU metrics
- For CPU/RAM power on Windows: LibreHardwareMonitor library DLL

Install dependencies manually (optional):

```powershell
pip install -r requirements.txt
```

Run:

```powershell
python main.py
```

The app can bootstrap itself:

- If `.venv311` does not exist, it creates it with Python 3.11.
- It installs/updates dependencies from `requirements.txt`.
- It relaunches itself from `.venv311` automatically.
- If `lib/LibreHardwareMonitorLib.dll` is missing, it restores the pinned LibreHardwareMonitor `0.9.6` runtime from `lhm_netfx.zip` or downloads that archive automatically.

## Notes

- If no NVIDIA telemetry is available, the app still runs and shows CPU/RAM information.
- Python 3.14 does not currently support pythonnet in this setup; the app uses Python 3.11 runtime bootstrap for LibreHardwareMonitor support.
- CPU and RAM power are read via LibreHardwareMonitor when available. The app searches for `LibreHardwareMonitorLib.dll` in:
  - current folder (`./LibreHardwareMonitorLib.dll`)
  - `./lib/LibreHardwareMonitorLib.dll`
  - `%ProgramFiles%/LibreHardwareMonitor/LibreHardwareMonitorLib.dll`
  - `%ProgramFiles(x86)%/LibreHardwareMonitor/LibreHardwareMonitorLib.dll`
  - `%LOCALAPPDATA%/Programs/LibreHardwareMonitor/LibreHardwareMonitorLib.dll`
- If `./lib/LibreHardwareMonitorLib.dll` is missing, the app first tries to restore the pinned `0.9.6` DLL set from `./lhm_netfx.zip`; if that archive is also missing, it downloads the same pinned archive automatically.
- You can explicitly set a path with environment variable `LHM_DLL_PATH`.
- If the DLL or `pythonnet` is missing, CPU/RAM wattage will display as `N/A`.

## LibreHardwareMonitor Setup (Windows)

1. Download a LibreHardwareMonitor release ZIP from the project releases page.
2. Extract `LibreHardwareMonitorLib.dll`.
3. Place it in this project root or in the `lib` folder.
4. Run the app normally; backend status is shown in the app UI.
