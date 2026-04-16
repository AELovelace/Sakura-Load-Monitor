from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional
from urllib.error import URLError
from urllib.request import urlretrieve
import zipfile

import psutil
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QColor, QFont, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QProgressBar,
    QScrollArea,
    QTabWidget,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

try:
    import pynvml
except Exception:
    pynvml = None

try:
    import clr  # type: ignore
except Exception:
    clr = None


PROJECT_ROOT = Path(__file__).resolve().parent
LIB_DIR = PROJECT_ROOT / "lib"
LHM_VERSION = "0.9.6"
LHM_ARCHIVE_NAME = "lhm_netfx.zip"
LHM_ARCHIVE_URLS = [
    (
        "https://github.com/LibreHardwareMonitor/LibreHardwareMonitor/releases/download/"
        f"v{LHM_VERSION}/LibreHardwareMonitor-net472.zip"
    ),
    (
        "https://sourceforge.net/projects/librehardwaremonitor.mirror/files/"
        f"v{LHM_VERSION}/LibreHardwareMonitor-net472.zip/download"
    ),
]


@dataclass
class GPUStats:
    name: str
    vram_total_mib: int
    vram_used_mib: int
    vram_percent: float
    core_clock_mhz: Optional[int]
    util_percent: Optional[float]
    power_watts: Optional[float]
    temperature_c: Optional[float] = None
    shared_total_mib: int = 0
    shared_used_mib: int = 0
    shared_percent: float = 0.0


@dataclass
class SystemStats:
    cpu_percent: float
    cpu_power_watts: Optional[float]
    cpu_temp_c: Optional[float]
    ram_total_gib: float
    ram_used_gib: float
    ram_percent: float
    ram_power_watts: Optional[float]
    ram_temp_c: Optional[float]


class CoverBackgroundWidget(QWidget):
    def __init__(self, image_path: str) -> None:
        super().__init__()
        self.setObjectName("AppBody")
        self._background = QPixmap(image_path) if os.path.exists(image_path) else QPixmap()

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#FFF8FB"))

        if not self._background.isNull() and self.width() > 0 and self.height() > 0:
            scaled = self._background.scaled(
                self.size(),
                Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                Qt.TransformationMode.SmoothTransformation,
            )
            x = (self.width() - scaled.width()) // 2
            y = (self.height() - scaled.height()) // 2
            painter.drawPixmap(x, y, scaled)

        super().paintEvent(event)


def ensure_runtime_and_relaunch() -> bool:
    project_root = PROJECT_ROOT
    target_python = project_root / ".venv311" / "Scripts" / "python.exe"

    current = Path(sys.executable).resolve()
    if current == target_python.resolve() if target_python.exists() else False:
        return False

    if os.environ.get("SAKURA_BOOTSTRAPPED") == "1":
        return False

    try:
        if not target_python.exists():
            subprocess.run(
                ["py", "-3.11", "-m", "venv", ".venv311"],
                cwd=str(project_root),
                check=True,
            )

        subprocess.run(
            [str(target_python), "-m", "pip", "install", "--upgrade", "pip"],
            cwd=str(project_root),
            check=True,
        )
        subprocess.run(
            [str(target_python), "-m", "pip", "install", "-r", "requirements.txt"],
            cwd=str(project_root),
            check=True,
        )

        env = os.environ.copy()
        env["SAKURA_BOOTSTRAPPED"] = "1"
        subprocess.Popen([str(target_python), str(project_root / "main.py")], cwd=str(project_root), env=env)
        return True
    except Exception as exc:
        print(f"Bootstrap failed: {exc}")
        print("Install Python 3.11 and run from .venv311 to enable full telemetry.")
        return False


def ensure_lhm_lib_available() -> tuple[Optional[Path], Optional[str]]:
    target_dll = LIB_DIR / "LibreHardwareMonitorLib.dll"
    if target_dll.exists():
        return target_dll.resolve(), None

    archive_path = PROJECT_ROOT / LHM_ARCHIVE_NAME

    try:
        LIB_DIR.mkdir(parents=True, exist_ok=True)

        downloaded = False
        if not archive_path.exists():
            last_error: Optional[Exception] = None
            for archive_url in LHM_ARCHIVE_URLS:
                try:
                    urlretrieve(archive_url, archive_path)
                    downloaded = True
                    break
                except (URLError, OSError) as exc:
                    last_error = exc
            if not downloaded:
                raise last_error or OSError("download failed")

        extracted_any = False
        with zipfile.ZipFile(archive_path) as archive:
            for member in archive.infolist():
                member_path = Path(member.filename)
                if member.is_dir():
                    continue
                if len(member_path.parts) != 1:
                    continue
                if member_path.suffix.lower() != ".dll":
                    continue

                destination = LIB_DIR / member_path.name
                with archive.open(member) as src, destination.open("wb") as dst:
                    dst.write(src.read())
                extracted_any = True

        if target_dll.exists():
            source_label = "download" if downloaded else "cache"
            return target_dll.resolve(), f"LibreHardwareMonitor: restored {LHM_VERSION} from {source_label}"

        if extracted_any:
            return None, "LibreHardwareMonitor: archive extracted but LibreHardwareMonitorLib.dll was not found"
        return None, "LibreHardwareMonitor: archive did not contain any root-level DLLs"
    except (URLError, OSError, zipfile.BadZipFile) as exc:
        return None, f"LibreHardwareMonitor: auto-fetch failed ({exc})"


class NvidiaMonitor:
    def __init__(self) -> None:
        self.ready = False
        self.handles = []

        if pynvml is None:
            return

        try:
            pynvml.nvmlInit()
            count = pynvml.nvmlDeviceGetCount()
            for index in range(count):
                self.handles.append(pynvml.nvmlDeviceGetHandleByIndex(index))
            self.ready = bool(self.handles)
        except Exception:
            self.ready = False

    def collect(self) -> list[GPUStats]:
        if not self.ready:
            return []

        stats: list[GPUStats] = []
        for handle in self.handles:
            try:
                name = pynvml.nvmlDeviceGetName(handle)
                if isinstance(name, bytes):
                    name = name.decode("utf-8", "ignore")

                mem = pynvml.nvmlDeviceGetMemoryInfo(handle)
                vram_total_mib = int(mem.total / (1024 * 1024))
                vram_used_mib = int(mem.used / (1024 * 1024))
                vram_percent = (vram_used_mib / vram_total_mib * 100.0) if vram_total_mib else 0.0

                core_clock_mhz: Optional[int] = None
                try:
                    core_clock_mhz = int(pynvml.nvmlDeviceGetClockInfo(handle, pynvml.NVML_CLOCK_GRAPHICS))
                except Exception:
                    core_clock_mhz = None

                util_percent: Optional[float] = None
                try:
                    util = pynvml.nvmlDeviceGetUtilizationRates(handle)
                    util_percent = float(util.gpu)
                except Exception:
                    util_percent = None

                power_watts: Optional[float] = None
                try:
                    power_mw = pynvml.nvmlDeviceGetPowerUsage(handle)
                    power_watts = power_mw / 1000.0
                except Exception:
                    power_watts = None

                temperature_c: Optional[float] = None
                try:
                    temperature_c = float(
                        pynvml.nvmlDeviceGetTemperature(handle, pynvml.NVML_TEMPERATURE_GPU)
                    )
                except Exception:
                    temperature_c = None

                stats.append(
                    GPUStats(
                        name=str(name),
                        vram_total_mib=vram_total_mib,
                        vram_used_mib=vram_used_mib,
                        vram_percent=vram_percent,
                        core_clock_mhz=core_clock_mhz,
                        util_percent=util_percent,
                        power_watts=power_watts,
                        temperature_c=temperature_c,
                    )
                )
            except Exception:
                continue

        return stats

    def shutdown(self) -> None:
        if not self.ready:
            return
        try:
            pynvml.nvmlShutdown()
        except Exception:
            pass


def _find_lhm_dll() -> Optional[Path]:
    candidates: list[Path] = []

    env_path = os.environ.get("LHM_DLL_PATH", "").strip()
    if env_path:
        candidates.append(Path(env_path))

    candidates.extend(
        [
            Path("LibreHardwareMonitorLib.dll"),
            Path("lib") / "LibreHardwareMonitorLib.dll",
            Path(os.environ.get("ProgramFiles", "")) / "LibreHardwareMonitor" / "LibreHardwareMonitorLib.dll",
            Path(os.environ.get("ProgramFiles(x86)", "")) / "LibreHardwareMonitor" / "LibreHardwareMonitorLib.dll",
            Path(os.environ.get("LOCALAPPDATA", ""))
            / "Programs"
            / "LibreHardwareMonitor"
            / "LibreHardwareMonitorLib.dll",
        ]
    )

    for candidate in candidates:
        if str(candidate).strip() and candidate.exists():
            return candidate.resolve()
    return None


class LibreHardwareMonitorBridge:
    def __init__(self) -> None:
        self.available = False
        self.status = "LibreHardwareMonitor: unavailable"
        self._computer: Any = None
        self._sensor_type_power = None
        self._sensor_type_load = None
        self._sensor_type_clock = None
        self._sensor_type_data = None
        self._sensor_type_small_data = None
        self._sensor_type_temperature = None

        if os.name != "nt":
            self.status = "LibreHardwareMonitor: Windows only"
            return
        if clr is None:
            if sys.version_info >= (3, 14):
                self.status = "LibreHardwareMonitor: pythonnet unavailable on Python 3.14 (use Python 3.11-3.13)"
            else:
                self.status = "LibreHardwareMonitor: pythonnet not installed"
            return

        dll_path = _find_lhm_dll()
        auto_dll: Optional[Path] = None
        auto_status: Optional[str] = None
        if dll_path is None:
            auto_dll, auto_status = ensure_lhm_lib_available()
            dll_path = auto_dll or _find_lhm_dll()
        if dll_path is None:
            self.status = auto_status or "LibreHardwareMonitor: DLL not found"
            return

        try:
            dll_dir = dll_path.parent
            for dependency in sorted(dll_dir.glob("*.dll")):
                if dependency.name.lower() == "librehardwaremonitorlib.dll":
                    continue
                try:
                    clr.AddReference(str(dependency))
                except Exception:
                    # Some companion DLLs may target optional features.
                    pass

            clr.AddReference(str(dll_path))
            from LibreHardwareMonitor.Hardware import Computer, SensorType  # type: ignore

            self._sensor_type_power = SensorType.Power
            self._sensor_type_load = SensorType.Load
            self._sensor_type_clock = SensorType.Clock
            self._sensor_type_data = SensorType.Data
            self._sensor_type_small_data = SensorType.SmallData
            self._sensor_type_temperature = SensorType.Temperature

            computer = Computer()
            computer.IsCpuEnabled = True
            computer.IsGpuEnabled = True
            computer.IsMemoryEnabled = True
            computer.IsMotherboardEnabled = True
            computer.IsControllerEnabled = True
            computer.Open()

            self._computer = computer
            self.available = True
            if auto_status and auto_dll and dll_path.resolve() == auto_dll.resolve():
                self.status = f"{auto_status}; connected ({dll_path.name})"
            else:
                self.status = f"LibreHardwareMonitor: connected ({dll_path.name})"
        except Exception as exc:
            self.status = f"LibreHardwareMonitor: load failed ({exc})"

    def _iter_hardware(self):
        if self._computer is None:
            return

        queue = [item for item in self._computer.Hardware]
        while queue:
            hw = queue.pop(0)
            yield hw
            for sub in hw.SubHardware:
                queue.append(sub)

    def collect_cpu_ram_telemetry(
        self,
    ) -> tuple[Optional[float], Optional[float], Optional[float], Optional[float]]:
        if not self.available:
            return None, None, None, None

        cpu_candidates: list[float] = []
        ram_candidates: list[float] = []
        cpu_temp_candidates: list[float] = []
        ram_temp_candidates: list[float] = []

        try:
            for hw in self._iter_hardware():
                hw.Update()
                hw_type = str(hw.HardwareType).lower()

                for sensor in hw.Sensors:
                    if sensor.Value is None:
                        continue

                    sensor_name = str(sensor.Name).lower()
                    value = float(sensor.Value)

                    if sensor.SensorType == self._sensor_type_power:
                        if "cpu" in hw_type:
                            # Prefer package-level or total CPU power over individual rails/cores.
                            if any(k in sensor_name for k in ("package", "total", "cpu", "ppt")):
                                cpu_candidates.append(value)
                        elif "memory" in hw_type or "ram" in hw_type:
                            ram_candidates.append(value)

                        if any(k in sensor_name for k in ("dram", "dimm", "ram", "memory")):
                            ram_candidates.append(value)

                    elif sensor.SensorType == self._sensor_type_temperature:
                        if "cpu" in hw_type:
                            if any(k in sensor_name for k in ("package", "cpu", "die", "tdie", "tctl")):
                                cpu_temp_candidates.append(value)
                        elif "memory" in hw_type or "ram" in hw_type:
                            ram_temp_candidates.append(value)

                        if any(k in sensor_name for k in ("dram", "dimm", "ram", "memory")):
                            ram_temp_candidates.append(value)
        except Exception:
            return None, None, None, None

        cpu_power = max(cpu_candidates) if cpu_candidates else None
        ram_power = max(ram_candidates) if ram_candidates else None
        cpu_temp = max(cpu_temp_candidates) if cpu_temp_candidates else None
        ram_temp = max(ram_temp_candidates) if ram_temp_candidates else None
        return cpu_power, ram_power, cpu_temp, ram_temp

    @staticmethod
    def _data_value_to_mib(value: float, sensor_type: Any) -> int:
        # LHM Data is typically GiB, while SmallData GPU memory is typically MiB.
        if str(sensor_type) == "SmallData":
            return int(round(value))
        if value <= 512:
            return int(round(value * 1024))
        return int(round(value))

    def collect_gpu_stats(self) -> list[GPUStats]:
        if not self.available:
            return []

        stats: list[GPUStats] = []
        try:
            for hw in self._iter_hardware():
                hw.Update()
                hw_type = str(hw.HardwareType).lower()
                if "gpu" not in hw_type:
                    continue

                util_percent: Optional[float] = None
                core_clock_mhz: Optional[int] = None
                power_watts: Optional[float] = None
                temperature_c: Optional[float] = None
                mem_used_mib: Optional[int] = None
                mem_total_mib: Optional[int] = None
                mem_free_mib: Optional[int] = None
                dedicated_used_mib: Optional[int] = None
                dedicated_total_mib: Optional[int] = None
                dedicated_free_mib: Optional[int] = None
                shared_used_mib: Optional[int] = None
                shared_total_mib: Optional[int] = None
                shared_free_mib: Optional[int] = None

                for sensor in hw.Sensors:
                    if sensor.Value is None:
                        continue

                    sensor_name = str(sensor.Name).lower()
                    value = float(sensor.Value)
                    sensor_type = sensor.SensorType

                    if sensor_type == self._sensor_type_load:
                        if "core" in sensor_name or "gpu" in sensor_name:
                            util_percent = value if util_percent is None else max(util_percent, value)

                    elif sensor_type == self._sensor_type_clock:
                        if "core" in sensor_name or "graphics" in sensor_name:
                            mhz = int(round(value))
                            core_clock_mhz = mhz if core_clock_mhz is None else max(core_clock_mhz, mhz)

                    elif sensor_type == self._sensor_type_power:
                        if "total" in sensor_name or "package" in sensor_name or "gpu" in sensor_name:
                            power_watts = value if power_watts is None else max(power_watts, value)

                    elif sensor_type == self._sensor_type_temperature:
                        if "core" in sensor_name or "gpu" in sensor_name or "hotspot" in sensor_name:
                            temperature_c = value if temperature_c is None else max(temperature_c, value)

                    elif sensor_type in (self._sensor_type_data, self._sensor_type_small_data):
                        if "d3d dedicated memory used" in sensor_name:
                            dedicated_used_mib = self._data_value_to_mib(value, sensor_type)
                        elif "d3d dedicated memory total" in sensor_name:
                            dedicated_total_mib = self._data_value_to_mib(value, sensor_type)
                        elif "d3d dedicated memory free" in sensor_name:
                            dedicated_free_mib = self._data_value_to_mib(value, sensor_type)
                        elif "d3d shared memory used" in sensor_name:
                            shared_used_mib = self._data_value_to_mib(value, sensor_type)
                        elif "d3d shared memory total" in sensor_name:
                            shared_total_mib = self._data_value_to_mib(value, sensor_type)
                        elif "d3d shared memory free" in sensor_name:
                            shared_free_mib = self._data_value_to_mib(value, sensor_type)
                        elif "memory used" in sensor_name:
                            mem_used_mib = self._data_value_to_mib(value, sensor_type)
                        elif "memory total" in sensor_name:
                            mem_total_mib = self._data_value_to_mib(value, sensor_type)
                        elif "memory free" in sensor_name:
                            mem_free_mib = self._data_value_to_mib(value, sensor_type)

                if dedicated_total_mib is None and dedicated_used_mib is not None and dedicated_free_mib is not None:
                    dedicated_total_mib = dedicated_used_mib + dedicated_free_mib
                if shared_total_mib is None and shared_used_mib is not None and shared_free_mib is not None:
                    shared_total_mib = shared_used_mib + shared_free_mib
                if mem_total_mib is None and mem_used_mib is not None and mem_free_mib is not None:
                    mem_total_mib = mem_used_mib + mem_free_mib

                vram_total_mib = dedicated_total_mib if dedicated_total_mib is not None else (mem_total_mib or 0)
                vram_used_mib = dedicated_used_mib if dedicated_used_mib is not None else (mem_used_mib or 0)
                vram_percent = (vram_used_mib / vram_total_mib * 100.0) if vram_total_mib else 0.0
                shared_total = shared_total_mib or 0
                shared_used = shared_used_mib or 0
                shared_percent = (shared_used / shared_total * 100.0) if shared_total else 0.0

                stats.append(
                    GPUStats(
                        name=str(hw.Name),
                        vram_total_mib=vram_total_mib,
                        vram_used_mib=vram_used_mib,
                        vram_percent=vram_percent,
                        shared_total_mib=shared_total,
                        shared_used_mib=shared_used,
                        shared_percent=shared_percent,
                        core_clock_mhz=core_clock_mhz,
                        util_percent=util_percent,
                        power_watts=power_watts,
                        temperature_c=temperature_c,
                    )
                )
        except Exception:
            return []

        return stats

    def shutdown(self) -> None:
        if self._computer is None:
            return
        try:
            self._computer.Close()
        except Exception:
            pass


class MetricRow(QWidget):
    def __init__(self, title: str, with_bar: bool = True) -> None:
        super().__init__()
        self.title = QLabel(title)
        self.title.setObjectName("MetricTitle")
        self.value = QLabel("-")
        self.value.setObjectName("MetricValue")
        self.value.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.bar = QProgressBar()
        self.bar.setObjectName("SakuraBar")
        self.bar.setRange(0, 100)
        self.bar.setValue(0)
        self.bar.setTextVisible(False)
        self.bar.setFixedHeight(14)
        if not with_bar:
            self.bar.hide()

        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.addWidget(self.title)
        top.addWidget(self.value)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(4)
        root.addLayout(top)
        root.addWidget(self.bar)

    @staticmethod
    def _blend_channel(start: int, end: int, ratio: float) -> int:
        return int(round(start + (end - start) * ratio))

    @classmethod
    def _interpolate_color(cls, percent: float) -> QColor:
        p = max(0.0, min(100.0, percent))
        # Piecewise gradient: light blue at 0, pink at 45, red at 80+.
        if p <= 45.0:
            ratio = p / 45.0 if 45.0 else 0.0
            start = QColor("#A8E4FF")
            end = QColor("#FF8FC9")
        elif p <= 80.0:
            ratio = (p - 45.0) / 35.0
            start = QColor("#FF8FC9")
            end = QColor("#FF4D4D")
        else:
            ratio = (p - 80.0) / 20.0
            start = QColor("#FF4D4D")
            end = QColor("#D7263D")

        return QColor(
            cls._blend_channel(start.red(), end.red(), ratio),
            cls._blend_channel(start.green(), end.green(), ratio),
            cls._blend_channel(start.blue(), end.blue(), ratio),
        )

    def _apply_bar_color(self, percent: float) -> None:
        base = self._interpolate_color(percent)
        glow = base.lighter(125)
        self.bar.setStyleSheet(
            (
                "QProgressBar {"
                "border: 1px solid #DDB2C1;"
                "border-radius: 7px;"
                "background: #FFEAF1;"
                "}"
                "QProgressBar::chunk {"
                "border-radius: 7px;"
                "background: qlineargradient("
                "x1: 0, y1: 0, x2: 1, y2: 0,"
                f"stop: 0 {glow.name()},"
                f"stop: 1 {base.name()}"
                ");"
                "}"
            )
        )

    def set_percent(self, percent: float, text: str) -> None:
        bounded = max(0, min(100, int(round(percent))))
        self.bar.setValue(bounded)
        self._apply_bar_color(float(bounded))
        self.value.setText(text)

    def set_text(self, text: str) -> None:
        self.value.setText(text)


class GPUCard(QGroupBox):
    def __init__(self, gpu_index: int, gpu_name: str) -> None:
        super().__init__(f"GPU {gpu_index}: {gpu_name}")
        self.setObjectName("DeviceCard")

        self.vram_row = MetricRow("GPU Memory")
        self.shared_row = MetricRow("Shared Memory")
        self.util_row = MetricRow("GPU Utilization")
        self.temp_row = MetricRow("Temperature")
        self.clock_row = MetricRow("Core Clock", with_bar=False)
        self.power_row = MetricRow("Power Draw", with_bar=False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 16, 14, 14)
        layout.setSpacing(12)
        layout.addWidget(self.vram_row)
        layout.addWidget(self.shared_row)
        layout.addWidget(self.util_row)
        layout.addWidget(self.temp_row)
        layout.addWidget(self.clock_row)
        layout.addWidget(self.power_row)

    def apply_stats(self, stats: GPUStats) -> None:
        if stats.vram_total_mib <= 0:
            self.vram_row.set_percent(0, "N/A")
        else:
            self.vram_row.set_percent(
                stats.vram_percent,
                f"{stats.vram_used_mib:,} / {stats.vram_total_mib:,} MiB ({stats.vram_percent:.1f}%)",
            )

        if stats.shared_total_mib <= 0:
            self.shared_row.set_percent(0, "N/A")
        else:
            self.shared_row.set_percent(
                stats.shared_percent,
                f"{stats.shared_used_mib:,} / {stats.shared_total_mib:,} MiB ({stats.shared_percent:.1f}%)",
            )

        if stats.util_percent is None:
            self.util_row.set_percent(0, "N/A")
        else:
            self.util_row.set_percent(stats.util_percent, f"{stats.util_percent:.1f}%")

        if stats.temperature_c is None:
            self.temp_row.set_percent(0, "N/A")
        else:
            temp_percent = max(0.0, min(100.0, stats.temperature_c))
            self.temp_row.set_percent(temp_percent, f"{stats.temperature_c:.1f} C")

        if stats.core_clock_mhz is None:
            self.clock_row.set_text("N/A")
        else:
            self.clock_row.set_text(f"{stats.core_clock_mhz:,} MHz")

        if stats.power_watts is None:
            self.power_row.set_text("N/A")
        else:
            self.power_row.set_text(f"{stats.power_watts:.1f} W")


class SystemCard(QGroupBox):
    def __init__(self) -> None:
        super().__init__("CPU + System Memory")
        self.setObjectName("DeviceCard")

        self.cpu_row = MetricRow("CPU Utilization")
        self.cpu_temp_row = MetricRow("CPU Temperature")
        self.cpu_power_row = MetricRow("CPU Power", with_bar=False)
        self.ram_row = MetricRow("System RAM")
        self.ram_temp_row = MetricRow("RAM Temperature")
        self.ram_power_row = MetricRow("RAM Power", with_bar=False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 16, 14, 14)
        layout.setSpacing(12)
        layout.addWidget(self.cpu_row)
        layout.addWidget(self.cpu_temp_row)
        layout.addWidget(self.cpu_power_row)
        layout.addWidget(self.ram_row)
        layout.addWidget(self.ram_temp_row)
        layout.addWidget(self.ram_power_row)

    def apply_stats(self, stats: SystemStats) -> None:
        self.cpu_row.set_percent(stats.cpu_percent, f"{stats.cpu_percent:.1f}%")

        if stats.cpu_temp_c is None:
            self.cpu_temp_row.set_percent(0, "N/A")
        else:
            cpu_temp_percent = max(0.0, min(100.0, stats.cpu_temp_c))
            self.cpu_temp_row.set_percent(cpu_temp_percent, f"{stats.cpu_temp_c:.1f} C")

        if stats.cpu_power_watts is None:
            self.cpu_power_row.set_text("N/A")
        else:
            self.cpu_power_row.set_text(f"{stats.cpu_power_watts:.1f} W")

        self.ram_row.set_percent(
            stats.ram_percent,
            f"{stats.ram_used_gib:.2f} / {stats.ram_total_gib:.2f} GiB ({stats.ram_percent:.1f}%)",
        )

        if stats.ram_temp_c is None:
            self.ram_temp_row.set_percent(0, "N/A")
        else:
            ram_temp_percent = max(0.0, min(100.0, stats.ram_temp_c))
            self.ram_temp_row.set_percent(ram_temp_percent, f"{stats.ram_temp_c:.1f} C")

        if stats.ram_power_watts is None:
            self.ram_power_row.set_text("N/A")
        else:
            self.ram_power_row.set_text(f"{stats.ram_power_watts:.1f} W")


class DiagnosticsCard(QGroupBox):
    def __init__(self) -> None:
        super().__init__("Runtime Diagnostics")
        self.setObjectName("DeviceCard")

        self.python_row = MetricRow("Python", with_bar=False)
        self.runtime_row = MetricRow("Expected Runtime", with_bar=False)
        self.nvml_row = MetricRow("NVML GPUs", with_bar=False)
        self.lhm_row = MetricRow("LHM GPUs", with_bar=False)
        self.lhm_status_row = MetricRow("LHM Status", with_bar=False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 16, 14, 14)
        layout.setSpacing(12)
        layout.addWidget(self.python_row)
        layout.addWidget(self.runtime_row)
        layout.addWidget(self.nvml_row)
        layout.addWidget(self.lhm_row)
        layout.addWidget(self.lhm_status_row)

    def apply_stats(
        self,
        python_display: str,
        runtime_display: str,
        nvml_count: int,
        lhm_count: int,
        lhm_status: str,
    ) -> None:
        self.python_row.set_text(python_display)
        self.runtime_row.set_text(runtime_display)
        self.nvml_row.set_text(str(nvml_count))
        self.lhm_row.set_text(str(lhm_count))
        self.lhm_status_row.set_text(lhm_status)


class Dashboard(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Sakura Load Monitor")
        self.resize(980, 680)

        self.monitor = NvidiaMonitor()
        self.lhm = LibreHardwareMonitorBridge()
        self.gpu_cards: list[GPUCard] = []

        root = CoverBackgroundWidget("res/bg.jpg")
        self.setCentralWidget(root)

        container = QVBoxLayout(root)
        container.setContentsMargins(18, 18, 18, 18)
        container.setSpacing(14)

        header = QFrame()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(18, 12, 18, 12)

        title = QLabel("Sakura Load Monitor")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)

        subtitle = QLabel("GPU / CPU / RAM usage and power")
        subtitle.setObjectName("Subtitle")

        title_wrap = QVBoxLayout()
        title_wrap.setContentsMargins(0, 0, 0, 0)
        title_wrap.setSpacing(2)
        title_wrap.addWidget(title)
        title_wrap.addWidget(subtitle)

        self.updated_label = QLabel("Last update: --")
        self.updated_label.setObjectName("UpdatedLabel")
        self.updated_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        header_layout.addLayout(title_wrap)
        header_layout.addWidget(self.updated_label)
        container.addWidget(header)

        tabs = QTabWidget()
        tabs.setObjectName("MainTabs")
        container.addWidget(tabs)

        # Telemetry tab: system and per-GPU bars.
        telemetry_tab = QWidget()
        telemetry_root = QVBoxLayout(telemetry_tab)
        telemetry_root.setContentsMargins(0, 0, 0, 0)

        telemetry_scroll = QScrollArea()
        telemetry_scroll.setWidgetResizable(True)
        telemetry_scroll.setFrameShape(QFrame.Shape.NoFrame)
        telemetry_root.addWidget(telemetry_scroll)

        telemetry_body = QWidget()
        self.telemetry_layout = QVBoxLayout(telemetry_body)
        self.telemetry_layout.setContentsMargins(4, 4, 4, 8)
        self.telemetry_layout.setSpacing(12)
        telemetry_scroll.setWidget(telemetry_body)

        self.system_card = SystemCard()
        self.telemetry_layout.addWidget(self.system_card)

        self.no_gpu_label = QLabel("No GPU telemetry found from NVML or LibreHardwareMonitor backends.")
        self.no_gpu_label.setWordWrap(True)
        self.no_gpu_label.setObjectName("Hint")
        self.no_gpu_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        self.telemetry_layout.addWidget(self.no_gpu_label)

        tabs.addTab(telemetry_tab, "Telemetry")

        # Runtime Diagnostics tab: backend/runtime checks only.
        diagnostics_tab = QWidget()
        diagnostics_root = QVBoxLayout(diagnostics_tab)
        diagnostics_root.setContentsMargins(4, 4, 4, 8)
        diagnostics_root.setSpacing(12)

        self.diagnostics_card = DiagnosticsCard()
        diagnostics_root.addWidget(self.diagnostics_card)

        self.power_backend_label = QLabel(self.lhm.status)
        self.power_backend_label.setObjectName("Hint")
        self.power_backend_label.setWordWrap(True)
        self.power_backend_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        diagnostics_root.addWidget(self.power_backend_label)
        diagnostics_root.addStretch(1)

        tabs.addTab(diagnostics_tab, "Runtime Diagnostics")

        self._build_gpu_cards()
        self._apply_theme()

        self.timer = QTimer(self)
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self.refresh)
        self.timer.start()

        # Prime psutil so the first display update is meaningful.
        psutil.cpu_percent(interval=None)
        self.refresh()

    @staticmethod
    def _normalize_gpu_name(name: str) -> str:
        normalized = "".join(ch.lower() for ch in name if ch.isalnum() or ch.isspace())
        for token in ("nvidia", "amd", "radeon", "geforce"):
            normalized = normalized.replace(token, " ")
        return " ".join(normalized.split())

    @staticmethod
    def _gpu_name_matches(left: str, right: str) -> bool:
        if left == right:
            return True
        if not left or not right:
            return False
        return left in right or right in left

    @staticmethod
    def _merge_gpu_stats(nvml_stats: list[GPUStats], lhm_stats: list[GPUStats]) -> list[GPUStats]:
        by_name: dict[str, GPUStats] = {}

        def norm(name: str) -> str:
            normalized = "".join(ch.lower() for ch in name if ch.isalnum() or ch.isspace())
            for token in ("nvidia", "amd", "radeon", "geforce"):
                normalized = normalized.replace(token, " ")
            return " ".join(normalized.split())

        def name_matches(left: str, right: str) -> bool:
            if left == right:
                return True
            if not left or not right:
                return False
            return left in right or right in left

        for item in nvml_stats:
            by_name[norm(item.name)] = item

        for item in lhm_stats:
            key = norm(item.name)
            existing = by_name.get(key)
            if existing is None:
                for existing_key, existing_value in by_name.items():
                    if name_matches(existing_key, key):
                        existing = existing_value
                        key = existing_key
                        break
            if existing is None:
                by_name[key] = item
                continue

            if existing.vram_total_mib <= 0 and item.vram_total_mib > 0:
                existing.vram_total_mib = item.vram_total_mib
                existing.vram_used_mib = item.vram_used_mib
                existing.vram_percent = item.vram_percent
            if existing.shared_total_mib <= 0 and item.shared_total_mib > 0:
                existing.shared_total_mib = item.shared_total_mib
                existing.shared_used_mib = item.shared_used_mib
                existing.shared_percent = item.shared_percent
            if existing.util_percent is None and item.util_percent is not None:
                existing.util_percent = item.util_percent
            if existing.core_clock_mhz is None and item.core_clock_mhz is not None:
                existing.core_clock_mhz = item.core_clock_mhz
            if existing.power_watts is None and item.power_watts is not None:
                existing.power_watts = item.power_watts
            if existing.temperature_c is None and item.temperature_c is not None:
                existing.temperature_c = item.temperature_c

        return list(by_name.values())

    def collect_gpu_sources(self) -> tuple[list[GPUStats], int, int]:
        nvml_stats = self.monitor.collect()
        lhm_stats = self.lhm.collect_gpu_stats()
        merged = self._merge_gpu_stats(nvml_stats, lhm_stats)
        return merged, len(nvml_stats), len(lhm_stats)

    def collect_gpu_stats(self) -> list[GPUStats]:
        merged, _, _ = self.collect_gpu_sources()
        return merged

    def _build_gpu_cards(self) -> None:
        for card in self.gpu_cards:
            card.setParent(None)
        self.gpu_cards.clear()

        gpu_stats = self.collect_gpu_stats()
        if not gpu_stats:
            self.no_gpu_label.show()
            return

        self.no_gpu_label.hide()

        for idx, stats in enumerate(gpu_stats):
            card = GPUCard(idx, stats.name)
            card.apply_stats(stats)
            self.gpu_cards.append(card)
            self.telemetry_layout.addWidget(card)

    def collect_system_stats(self) -> SystemStats:
        cpu_percent = float(psutil.cpu_percent(interval=None))

        mem = psutil.virtual_memory()
        ram_total_gib = mem.total / (1024 ** 3)
        ram_used_gib = (mem.total - mem.available) / (1024 ** 3)
        ram_percent = float(mem.percent)

        cpu_power_watts, ram_power_watts, cpu_temp_c, ram_temp_c = self.lhm.collect_cpu_ram_telemetry()

        return SystemStats(
            cpu_percent=cpu_percent,
            cpu_power_watts=cpu_power_watts,
            cpu_temp_c=cpu_temp_c,
            ram_total_gib=ram_total_gib,
            ram_used_gib=ram_used_gib,
            ram_percent=ram_percent,
            ram_power_watts=ram_power_watts,
            ram_temp_c=ram_temp_c,
        )

    def refresh(self) -> None:
        gpu_stats, nvml_count, lhm_count = self.collect_gpu_sources()

        if len(gpu_stats) != len(self.gpu_cards):
            self._build_gpu_cards()
            gpu_stats = self.collect_gpu_stats()

        for card, stats in zip(self.gpu_cards, gpu_stats):
            card.apply_stats(stats)

        self.system_card.apply_stats(self.collect_system_stats())
        project_root = PROJECT_ROOT
        target_python = project_root / ".venv311" / "Scripts" / "python.exe"
        self.diagnostics_card.apply_stats(
            python_display=f"{sys.version.split()[0]} ({Path(sys.executable).name})",
            runtime_display="OK" if Path(sys.executable).resolve() == target_python.resolve() else "Fallback",
            nvml_count=nvml_count,
            lhm_count=lhm_count,
            lhm_status=self.lhm.status,
        )
        self.updated_label.setText(f"Last update: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    def closeEvent(self, event) -> None:
        self.monitor.shutdown()
        self.lhm.shutdown()
        super().closeEvent(event)

    def _apply_theme(self) -> None:
        self.setStyleSheet(
            f"""
            QWidget {{
                color: #5A2433;
                font-family: 'Segoe UI', 'Noto Sans', sans-serif;
                font-size: 10.5pt;
                background: #FFF6FA;
            }}
            QMainWindow {{
                background: #FFF8FB;
            }}
            QWidget#AppBody {{
                background: #FFF8FB;
            }}
            QTabWidget#MainTabs::pane {{
                border: 1px solid #B06A7D;
                border-radius: 12px;
                top: -1px;
                background: rgba(255, 248, 251, 196);
            }}
            QTabBar::tab {{
                background: #6F2E42;
                color: #FFE8EF;
                border: 1px solid #5D2435;
                border-bottom: none;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                padding: 8px 16px;
                margin-right: 4px;
                font-weight: 600;
            }}
            QTabBar::tab:selected {{
                background: #F39AB5;
                color: #5A2433;
                border: 1px solid #D87093;
                border-bottom: 1px solid rgba(255, 248, 251, 196);
            }}
            QTabBar::tab:hover:!selected {{
                background: #7E374E;
            }}
            QFrame {{
                background: rgba(255, 245, 250, 224);
                border: 1px solid #E9C5D1;
                border-radius: 14px;
            }}
            QGroupBox#DeviceCard {{
                background: rgba(255, 247, 251, 236);
                border: 1px solid #E5BAC8;
                border-radius: 14px;
                margin-top: 8px;
                padding-top: 8px;
                font-weight: 600;
                color: #6B2A3E;
            }}
            QGroupBox#DeviceCard::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 6px;
            }}
            QLabel#Subtitle {{
                color: #8A5061;
            }}
            QLabel#UpdatedLabel {{
                color: #7B3A4D;
                font-weight: 600;
            }}
            QLabel#Hint {{
                color: #8F5F6F;
                background: rgba(255, 240, 246, 220);
                border: 1px dashed #D9A7B7;
                border-radius: 10px;
                padding: 10px;
            }}
            QLabel#MetricTitle {{
                font-weight: 600;
                color: #6A2A3F;
            }}
            QLabel#MetricValue {{
                color: #7A3A4E;
                font-weight: 600;
            }}
            QProgressBar#SakuraBar {{
                border: 1px solid #DDB2C1;
                border-radius: 7px;
                background: #FFEAF1;
            }}
            QProgressBar#SakuraBar::chunk {{
                border-radius: 7px;
                background: #F39AB5;
            }}
            QScrollArea {{
                border: none;
                background: transparent;
            }}
            """
        )


def main() -> int:
    if ensure_runtime_and_relaunch():
        return 0

    app = QApplication(sys.argv)
    win = Dashboard()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
