import sys
import winreg
import subprocess
import json
import time
import os
import datetime
from typing import List, Optional
from dataclasses import dataclass
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView, QGroupBox, QProgressBar,
    QSplashScreen, QToolBar, QDialog, QTextEdit, QCheckBox, QProgressDialog
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QFont, QColor, QPixmap, QPainter, QAction, QIcon

import resources  #:noqa: F401


startupinfo = None
if sys.platform == "win32":
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW


@dataclass
class CameraDevice:
    """Data class for USB camera information"""
    name: str
    device_id: str
    registry_path: str
    friendly_name: str
    hardware_id: str
    is_connected: bool = True


class RegistrySearchDialog(QDialog):
    """Dialog to show registry search progress"""

    def __init__(self, parent=None, search_options=None):
        super().__init__(parent)
        self.search_options = search_options or {}
        self.setWindowTitle("Registry Search Progress")
        self.setFixedSize(600, 400)
        self.setModal(True)

        layout = QVBoxLayout()

        # Title
        title = QLabel("🔍 Searching Registry Entries...")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        layout.addWidget(title)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)

        # Status label
        self.status_label = QLabel("Initializing search...")
        layout.addWidget(self.status_label)

        # Search results text area
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setMinimumHeight(250)
        layout.addWidget(self.results_text)

        # Close button
        self.close_button = QPushButton("OK")
        self.close_button.clicked.connect(self.accept)
        self.close_button.setEnabled(False)  # Disabled until search completes
        layout.addWidget(self.close_button)

        self.setLayout(layout)

        # Auto-close timer setup
        self.auto_close_timer = QTimer()
        self.auto_close_timer.timeout.connect(self.accept)
        self.countdown_timer = QTimer()
        self.countdown_timer.timeout.connect(self.update_countdown)
        self.countdown_seconds = 4

    def update_progress(self, value, status):
        """Update progress bar and status"""
        self.progress_bar.setValue(value)
        self.status_label.setText(status)

    def add_result(self, text):
        """Add text to results area"""
        self.results_text.append(text)

    def search_completed(self):
        """Called when search is completed"""
        self.close_button.setEnabled(True)
        self.status_label.setText("Search completed!")

        # Check if auto-close is enabled
        if self.search_options.get("skip_next_btn", False):
            self.start_auto_close()

    def start_auto_close(self):
        """Start the auto-close countdown"""
        self.countdown_seconds = 4
        self.update_countdown()
        self.countdown_timer.start(1000)  # Update every second
        self.auto_close_timer.start(4000)  # Close after 3 seconds

    def update_countdown(self):
        """Update the countdown display"""
        if self.countdown_seconds > 0:
            self.close_button.setText(f"OK (auto-close in {self.countdown_seconds}s)")
            self.countdown_seconds -= 1
        else:
            self.countdown_timer.stop()

class EnhancedRegistrySearchThread(QThread):
    """Enhanced thread for comprehensive registry search"""
    search_completed = Signal(list)
    progress_updated = Signal(int, str)
    result_found = Signal(str)

    def __init__(self, camera: CameraDevice, search_options: dict = None):
        super().__init__()
        self.camera = camera
        self.search_options = search_options or self.load_default_search_options()

    def load_default_search_options(self):
        """Load search options from settings or return defaults"""
        defaults = {
            "device_manager_friendly_name": False,
            "standard_device_paths": True,
            "device_classes": False,
            "usb_interfaces": False,
            "system_drivers": False,
            "control_entries": False,
            "powershell_extended": False,
            "vid_pid_matching": True,
            "friendly_name_search": False,
            "skip_next_btn": True
        }

        try:
            import json
            if os.path.exists("search_options.json"):
                with open("search_options.json", "r") as f:
                    saved_options = json.load(f)
                    defaults.update(saved_options)
        except Exception:
            pass

        return defaults

    def run(self):
        """Comprehensive registry search with detailed progress"""
        try:
            self.progress_updated.emit(0, "Starting registry search...")
            registry_paths = self.comprehensive_registry_search()
            self.search_completed.emit(registry_paths)
        except Exception as e:
            self.result_found.emit(f"❌ Error in registry search: {str(e)}")
            self.search_completed.emit([])

    def comprehensive_registry_search(self) -> List[str]:
        """Comprehensive registry search with configurable strategies"""
        registry_paths = []
        progress = 0
        total_steps = sum(1 for key, value in self.search_options.items() if value and key != 'vid_pid_matching' and key != 'friendly_name_search')
        step_increment = 90 / max(total_steps, 1)  # Reserve 10% for final processing

        # Flag für kontinuierliche Animation setzen
        self.animate_progress = True
        self.current_step_progress = 0
        self.target_progress = 0

        # Separater Worker-Thread für Animation (läuft parallel)
        def animation_worker():
            while self.animate_progress:
                # Sanfte Bewegung zum Zielwert
                if abs(self.current_step_progress - self.target_progress) > 0.1:
                    # Bewege sich langsam zum Ziel
                    diff = self.target_progress - self.current_step_progress
                    self.current_step_progress += diff * 0.1  # 10% pro Update

                    self.progress_updated.emit(int(self.current_step_progress), f"Arbeite... ({int(self.current_step_progress)}%)")
                else:
                    # Am Ziel angekommen - kleine zufällige Bewegung
                    import random
                    jitter = random.uniform(-0.3, 0.3)
                    display_progress = max(0, min(100, self.current_step_progress + jitter))

                    self.progress_updated.emit(int(display_progress), f"Verarbeite... ({int(self.current_step_progress)}%)")

                self.msleep(100)  # Update alle 100ms für sichtbare Animation

        # Animation-Thread starten
        import threading
        animation_thread = threading.Thread(target=animation_worker, daemon=True)
        animation_thread.start()

        def set_target_progress(target, status):
            """Setzt neuen Zielwert für Animation"""
            self.target_progress = float(target)
            # Sofortiges Update für Status-Text
            self.progress_updated.emit(int(self.current_step_progress), status)

        # Emit initial progress
        set_target_progress(5, "Initialisiere Registry-Suche...")
        # Extract VID/PID if matching is enabled
        vid_pid = ""
        if self.search_options.get("vid_pid_matching", True):
            vid_pid = self.extract_vid_pid()
            self.result_found.emit(f"🔍 VID/PID extraction: {'Enabled - ' + vid_pid if vid_pid else 'Enabled but no VID/PID found'}")
        else:
            self.result_found.emit("⏭️ VID/PID matching: Disabled (skipped)")

        # Step 0: Device Manager FriendlyName search
        if self.search_options.get("device_manager_friendly_name", True):
            progress += step_increment
            set_target_progress(int(progress), "Searching Device Manager FriendlyName entries...")
            dm_paths = self.search_device_manager_friendly_name(self.camera)
            registry_paths.extend(dm_paths)
        else:
            self.result_found.emit("⏭️ Device Manager FriendlyName: Disabled (skipped)")


        # Step 1: Standard device ID paths
        if self.search_options.get("standard_device_paths", True):
            progress += step_increment
            set_target_progress(int(progress), "Searching standard device paths...")
            standard_paths = self.search_standard_device_paths(self.camera.device_id.replace("\\", "#"))
            registry_paths.extend(standard_paths)
        else:
            self.result_found.emit("⏭️ Standard Device Paths: Disabled (skipped)")

        # Step 2: Device Classes search
        if self.search_options.get("device_classes", True):
            progress += step_increment
            set_target_progress(int(progress), "Searching Device Classes...")
            device_class_paths = self.search_device_classes(vid_pid)
            registry_paths.extend(device_class_paths)
        else:
            self.result_found.emit("⏭️ Device Classes: Disabled (skipped)")

        # Step 3: USB interfaces and storage paths
        if self.search_options.get("usb_interfaces", True):
            progress += step_increment
            set_target_progress(int(progress), "Searching USB interfaces...")
            usb_paths = self.search_usb_interfaces(vid_pid)
            registry_paths.extend(usb_paths)
        else:
            self.result_found.emit("⏭️ USB Interfaces: Disabled (skipped)")

        # Step 4: System driver entries
        if self.search_options.get("system_drivers", False):
            progress += step_increment
            set_target_progress(int(progress), "Searching system drivers...")
            driver_paths = self.search_system_drivers(vid_pid)
            registry_paths.extend(driver_paths)
        else:
            self.result_found.emit("⏭️ System Drivers: Disabled (skipped)")

        # Step 5: Control panel and device manager entries
        if self.search_options.get("control_entries", False):
            progress += step_increment
            set_target_progress(int(progress), "Searching control panel entries...")
            control_paths = self.search_control_entries(vid_pid)
            registry_paths.extend(control_paths)
        else:
            self.result_found.emit("⏭️ Control Panel Entries: Disabled (skipped)")

        # Step 6: Extended PowerShell-based search
        if self.search_options.get("powershell_extended", False):
            progress += step_increment
            set_target_progress(int(progress), "Performing extended PowerShell search...")
            ps_paths = self.powershell_comprehensive_search(vid_pid)
            registry_paths.extend(ps_paths)
        else:
            self.result_found.emit("⏭️ PowerShell Extended Search: Disabled (skipped)")

        # Step 7: Friendly name search
        # if self.search_options.get("friendly_name_search", True):
        #     progress += step_increment
        #     self.progress_updated.emit(int(progress), "Searching by friendly names...")
        #     friendly_paths = self.search_by_friendly_name()
        #     registry_paths.extend(friendly_paths)
        # else:
        #     self.result_found.emit("⏭️ Friendly Name Search: Disabled (skipped)")

        # Remove duplicates and filter valid paths
        set_target_progress(99, "Filtering and validating results...")
        unique_paths = list(set(registry_paths))
        valid_paths = [path for path in unique_paths if path and len(path) > 10]

        # Show summary
        enabled_criteria = [key.replace('_', ' ').title() for key, value in self.search_options.items() if value]
        disabled_criteria = [key.replace('_', ' ').title() for key, value in self.search_options.items() if not value]

        self.result_found.emit(f"\n📊 Search Summary:")
        self.result_found.emit(f"✅ Enabled criteria: {', '.join(enabled_criteria) if enabled_criteria else 'None'}")
        self.result_found.emit(f"⏭️ Skipped criteria: {', '.join(disabled_criteria) if disabled_criteria else 'None'}")
        self.result_found.emit(f"✅ Found {len(valid_paths)} unique registry paths")

        for path in valid_paths:
            # print(path)
            self.result_found.emit(f"📝 {path}")



        self.animate_progress = False
        self.progress_updated.emit(100, "Search completed!")

        return valid_paths

    def extract_vid_pid(self) -> str:
        """Extract VID and PID from hardware ID"""
        import re
        if not self.camera.hardware_id:
            return ""

        match = re.search(r'VID_([0-9A-F]{4})&PID_([0-9A-F]{4})', self.camera.hardware_id, re.IGNORECASE)
        if match:
            return f"VID_{match.group(1)}&PID_{match.group(2)}"
        return ""

    def search_standard_device_paths(self, device_id) -> List[str]:
        """Search for registry paths containing FriendlyName entries"""
        if not device_id:
            return []

        powershell_cmd = f"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $vidPid = "{device_id}"
        $foundPaths = @()

        # Comprehensive Device Classes search
        $deviceClassesPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceClasses"

        # Extended list of relevant GUIDs for cameras and multimedia devices
        $relevantGUIDs = @(
            "{{65e8773d-8f56-11d0-a3b9-00a0c9223196}}",  # Image devices
            "{{e5323777-f976-4f5b-9b55-b94699c46e44}}",  # Camera devices
            "{{6994AD05-93EF-11D0-A3CC-00A0C9223196}}",  # Image class
            "{{4D36E96C-E325-11CE-BFC1-08002BE10318}}",  # Sound/Video devices
            "{{6bdd1fc6-810f-11d0-bec7-08002be2092f}}",  # USB devices
            "{{4d36e972-e325-11ce-bfc1-08002be10318}}",  # Multimedia devices
            "{{c06ff265-ae09-48f0-812c-16753d7cba83}}",  # WDM streaming devices
            "{{6994ad04-93ef-11d0-a3cc-00a0c9223196}}"   # Still image devices
        )

        foreach ($guid in $relevantGUIDs) {{
            $guidPath = Join-Path $deviceClassesPath $guid
            if (Test-Path $guidPath) {{
                try {{
                    $subKeys = Get-ChildItem $guidPath -ErrorAction SilentlyContinue
                    foreach ($subKey in $subKeys) {{
                        if ($subKey.Name -like "*$vidPid*") {{
                            $relativePath = $subKey.Name -replace "HKEY_LOCAL_MACHINE\\\\", ""

                            # Check specifically for #GLOBAL\Device Parameters path with FriendlyName
                            $friendlyNamePath = Join-Path $subKey.PSPath "#GLOBAL\\Device Parameters"
                            if (Test-Path $friendlyNamePath) {{
                                try {{
                                    $friendlyNameValue = Get-ItemProperty -Path $friendlyNamePath -Name "FriendlyName" -ErrorAction SilentlyContinue
                                    if ($friendlyNameValue -and $friendlyNameValue.FriendlyName) {{
                                        $foundPaths += "$relativePath\\#GLOBAL\\Device Parameters"
                                    }}
                                }} catch {{
                                    # Continue if FriendlyName property doesn't exist
                                }}
                            }}
                        }}
                    }}
                }} catch {{
                    # Continue on access errors
                }}
            }}
        }}

        $foundPaths | Where-Object {{$_ -ne ""}} | ForEach-Object {{ Write-Output $_ }}
        """

        paths = self.execute_powershell(powershell_cmd)
        self.result_found.emit(f"🎯 Found {len(paths)} registry paths with FriendlyName")
        return paths

    def search_device_manager_friendly_name(self, camera: CameraDevice) -> List[str]:
        """Search standard device enumeration paths"""
        paths = []
        # use fixed standard path
        # Standard enumeration path
        base_path = f"SYSTEM\\CurrentControlSet\\Enum\\{self.camera.device_id}"
        paths.append(base_path)

        # Add Device Parameters subkey
        # paths.append(f"{base_path}\\Device Parameters")

        # Add LogConf subkey
        # paths.append(f"{base_path}\\LogConf")

        self.result_found.emit(f"📂 Added {len(paths)} Device Manager FriendlyName entries")
        return paths



    def search_device_classes(self, vid_pid: str) -> List[str]:
        """Search in Device Classes registry"""
        if not vid_pid:
            return []

        powershell_cmd = f"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $vidPid = "{vid_pid}"
        $foundPaths = @()

        # Comprehensive Device Classes search
        $deviceClassesPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceClasses"

        # Extended list of relevant GUIDs for cameras and multimedia devices
        $relevantGUIDs = @(
            "{{65e8773d-8f56-11d0-a3b9-00a0c9223196}}",  # Image devices
            "{{e5323777-f976-4f5b-9b55-b94699c46e44}}",  # Camera devices
            "{{6994AD05-93EF-11D0-A3CC-00A0C9223196}}",  # Image class
            "{{4D36E96C-E325-11CE-BFC1-08002BE10318}}",  # Sound/Video devices
            "{{6bdd1fc6-810f-11d0-bec7-08002be2092f}}",  # USB devices
            "{{4d36e972-e325-11ce-bfc1-08002be10318}}",  # Multimedia devices
            "{{c06ff265-ae09-48f0-812c-16753d7cba83}}",  # WDM streaming devices
            "{{6994ad04-93ef-11d0-a3cc-00a0c9223196}}"   # Still image devices
        )

        foreach ($guid in $relevantGUIDs) {{
            $guidPath = Join-Path $deviceClassesPath $guid
            if (Test-Path $guidPath) {{
                try {{
                    $subKeys = Get-ChildItem $guidPath -ErrorAction SilentlyContinue
                    foreach ($subKey in $subKeys) {{
                        if ($subKey.Name -like "*$vidPid*") {{
                            $relativePath = $subKey.Name -replace "HKEY_LOCAL_MACHINE\\\\", ""
                            $foundPaths += $relativePath

                            # Check for nested paths
                            $nestedPaths = @("#GLOBAL", "Control", "Device Parameters")
                            foreach ($nested in $nestedPaths) {{
                                $nestedPath = Join-Path $subKey.PSPath $nested
                                if (Test-Path $nestedPath) {{
                                    $foundPaths += "$relativePath\\$nested"
                                }}
                            }}
                        }}
                    }}
                }} catch {{
                    # Continue on access errors
                }}
            }}
        }}

        $foundPaths | Where-Object {{$_ -ne ""}} | ForEach-Object {{ Write-Output $_ }}
        """

        paths = self.execute_powershell(powershell_cmd)
        self.result_found.emit(f"🎯 Found {len(paths)} Device Classes entries")
        return paths

    def search_usb_interfaces(self, vid_pid: str) -> List[str]:
        """Search USB interface and hub entries"""
        if not vid_pid:
            return []

        powershell_cmd = f"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $vidPid = "{vid_pid}"
        $foundPaths = @()
        
        # Search USB-specific registry locations
        $usbPaths = @(
            "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\USB",
            "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\usbhub\\Enum",
            "HKLM:\\SYSTEM\\CurrentControlSet\\Services\\usbccgp\\Enum"
        )
        
        foreach ($usbPath in $usbPaths) {{
            if (Test-Path $usbPath) {{
                try {{
                    Get-ChildItem $usbPath -Recurse -ErrorAction SilentlyContinue | 
                    Where-Object {{$_.Name -like "*$vidPid*"}} |
                    ForEach-Object {{
                        $relativePath = $_.Name -replace "HKEY_LOCAL_MACHINE\\\\", ""
                        $foundPaths += $relativePath
                    }}
                }} catch {{
                    # Continue on errors
                }}
            }}
        }}
        
        $foundPaths | Where-Object {{$_ -ne ""}} | ForEach-Object {{ Write-Output $_ }}
        """

        paths = self.execute_powershell(powershell_cmd)
        self.result_found.emit(f"🔌 Found {len(paths)} USB interface entries")
        return paths

    def search_system_drivers(self, vid_pid: str) -> List[str]:
        """Search system driver registry entries"""
        if not vid_pid:
            return []

        powershell_cmd = f"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $vidPid = "{vid_pid}"
        $foundPaths = @()
        
        # Search in Services for driver entries
        $servicesPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Services"
        
        # Common camera/video driver services
        $driverServices = @(
            "usbvideo", "ksthunk", "stream", "swenum", "usbaudio", 
            "usbccgp", "usbhub", "winusb", "wudfrd", "WUDFRd"
        )
        
        foreach ($service in $driverServices) {{
            $servicePath = Join-Path $servicesPath $service
            if (Test-Path $servicePath) {{
                try {{
                    $enumPath = Join-Path $servicePath "Enum"
                    if (Test-Path $enumPath) {{
                        $enumProps = Get-ItemProperty $enumPath -ErrorAction SilentlyContinue
                        $enumProps.PSObject.Properties | Where-Object {{
                            $_.Value -like "*$vidPid*"
                        }} | ForEach-Object {{
                            $relativePath = "SYSTEM\\CurrentControlSet\\Services\\$service\\Enum"
                            $foundPaths += $relativePath
                        }}
                    }}
                }} catch {{
                    # Continue
                }}
            }}
        }}
        
        $foundPaths | Where-Object {{$_ -ne ""}} | ForEach-Object {{ Write-Output $_ }}
        """

        paths = self.execute_powershell(powershell_cmd)
        self.result_found.emit(f"⚙️ Found {len(paths)} system driver entries")
        return paths

    def search_control_entries(self, vid_pid: str) -> List[str]:
        """Search control panel and device manager entries"""
        if not vid_pid:
            return []

        powershell_cmd = f"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $vidPid = "{vid_pid}"
        $foundPaths = @()
        
        # Search in Class registry for device classes
        $classPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class"
        
        if (Test-Path $classPath) {{
            try {{
                Get-ChildItem $classPath -ErrorAction SilentlyContinue | ForEach-Object {{
                    $classGuidPath = $_.PSPath
                    try {{
                        Get-ChildItem $classGuidPath -ErrorAction SilentlyContinue | ForEach-Object {{
                            $subKeyPath = $_.PSPath
                            try {{
                                $props = Get-ItemProperty $subKeyPath -ErrorAction SilentlyContinue
                                if ($props.MatchingDeviceId -like "*$vidPid*" -or 
                                    $props.HardwareID -like "*$vidPid*") {{
                                    $relativePath = $_.Name -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                    $foundPaths += $relativePath
                                }}
                            }} catch {{
                                # Continue
                            }}
                        }}
                    }} catch {{
                        # Continue
                    }}
                }}
            }} catch {{
                # Continue
            }}
        }}
        
        $foundPaths | Where-Object {{$_ -ne ""}} | ForEach-Object {{ Write-Output $_ }}
        """

        paths = self.execute_powershell(powershell_cmd)
        self.result_found.emit(f"🎛️ Found {len(paths)} control panel entries")
        return paths

    def powershell_comprehensive_search(self, vid_pid: str) -> List[str]:
        """Comprehensive PowerShell-based search"""
        if not vid_pid:
            return []

        powershell_cmd = f"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $vidPid = "{vid_pid}"
        $deviceId = "{self.camera.device_id}"
        $foundPaths = @()
        
        # Search for any registry keys containing our VID/PID
        $searchRoots = @(
            "HKLM:\\SYSTEM\\CurrentControlSet",
            "HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion",
            "HKLM:\\SOFTWARE\\Classes"
        )
        
        foreach ($root in $searchRoots) {{
            if (Test-Path $root) {{
                try {{
                    # Use Get-ChildItem with specific depth to avoid infinite recursion
                    Get-ChildItem $root -Recurse -Depth 3 -ErrorAction SilentlyContinue |
                    Where-Object {{$_.Name -like "*$vidPid*" -or $_.Name -like "*$deviceId*"}} |
                    ForEach-Object {{
                        $relativePath = $_.Name -replace "HKEY_LOCAL_MACHINE\\\\", ""
                        $foundPaths += $relativePath
                    }}
                }} catch {{
                    # Continue on access errors
                }}
            }}
        }}
        
        # Also search for FriendlyName entries that might reference our device
        try {{
            Get-ChildItem "HKLM:\\SYSTEM\\CurrentControlSet\\Enum" -Recurse -ErrorAction SilentlyContinue |
            Where-Object {{
                try {{
                    $props = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
                    $props.FriendlyName -like "*{self.camera.friendly_name.split()[0]}*" -or
                    $props.DeviceDesc -like "*{self.camera.friendly_name.split()[0]}*"
                }} catch {{
                    $false
                }}
            }} |
            ForEach-Object {{
                $relativePath = $_.Name -replace "HKEY_LOCAL_MACHINE\\\\", ""
                $foundPaths += $relativePath
            }}
        }} catch {{
            # Continue
        }}
        
        $foundPaths | Where-Object {{$_ -ne ""}} | Select-Object -Unique | ForEach-Object {{ Write-Output $_ }}
        """

        paths = self.execute_powershell(powershell_cmd)
        self.result_found.emit(f"🔍 Found {len(paths)} comprehensive search results")
        return paths

    def execute_powershell(self, cmd: str) -> List[str]:
        """Execute PowerShell command and return results"""
        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", cmd],
                startupinfo=startupinfo,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                timeout=30
            )

            if result.returncode == 0 and result.stdout.strip():
                paths = []
                for line in result.stdout.strip().split('\n'):
                    path = line.strip()
                    if path:
                        paths.append(path)
                return paths
        except Exception as e:
            self.result_found.emit(f"⚠️ PowerShell execution error: {str(e)}")

        return []


# Update the original RegistrySearchThread class to use the enhanced version
class RegistrySearchThread(EnhancedRegistrySearchThread):
    """Updated registry search thread - inherits from enhanced version"""
    pass


class RegistrySearchThread(QThread):
    """Thread for searching registry entries without blocking UI"""
    search_completed = Signal(list)
    progress_updated = Signal(str)

    def __init__(self, camera: CameraDevice):
        super().__init__()
        self.camera = camera

    def run(self):
        """Search for registry paths in background thread"""
        try:
            self.progress_updated.emit("Searching registry entries...")
            registry_paths = self.find_registry_paths_optimized()
            self.search_completed.emit(registry_paths)
        except Exception as e:
            print(f"Error in registry search thread: {e}")
            self.search_completed.emit([])

    def find_registry_paths_optimized(self) -> List[str]:
        """Optimized registry search with faster PowerShell queries"""
        registry_paths = []

        # Standard device path
        device_path = f"SYSTEM\\CurrentControlSet\\Enum\\{self.camera.device_id}"
        registry_paths.append(device_path)

        # Optimized PowerShell search with limited scope
        powershell_cmd = r"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $deviceID = "{}"
        $hardwareID = "{}"
        
        # Extract VID and PID for targeted search
        $vidPid = ""
        if ($hardwareID -match "VID_([0-9A-F]{{4}})&PID_([0-9A-F]{{4}})") {{
            $vidPid = "VID_$($matches[1])&PID_$($matches[2])"
        }}
        
        $foundPaths = @()
        
        # TARGETED DeviceClasses search - much faster
        if ($vidPid -ne "") {{
            $deviceClassesPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceClasses"
            
            # Only search in common camera GUIDs to speed up
            $cameraGUIDs = @(
                "{{65e8773d-8f56-11d0-a3b9-00a0c9223196}}",
                "{{e5323777-f976-4f5b-9b55-b94699c46e44}}",
                "{{6994AD05-93EF-11D0-A3CC-00A0C9223196}}",
                "{{4D36E96C-E325-11CE-BFC1-08002BE10318}}"
            )
            
            foreach ($guid in $cameraGUIDs) {{
                $guidPath = Join-Path $deviceClassesPath $guid
                if (Test-Path $guidPath) {{
                    try {{
                        # Use Get-ChildItem with Name filter for speed
                        $matchingKeys = Get-ChildItem $guidPath -Name "*$vidPid*" -ErrorAction SilentlyContinue
                        
                        foreach ($keyName in $matchingKeys) {{
                            $fullKeyPath = "SYSTEM\\CurrentControlSet\\Control\\DeviceClasses\\$guid\\$keyName"
                            $foundPaths += $fullKeyPath
                            
                            # Check for #GLOBAL\Device Parameters
                            $globalDeviceParamsPath = "$guidPath\\$keyName\\#GLOBAL\\Device Parameters"
                            if (Test-Path $globalDeviceParamsPath) {{
                                $foundPaths += "$fullKeyPath\\#GLOBAL\\Device Parameters"
                            }}
                            
                            # Check for direct Device Parameters  
                            $directDeviceParamsPath = "$guidPath\\$keyName\\Device Parameters"
                            if (Test-Path $directDeviceParamsPath) {{
                                $foundPaths += "$fullKeyPath\\Device Parameters"
                            }}
                        }}
                    }} catch {{
                        # Continue on access errors
                    }}
                }}
            }}
        }}
        
        # Quick search in Class registry for device-specific entries
        $classPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\Class"
        if (Test-Path $classPath) {{
            try {{
                # Only search in Image and Media class GUIDs
                $relevantClasses = @(
                    "{{6bdd1fc6-810f-11d0-bec7-08002be2092f}}",  # USB
                    "{{4d36e96c-e325-11ce-bfc1-08002be10318}}",  # Sound/Video
                    "{{6994ad05-93ef-11d0-a3cc-00a0c9223196}}"   # Image
                )
                
                foreach ($classGuid in $relevantClasses) {{
                    $classGuidPath = Join-Path $classPath $classGuid
                    if (Test-Path $classGuidPath) {{
                        try {{
                            $classSubKeys = Get-ChildItem $classGuidPath -ErrorAction SilentlyContinue
                            foreach ($subKey in $classSubKeys) {{
                                $subKeyPath = $subKey.PSPath
                                try {{
                                    $matchingDevicesProp = Get-ItemProperty $subKeyPath -Name "MatchingDeviceId" -ErrorAction SilentlyContinue
                                    if ($matchingDevicesProp -and $matchingDevicesProp.MatchingDeviceId -like "*$vidPid*") {{
                                        $relativePath = $subKey.Name -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                        $foundPaths += $relativePath
                                    }}
                                }} catch {{
                                    # Continue
                                }}
                            }}
                        }} catch {{
                            # Continue
                        }}
                    }}
                }}
            }} catch {{
                # Continue
            }}
        }}
        
        # Output unique paths
        $foundPaths | Select-Object -Unique | ForEach-Object {{ Write-Output $_ }}
        """.format(self.camera.device_id, self.camera.hardware_id)

        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd],
                startupinfo=startupinfo,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                timeout=15  # Reduced timeout
            )

            if result.returncode == 0 and result.stdout.strip():
                for line in result.stdout.strip().split('\n'):
                    path = line.strip()
                    if path and path not in registry_paths:
                        registry_paths.append(path)

        except subprocess.TimeoutExpired:
            print("Registry search timeout - using fallback")
        except Exception as e:
            print(f"Registry search error: {e}")

        return registry_paths


class CameraScanner(QThread):
    """Thread for scanning USB cameras"""
    cameras_found = Signal(list)
    progress_updated = Signal(int)
    status_updated = Signal(str)

    def run(self):
        """Scans for connected USB cameras"""
        try:
            self.status_updated.emit("Scanning for USB cameras...")
            cameras = []

            # PowerShell command for camera detection with UTF-8 output
            powershell_cmd = """
            [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            Get-PnpDevice | Where-Object {
                $_.Class -eq 'Camera' -or 
                $_.Class -eq 'Image' -or
                ($_.HardwareID -like '*USB*' -and $_.FriendlyName -like '*camera*') -or
                ($_.HardwareID -like '*USB*' -and $_.FriendlyName -like '*webcam*') -or
                ($_.HardwareID -like '*USB*' -and $_.FriendlyName -like '*cam*')
            } | Select-Object FriendlyName, InstanceId, HardwareID, Status | ConvertTo-Json -Depth 2
            """

            self.progress_updated.emit(25)

            # Execute PowerShell with explicit UTF-8 encoding
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd],
                startupinfo=startupinfo,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                timeout=30
            )

            self.progress_updated.emit(50)

            if result.returncode == 0 and result.stdout.strip():
                try:
                    devices_data = json.loads(result.stdout)
                    if not isinstance(devices_data, list):
                        devices_data = [devices_data]

                    self.progress_updated.emit(75)

                    for device in devices_data:
                        if device and isinstance(device, dict):
                            friendly_name = device.get('FriendlyName', 'Unknown')
                            instance_id = device.get('InstanceId', '')
                            hardware_id = device.get('HardwareID', [''])[0] if device.get('HardwareID') else ''
                            status = device.get('Status', 'Unknown')

                            # Construct registry path
                            registry_path = f"SYSTEM\\CurrentControlSet\\Enum\\{instance_id}"

                            camera = CameraDevice(
                                name=friendly_name,
                                device_id=instance_id,
                                registry_path=registry_path,
                                friendly_name=friendly_name,
                                hardware_id=hardware_id,
                                is_connected=(status == 'OK')
                            )
                            cameras.append(camera)

                    self.progress_updated.emit(100)
                    self.status_updated.emit(f"{len(cameras)} camera(s) found")

                except json.JSONDecodeError as e:
                    self.status_updated.emit(f"Error parsing camera data: {str(e)}")
            else:
                if result.stderr:
                    self.status_updated.emit(f"PowerShell error: {result.stderr}")
                else:
                    self.status_updated.emit("No cameras found")

            self.cameras_found.emit(cameras)

        except subprocess.TimeoutExpired:
            self.status_updated.emit("Timeout while scanning cameras")
            self.cameras_found.emit([])
        except Exception as e:
            self.status_updated.emit(f"Error while scanning: {str(e)}")
            self.cameras_found.emit([])


class ExitDialog(QDialog):
    """Custom exit dialog with donation links"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Thank you for using CamRenamer!")
        self.setFixedSize(450, 350)
        self.setModal(True)

        layout = QVBoxLayout()

        # Title
        title = QLabel("ℹ️ Thank you for using CamRenamer!")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        layout.addWidget(title)

        # Main message
        message = QLabel("If you found this tool helpful, consider supporting its development:")
        message.setAlignment(Qt.AlignmentFlag.AlignCenter)
        message.setFont(QFont("Arial", 11))
        message.setStyleSheet("color: #cccccc; margin: 10px;")
        message.setWordWrap(True)
        layout.addWidget(message)

        # Support options
        support_layout = QVBoxLayout()

        # Share option
        share_label = QLabel("• Share it with others")
        share_label.setFont(QFont("Arial", 10))
        share_label.setStyleSheet("color: #ffffff; margin: 5px;")
        support_layout.addWidget(share_label)

        # GitHub link
        github_label = QLabel('• Leave a star on the <a href="https://github.com/oe7set/camrenamer" style="color: #4a90e2; text-decoration: none;">🐙 GitHub Repository</a>')
        github_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        github_label.setOpenExternalLinks(True)
        github_label.setFont(QFont("Arial", 10))
        github_label.setStyleSheet("color: #ffffff; margin: 5px;")
        support_layout.addWidget(github_label)

        # Donation text
        donation_text = QLabel("• Consider a donation ☕❤️")
        donation_text.setFont(QFont("Arial", 10))
        donation_text.setStyleSheet("color: #ffffff; margin: 5px;")
        support_layout.addWidget(donation_text)

        layout.addLayout(support_layout)

        # Donation buttons
        donation_buttons_layout = QHBoxLayout()
        donation_buttons_layout.setContentsMargins(40, 10, 40, 10)

        # Buy Me A Coffee button
        coffee_button = QPushButton("☕ Buy Me A Coffee")
        coffee_button.setMinimumHeight(35)
        coffee_button.setStyleSheet("""
            QPushButton {
                background-color: #ff813f;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #ff9147;
            }
            QPushButton:pressed {
                background-color: #e6732f;
            }
        """)
        coffee_button.clicked.connect(lambda: self.open_url("https://www.buymeacoffee.com/oe7set"))
        donation_buttons_layout.addWidget(coffee_button)

        # Ko-fi button
        kofi_button = QPushButton("💙 Ko-fi")
        kofi_button.setMinimumHeight(35)
        kofi_button.setStyleSheet("""
            QPushButton {
                background-color: #29abe0;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #31b8e8;
            }
            QPushButton:pressed {
                background-color: #2399c7;
            }
        """)
        kofi_button.clicked.connect(lambda: self.open_url("https://ko-fi.com/O5O31L3XGA"))
        donation_buttons_layout.addWidget(kofi_button)

        layout.addLayout(donation_buttons_layout)

        # Exit message
        exit_message = QLabel("The application will now exit.")
        exit_message.setAlignment(Qt.AlignmentFlag.AlignCenter)
        exit_message.setFont(QFont("Arial", 10))
        exit_message.setStyleSheet("color: #aaaaaa; margin: 15px 5px 5px 5px;")
        layout.addWidget(exit_message)

        # OK Button
        ok_button = QPushButton("OK - Exit Application")
        ok_button.clicked.connect(self.accept)
        ok_button.setMinimumHeight(40)
        ok_button.setStyleSheet("""
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 12px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 12px;
                margin: 10px;
            }
            QPushButton:hover {
                background-color: #5ba0f2;
            }
            QPushButton:pressed {
                background-color: #3a80d2;
            }
        """)
        layout.addWidget(ok_button)

        self.setLayout(layout)

    def open_url(self, url):
        """Opens URL in default browser"""
        import webbrowser
        webbrowser.open(url)


class BackupThread(QThread):
    """Thread for creating registry backups without blocking UI"""
    backup_completed = Signal(str)
    backup_failed = Signal(str)
    progress_updated = Signal(str)

    def __init__(self, camera: CameraDevice, registry_paths: List[str]):
        super().__init__()
        self.camera = camera
        self.registry_paths = registry_paths

    def run(self):
        """Execute backup creation in background thread"""
        try:
            self.progress_updated.emit("Initialisiere Backup...")
            backup_path = self.create_registry_backup(self.camera, self.registry_paths)

            if backup_path:
                self.backup_completed.emit(backup_path)
            else:
                self.backup_failed.emit("Backup konnte nicht erstellt werden")

        except Exception as e:
            self.backup_failed.emit(f"Backup-Error: {str(e)}")

    def create_backup_folder(self):
        """Erstellt den Backup-Ordner im Benutzer-Documents-Verzeichnis"""
        documents_folder = os.path.join(os.path.expanduser("~"), "Documents")
        backup_folder = os.path.join(documents_folder, "CamRenamer_Backups")
        os.makedirs(backup_folder, exist_ok=True)
        return backup_folder

    def create_registry_backup(self, camera: CameraDevice, registry_paths: List[str]) -> str:
        """Creates a backup .reg file for the camera registry entries - optimized version"""
        try:
            self.progress_updated.emit("Create Backup-File...")

            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_device_id = camera.device_id.replace("\\", "_").replace("/", "_").replace(":", "_")
            safe_hardware_id = camera.hardware_id[:20].replace("\\", "_").replace("/", "_").replace(":",
                                                                                                    "_") if camera.hardware_id else "unknown"

            backup_filename = f"CamRenamer_Backup_{safe_hardware_id}_{safe_device_id}_{timestamp}.reg"
            backup_folder = self.create_backup_folder()
            backup_path = os.path.join(backup_folder, backup_filename)

            self.progress_updated.emit("Reading Registry-entries...")

            # Create single PowerShell script that processes all paths at once
            paths_array = "', '".join([path.replace('"', '').strip() for path in registry_paths])

            powershell_cmd = f"""
            $paths = @('{paths_array}')

            foreach ($regPath in $paths) {{
                $fullPath = "HKLM:\\$regPath"
                if (Test-Path $fullPath) {{
                    try {{
                        $key = Get-Item $fullPath -ErrorAction SilentlyContinue
                        if ($key) {{
                            Write-Output "[HKEY_LOCAL_MACHINE\\$regPath]"

                            $key.GetValueNames() | ForEach-Object {{
                                $valueName = $_
                                $value = $key.GetValue($valueName)
                                $valueType = $key.GetValueKind($valueName)

                                if ($valueName -eq "") {{
                                    $regValueName = "@"
                                }} else {{
                                    $regValueName = "`"$valueName`""
                                }}

                                switch ($valueType) {{
                                    "String" {{
                                        $escapedValue = $value -replace '\\\\', '\\\\\\\\' -replace '"', '\\"'
                                        Write-Output "$regValueName=`"$escapedValue`""
                                    }}
                                    "DWord" {{
                                        $hexValue = [System.Convert]::ToString([int]$value, 16).PadLeft(8, '0')
                                        Write-Output "$regValueName=dword:$hexValue"
                                    }}
                                    "QWord" {{
                                        $hexValue = [System.Convert]::ToString([long]$value, 16).PadLeft(16, '0')
                                        Write-Output "$regValueName=qword:$hexValue"
                                    }}
                                    "Binary" {{
                                        $hexString = ($value | ForEach-Object {{ [System.Convert]::ToString($_, 16).PadLeft(2, '0') }}) -join ','
                                        Write-Output "$regValueName=hex:$hexString"
                                    }}
                                    "MultiString" {{
                                        $hexBytes = [System.Text.Encoding]::Unicode.GetBytes(($value -join "`0") + "`0`0")
                                        $hexString = ($hexBytes | ForEach-Object {{ [System.Convert]::ToString($_, 16).PadLeft(2, '0') }}) -join ','
                                        Write-Output "$regValueName=hex(7):$hexString"
                                    }}
                                    "ExpandString" {{
                                        $hexBytes = [System.Text.Encoding]::Unicode.GetBytes($value + "`0")
                                        $hexString = ($hexBytes | ForEach-Object {{ [System.Convert]::ToString($_, 16).PadLeft(2, '0') }}) -join ','
                                        Write-Output "$regValueName=hex(2):$hexString"
                                    }}
                                }}
                            }}
                            Write-Output ""
                        }}
                    }} catch {{
                        Write-Output "; Error accessing: $regPath"
                        Write-Output ""
                    }}
                }} else {{
                    Write-Output "; Registry key not found: $regPath"
                    Write-Output ""
                }}
            }}
            """

            self.progress_updated.emit("Running Registry backup...")

            # Single PowerShell execution for all paths
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd],
                startupinfo=startupinfo,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                timeout=60
            )

            self.progress_updated.emit("Writing backup file...")

            # Write to file
            with open(backup_path, 'w', encoding='utf-16le') as f:
                f.write('\ufeffWindows Registry Editor Version 5.00\n\n')
                f.write(f'; CamRenamer Backup created on {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
                f.write(f'; Camera: {camera.friendly_name}\n')
                f.write(f'; Device ID: {camera.device_id}\n')
                f.write(f'; Hardware ID: {camera.hardware_id}\n\n')

                if result.returncode == 0 and result.stdout.strip():
                    f.write(result.stdout)
                else:
                    f.write(f'; Error during backup: {result.stderr}\n')

            self.progress_updated.emit("Backup created successfully!")
            return backup_path

        except Exception as e:
            print(f"Error creating backup: {e}")
            return ""


class AboutDialog(QDialog):
    """About dialog"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About CamRenamer")
        self.setFixedSize(500, 450)
        self.setModal(True)

        layout = QVBoxLayout()

        # Title
        title = QLabel("CamRenamer")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        layout.addWidget(title)

        # Version
        version = QLabel("Version 1.1")
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version.setFont(QFont("Arial", 12))
        layout.addWidget(version)

        # Author
        author = QLabel("© 2025 OE7SET - Erwin Spitaler")
        author.setAlignment(Qt.AlignmentFlag.AlignCenter)
        author.setFont(QFont("Arial", 10))
        layout.addWidget(author)

        # GitHub Link
        github_label = QLabel('<a href="https://github.com/oe7set/camrenamer" style="color: #4a90e2; text-decoration: none;">🐙 GitHub Repository</a>')
        github_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        github_label.setOpenExternalLinks(True)
        github_label.setFont(QFont("Arial", 11))
        layout.addWidget(github_label)

        # Description
        desc = QTextEdit()
        desc.setReadOnly(True)
        desc.setMaximumHeight(180)
        desc.setPlainText(
            "CamRenamer allows you to manage USB cameras and change their names "
            "in the Windows Registry. This is especially useful for "
            "streaming applications like OBS, where unique camera names are required.\n\n"
            "Features:\n"
            "• Automatic detection of USB cameras\n"
            "• Rename cameras in the Registry\n"
            "• PowerShell and Registry integration"
        )
        layout.addWidget(desc)

        # Donation section
        donation_group = QGroupBox("💖 Support Development")
        donation_layout = QVBoxLayout(donation_group)

        donation_text = QLabel("If you find this tool useful, consider supporting its development:")
        donation_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        donation_text.setFont(QFont("Arial", 10))
        donation_text.setStyleSheet("color: #cccccc; margin: 5px;")
        donation_layout.addWidget(donation_text)

        # Donation buttons layout
        buttons_layout = QHBoxLayout()

        # Buy Me A Coffee button
        coffee_button = QPushButton("☕ Buy Me A Coffee")
        coffee_button.setMinimumHeight(35)
        coffee_button.setStyleSheet("""
            QPushButton {
                background-color: #ff813f;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #ff9147;
            }
            QPushButton:pressed {
                background-color: #e6732f;
            }
        """)
        coffee_button.clicked.connect(lambda: self.open_url("https://www.buymeacoffee.com/oe7set"))
        buttons_layout.addWidget(coffee_button)

        # Ko-fi button
        kofi_button = QPushButton("💙 Ko-fi")
        kofi_button.setMinimumHeight(35)
        kofi_button.setStyleSheet("""
            QPushButton {
                background-color: #29abe0;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #31b8e8;
            }
            QPushButton:pressed {
                background-color: #2399c7;
            }
        """)
        kofi_button.clicked.connect(lambda: self.open_url("https://ko-fi.com/O5O31L3XGA"))
        buttons_layout.addWidget(kofi_button)

        donation_layout.addLayout(buttons_layout)
        layout.addWidget(donation_group)

        # OK Button
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        ok_button.setMinimumHeight(35)
        layout.addWidget(ok_button)

        self.setLayout(layout)

    def open_url(self, url):
        """Opens URL in default browser"""
        import webbrowser
        webbrowser.open(url)


class ProportionalHeaderView(QHeaderView):
    """Custom header view that maintains proportional column widths"""

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        self._proportions = []
        self._initial_widths_set = False

    def setProportionalWidths(self, widths):
        """Set the proportional widths for columns"""
        total = sum(widths)
        self._proportions = [w / total for w in widths]
        self._initial_widths_set = True
        self._updateColumnWidths()

    def _updateColumnWidths(self):
        """Update column widths based on current table width and proportions"""
        if not self._proportions or not self._initial_widths_set:
            return

        # Get available width (subtract margins and scrollbar)
        available_width = self.parent().width() - 50
        if available_width <= 0:
            return

        # Calculate new widths based on proportions
        for i, proportion in enumerate(self._proportions):
            new_width = int(available_width * proportion)
            if i < self.count():
                self.resizeSection(i, max(new_width, 50))  # Minimum width of 50px

    def resizeEvent(self, event):
        """Handle resize events to maintain proportions"""
        super().resizeEvent(event)
        if self._initial_widths_set:
            QTimer.singleShot(10, self._updateColumnWidths)  # Small delay for smooth resizing


class CamRenamerMainWindow(QMainWindow):
    """Main window of the application"""

    def __init__(self):
        super().__init__()
        self.cameras: List[CameraDevice] = []
        self.scanner_thread: Optional[CameraScanner] = None
        self.successful_rename_occurred = False  # Track if any successful rename happened

        self.setWindowTitle("CamRenamer - USB Camera Manager v1.1")
        self.setMinimumSize(640, 480)
        self.resize(1200, 750)

        # Apply modern design
        self.apply_modern_style()

        # Create UI components
        self.setup_ui()

        # Create menu and toolbar
        self.setup_menu_and_toolbar()

        # Enable Enter key for renaming
        self.new_name_edit.returnPressed.connect(self.rename_selected_camera)

        # Initial scan after short delay
        QTimer.singleShot(500, self.scan_cameras)

    def apply_modern_style(self):
        """Applies a modern dark theme"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }

            QTableWidget {
                background-color: #3c3c3c;
                border: 1px solid #555555;
                gridline-color: #555555;
                color: #ffffff;
                selection-background-color: #4a90e2;
                alternate-background-color: #404040;
            }

            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #555555;
            }

            QTableWidget::item:selected {
                background-color: #4a90e2;
                color: white;
            }

            QHeaderView::section {
                background-color: #3c3c3c;
                color: #ffffff;
                padding: 12px;
                border: none;
                border-right: 1px solid #555555;
                border-bottom: 1px solid #555555;
                font-weight: bold;
                font-size: 11px;
            }

            QHeaderView::section:hover {
                background-color: #4a90e2;
            }

            QTableWidget QHeaderView {
                background-color: #3c3c3c;
            }

            QTableWidget QHeaderView::section {
                background-color: #3c3c3c;
                color: #ffffff;
                padding: 1px 12px;
                border: none;
                border-right: 1px solid #555555;
                border-bottom: 1px solid #555555;
                font-weight: bold;
                font-size: 14px;
            }

            QTableWidget QHeaderView::section:hover {
                background-color: #4a90e2;
            }

            QTableWidget QTableCornerButton::section {
                background-color: #3c3c3c;
                border: 1px solid #555555;
            }

            QTableCornerButton::section {
                background-color: #3c3c3c;
                border: 1px solid #555555;
            }

            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 12px 20px;
                border-radius: 6px;
                font-weight: bold;
                min-width: 100px;
                font-size: 12px;
            }

            QPushButton:hover {
                background-color: #5ba0f2;
            }

            QPushButton:pressed {
                background-color: #3a80d2;
            }

            QPushButton:disabled {
                background-color: #666666;
                color: #aaaaaa;
            }

            QLineEdit {
                background-color: #3c3c3c;
                border: 2px solid #555555;
                border-radius: 6px;
                padding: 10px;
                color: #ffffff;
                font-size: 12px;
            }

            QLineEdit:focus {
                border-color: #4a90e2;
                background-color: #454545;
            }

            QLabel {
                color: #ffffff;
                font-size: 12px;
            }

            QGroupBox {
                font-weight: bold;
                border: 2px solid #555555;
                border-radius: 8px;
                margin-top: 1ex;
                padding-top: 15px;
                color: #ffffff;
                font-size: 13px;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 8px 0 8px;
            }

            QProgressBar {
                border: 2px solid #555555;
                border-radius: 6px;
                background-color: #3c3c3c;
                color: #ffffff;
                text-align: center;
                font-weight: bold;
            }

            QProgressBar::chunk {
                background-color: #4a90e2;
                border-radius: 4px;
            }

            QStatusBar {
                background-color: #404040;
                color: #ffffff;
                border-top: 1px solid #555555;
                font-size: 11px;
            }

            QToolBar {
                background-color: #404040;
                border: none;
                spacing: 6px;
                padding: 6px;
                color: #ffffff;
            }

            QToolBar QToolButton {
                color: #ffffff;
                background-color: transparent;
                border: none;
                padding: 4px;
                border-radius: 4px;
                font-size: 12px;
                font-weight: normal;
                min-width: 80px;
            }

            QToolBar QToolButton:hover {
                background-color: #4a90e2;
                color: #ffffff;
            }

            QToolBar QToolButton:pressed {
                background-color: #3a80d2;
                color: #ffffff;
            }

            QMenuBar {
                background-color: #404040;
                color: #ffffff;
                border-bottom: 1px solid #555555;
            }

            QMenuBar::item {
                padding: 8px 12px;
                background-color: transparent;
            }

            QMenuBar::item:selected {
                background-color: #4a90e2;
                border-radius: 4px;
            }

            QMenu {
                background-color: #3c3c3c;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: 4px;
            }

            QMenu::item {
                padding: 8px 20px;
            }

            QMenu::item:selected {
                background-color: #4a90e2;
            }

            QDialog {
                background-color: #2b2b2b;
                color: #ffffff;
            }

            QTextEdit {
                background-color: #3c3c3c;
                border: 1px solid #555555;
                border-radius: 4px;
                color: #ffffff;
                padding: 8px;
            }

            /* Scrollbar Styles for Dark Theme */
            QTableWidget QScrollBar:vertical {
                background-color: #2b2b2b;
                width: 16px;
                border: 1px solid #555555;
                border-radius: 8px;
                margin: 0px;
            }

            QTableWidget QScrollBar::handle:vertical {
                background-color: #4a90e2;
                min-height: 20px;
                border-radius: 6px;
                margin: 2px;
            }

            QTableWidget QScrollBar::handle:vertical:hover {
                background-color: #5ba0f2;
            }

            QTableWidget QScrollBar::handle:vertical:pressed {
                background-color: #3a80d2;
            }

            QTableWidget QScrollBar::add-line:vertical,
            QTableWidget QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 0px;
            }

            QTableWidget QScrollBar::up-arrow:vertical,
            QTableWidget QScrollBar::down-arrow:vertical {
                background: none;
                border: none;
            }

            QTableWidget QScrollBar::add-page:vertical,
            QTableWidget QScrollBar::sub-page:vertical {
                background: none;
            }

            QTableWidget QScrollBar:horizontal {
                background-color: #2b2b2b;
                height: 16px;
                border: 1px solid #555555;
                border-radius: 8px;
                margin: 0px;
            }

            QTableWidget QScrollBar::handle:horizontal {
                background-color: #4a90e2;
                min-width: 20px;
                border-radius: 6px;
                margin: 2px;
            }

            QTableWidget QScrollBar::handle:horizontal:hover {
                background-color: #5ba0f2;
            }

            QTableWidget QScrollBar::handle:horizontal:pressed {
                background-color: #3a80d2;
            }

            QTableWidget QScrollBar::add-line:horizontal,
            QTableWidget QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
                width: 0px;
            }

            QTableWidget QScrollBar::left-arrow:horizontal,
            QTableWidget QScrollBar::right-arrow:horizontal {
                background: none;
                border: none;
            }

            QTableWidget QScrollBar::add-page:horizontal,
            QTableWidget QScrollBar::sub-page:horizontal {
                background: none;
            }

            /* Corner widget between scrollbars */
            QTableWidget QScrollBar::corner {
                background-color: #2b2b2b;
                border: 1px solid #555555;
            }
        """)

    def setup_menu_and_toolbar(self):
        """Creates menu and toolbar"""
        # Menu bar
        menubar = self.menuBar()

        # File Menu
        file_menu = menubar.addMenu("&File")

        refresh_action = QAction("🔄 &Refresh", self)
        refresh_action.setShortcut("F5")
        refresh_action.setStatusTip("Reload camera list")
        refresh_action.triggered.connect(self.scan_cameras)
        file_menu.addAction(refresh_action)

        file_menu.addSeparator()

        exit_action = QAction("❌ &Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Exit application")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Tools Menu
        tools_menu = menubar.addMenu("&Tools")

        clear_action = QAction("🧹 &Clear table", self)
        clear_action.setShortcut("Ctrl+L")
        clear_action.setStatusTip("Clear camera table")
        clear_action.triggered.connect(self.clear_table)
        tools_menu.addAction(clear_action)

        # Help Menu
        help_menu = menubar.addMenu("&Help")

        about_action = QAction("ℹ️ &About", self)
        about_action.setShortcut("F1")
        about_action.setStatusTip("About this application")
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        # Toolbar
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)

        toolbar.addAction(refresh_action)
        toolbar.addSeparator()
        toolbar.addAction(clear_action)
        toolbar.addSeparator()
        toolbar.addAction(about_action)

        # Status Bar
        self.statusBar().showMessage("Ready - Press F5 to refresh")

    def setup_ui(self):
        """Creates the user interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(25, 25, 25, 25)

        # Header
        header_layout = QHBoxLayout()

        title_label = QLabel("🎥 USB Camera Manager")
        title_label.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        header_layout.addWidget(title_label)

        header_layout.addStretch()

        self.scan_button = QPushButton("🔄 Scan cameras")
        self.scan_button.clicked.connect(self.scan_cameras)
        self.scan_button.setMinimumHeight(40)
        header_layout.addWidget(self.scan_button)

        main_layout.addLayout(header_layout)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(25)
        main_layout.addWidget(self.progress_bar)

        # Camera Table
        camera_group = QGroupBox("📋 Found USB Cameras")
        camera_layout = QVBoxLayout(camera_group)

        self.camera_table = QTableWidget()
        self.camera_table.setColumnCount(5)
        self.camera_table.setHorizontalHeaderLabels([
            "🎥 Camera Name", "🔧 Device ID", "💾 Hardware ID", "🔌 Status", "⚙️ Actions"
        ])

        # Use custom proportional header view
        proportional_header = ProportionalHeaderView(Qt.Orientation.Horizontal, self.camera_table)
        self.camera_table.setHorizontalHeader(proportional_header)

        # Set proportional widths: Camera Name(20%), Device ID(30%), Hardware ID(25%), Status(12%), Actions(13%)
        proportional_header.setProportionalWidths([200, 300, 250, 120, 130])

        # Disable column moving but allow manual resizing
        proportional_header.setSectionsMovable(False)
        proportional_header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        # Table behavior settings
        self.camera_table.setAlternatingRowColors(True)
        self.camera_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.camera_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)  # Disable editing
        self.camera_table.setTextElideMode(Qt.TextElideMode.ElideMiddle)  # Enable text selection with ellipsis
        self.camera_table.setMinimumHeight(50)

        # Horizontales Scrollen aktivieren
        self.camera_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.camera_table.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)

        # Enable text selection in individual cells
        self.camera_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)

        camera_layout.addWidget(self.camera_table)
        main_layout.addWidget(camera_group)

        # Rename Section
        rename_group = QGroupBox("✏️ Rename camera")
        rename_layout = QVBoxLayout(rename_group)

        rename_form_layout = QHBoxLayout()

        rename_form_layout.addWidget(QLabel("📝 New name:"))

        self.new_name_edit = QLineEdit()
        self.new_name_edit.setPlaceholderText("Enter the new name...")
        self.new_name_edit.setMinimumHeight(35)
        rename_form_layout.addWidget(self.new_name_edit)

        self.rename_button = QPushButton("✅ Rename")
        self.rename_button.clicked.connect(self.rename_selected_camera)
        self.rename_button.setEnabled(False)
        self.rename_button.setMinimumHeight(35)
        rename_form_layout.addWidget(self.rename_button)

        rename_layout.addLayout(rename_form_layout)

        # Hint label
        hint_label = QLabel("💡 Tip: Select a camera from the table and enter a new name. Column widths adjust proportionally when resizing the window.")
        hint_label.setStyleSheet("color: #aaaaaa; font-size: 11px; font-style: italic;")
        hint_label.setWordWrap(True)
        rename_layout.addWidget(hint_label)

        main_layout.addWidget(rename_group)

        # Table selection event
        self.camera_table.itemSelectionChanged.connect(self.on_camera_selection_changed)

    def resizeEvent(self, event):
        """Handle window resize events to update table column proportions"""
        super().resizeEvent(event)
        # Trigger proportional resize with a small delay
        QTimer.singleShot(50, self._updateTableColumnWidths)

    def _updateTableColumnWidths(self):
        """Update table column widths proportionally"""
        header = self.camera_table.horizontalHeader()
        if hasattr(header, '_updateColumnWidths'):
            header._updateColumnWidths()

    def clear_table(self):
        """Clears the camera table"""
        self.cameras.clear()
        self.camera_table.setRowCount(0)
        self.new_name_edit.clear()
        self.rename_button.setEnabled(False)
        self.statusBar().showMessage("Table cleared")

    def show_about(self):
        """Shows the about dialog"""
        dialog = AboutDialog(self)
        dialog.exec()

    def on_camera_selection_changed(self):
        """Called when a camera is selected"""
        selected_rows = self.camera_table.selectionModel().selectedRows()

        if selected_rows:
            row = selected_rows[0].row()
            if row < len(self.cameras):
                camera = self.cameras[row]
                self.new_name_edit.setText(camera.friendly_name)
                self.rename_button.setEnabled(True)
                self.statusBar().showMessage(f"Camera selected: {camera.friendly_name}")
                self.new_name_edit.setFocus()
                self.new_name_edit.selectAll()
            else:
                self.rename_button.setEnabled(False)
        else:
            self.rename_button.setEnabled(False)
            self.new_name_edit.clear()
            self.statusBar().showMessage("Ready")

    def scan_cameras(self):
        """Starts the camera scan"""
        if self.scanner_thread and self.scanner_thread.isRunning():
            return

        self.scan_button.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # Visual progress timer for smooth animation
        self.visual_progress = 0
        self.scan_completed = False

        self.visual_timer = QTimer()
        self.visual_timer.timeout.connect(self.update_visual_progress)
        self.visual_timer.start(100)  # Update every 100ms

        self.scanner_thread = CameraScanner()
        self.scanner_thread.cameras_found.connect(self.on_cameras_found)
        self.scanner_thread.progress_updated.connect(self.progress_bar.setValue)
        self.scanner_thread.status_updated.connect(self.statusBar().showMessage)
        self.scanner_thread.finished.connect(self.on_scan_finished)
        self.scanner_thread.start()

    def update_visual_progress(self):
        """Updates visual progress with smooth animation"""
        if self.scan_completed:
            # Fast completion when scan is done
            if self.visual_progress < 100:
                self.visual_progress = min(100, self.visual_progress + 10)
                self.progress_bar.setValue(self.visual_progress)
            else:
                self.visual_timer.stop()
            return

        # Normal progress animation
        if self.visual_progress < 90:
            # Fast progress until 90%
            self.visual_progress = min(90, self.visual_progress + 2.2)
        else:
            # Slower progress after 90%
            if self.visual_progress < 98:
                self.visual_progress = min(98, self.visual_progress + 0.2)

        self.progress_bar.setValue(int(self.visual_progress))

    def on_real_progress_updated(self, value):
        """Handle real progress updates from scanner thread"""
        # Optional: sync visual progress with real progress if needed
        if value > self.visual_progress:
            self.visual_progress = min(value, 98)

    def on_scan_finished(self):
        """Called when the scan is finished"""
        self.scan_completed = True
        self.scan_button.setEnabled(True)

        # Hide progress bar after completion animation
        QTimer.singleShot(500, lambda: self.progress_bar.setVisible(False))

    def on_scan_finished(self):
        """Called when the scan is finished"""
        self.scan_button.setEnabled(True)
        self.progress_bar.setVisible(False)

    def on_cameras_found(self, cameras: List[CameraDevice]):
        """Called when cameras are found"""
        self.cameras = cameras
        self.update_camera_table()

    def update_camera_table(self):
        """Updates the camera table"""
        self.camera_table.setRowCount(len(self.cameras))

        for row, camera in enumerate(self.cameras):
            # Name
            name_item = QTableWidgetItem(camera.friendly_name)
            name_item.setToolTip(f"Full name: {camera.friendly_name}")
            name_item.setFlags(name_item.flags() | Qt.ItemFlag.ItemIsSelectable)
            self.camera_table.setItem(row, 0, name_item)

            # Device ID (shortened)
            device_id = camera.device_id[:50] + "..." if len(camera.device_id) > 50 else camera.device_id
            device_item = QTableWidgetItem(device_id)
            device_item.setToolTip(f"Full Device ID: {camera.device_id}")
            device_item.setFlags(device_item.flags() | Qt.ItemFlag.ItemIsSelectable)
            self.camera_table.setItem(row, 1, device_item)

            # Hardware ID (shortened)
            hardware_id = camera.hardware_id[:30] + "..." if len(camera.hardware_id) > 30 else camera.hardware_id
            hardware_item = QTableWidgetItem(hardware_id)
            hardware_item.setToolTip(f"Full Hardware ID: {camera.hardware_id}")
            hardware_item.setFlags(hardware_item.flags() | Qt.ItemFlag.ItemIsSelectable)
            self.camera_table.setItem(row, 2, hardware_item)

            # Status
            status_text = "🟢 Connected" if camera.is_connected else "🔴 Disconnected"
            status_item = QTableWidgetItem(status_text)
            status_item.setToolTip(f"Status: {'Active and ready' if camera.is_connected else 'Not available'}")
            status_item.setFlags(status_item.flags() | Qt.ItemFlag.ItemIsSelectable)
            self.camera_table.setItem(row, 3, status_item)

            # Actions
            action_item = QTableWidgetItem("👆 Select to rename")
            action_item.setToolTip("Click this row to select the camera")
            action_item.setFlags(action_item.flags() | Qt.ItemFlag.ItemIsSelectable)
            self.camera_table.setItem(row, 4, action_item)

    def rename_selected_camera(self):
        """Renames the selected camera with enhanced registry search dialog"""
        selected_rows = self.camera_table.selectionModel().selectedRows()

        if not selected_rows:
            QMessageBox.warning(self, "⚠️ Warning", "Please select a camera from the table.")
            return

        new_name = self.new_name_edit.text().strip()
        if not new_name:
            QMessageBox.warning(self, "⚠️ Warning", "Please enter a new name.")
            self.new_name_edit.setFocus()
            return

        if len(new_name) > 255:
            QMessageBox.warning(self, "⚠️ Warning", "The name is too long. Maximum 255 characters allowed.")
            return

        row = selected_rows[0].row()
        camera = self.cameras[row]

        if new_name == camera.friendly_name:
            QMessageBox.information(self, "ℹ️ Information", "The new name is identical to the current name.")
            return

        # Confirmation
        reply = QMessageBox.question(
            self, "❓ Confirmation",
            f"Do you really want to rename '{camera.friendly_name}' to '{new_name}'?\n\n"
            "⚠️ Note: This requires administrator rights and may require a restart.\n"
            "The change will be stored in the Windows Registry.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Show registry search dialog
            # search_dialog = RegistrySearchDialog(self)

            # Create and setup the enhanced registry search thread
            registry_search = EnhancedRegistrySearchThread(camera)
            search_dialog = RegistrySearchDialog(self, registry_search.search_options)
            registry_paths = []

            # Connect signals to the dialog
            def on_progress_updated(value, status):
                search_dialog.update_progress(value, status)

            def on_result_found(text):
                search_dialog.add_result(text)

            def on_search_completed(paths):
                nonlocal registry_paths
                registry_paths = paths
                search_dialog.search_completed()
                if paths:
                    search_dialog.add_result(f"\n🎉 Registry search completed! Found {len(paths)} paths total.")
                else:
                    search_dialog.add_result("\n❌ No registry paths found. The camera might not be properly detected.")

            registry_search.progress_updated.connect(on_progress_updated)
            registry_search.result_found.connect(on_result_found)
            registry_search.search_completed.connect(on_search_completed)

            # Start the search
            registry_search.start()

            # Show the dialog (it will be modal)
            search_dialog.exec()

            # Wait for search to complete if still running
            if registry_search.isRunning():
                registry_search.wait()

            # Now proceed with the actual renaming if we found paths
            if registry_paths:
                success = self.update_camera_name_in_registry_with_paths(camera, new_name, registry_paths)

                if success:
                    # Update table
                    camera.friendly_name = new_name
                    camera.name = new_name
                    self.update_camera_table()
                    self.statusBar().showMessage(f"Camera successfully renamed to: {new_name}")
                    self.successful_rename_occurred = True
                else:
                    QMessageBox.critical(
                        self, "❌ Error",
                        "Error renaming the camera!\n\n"
                        "Possible causes:\n"
                        "• No administrator rights\n"
                        "• Camera is currently in use\n"
                        "• Registry access denied\n\n"
                        "Try running the application as administrator."
                    )
                    self.statusBar().showMessage("Error renaming")
            else:
                QMessageBox.warning(
                    self, "⚠️ No Registry Paths Found",
                    f"Could not find any registry entries for camera '{camera.friendly_name}'.\n\n"
                    "This might happen if:\n"
                    "• The camera is not properly installed\n"
                    "• Access to registry is restricted\n"
                    "• The camera uses a different naming structure\n\n"
                    "Try running as administrator or check if the camera is properly connected."
                )

    def update_camera_name_in_registry_with_paths(self, camera: CameraDevice, new_name: str, registry_paths: List[str]) -> bool:
        """Updates the camera name in the registry using provided paths with threaded backup"""
        try:
            # Backup-Progress-Dialog erstellen (non-modal)
            backup_dialog = QProgressDialog("Erstelle Registry-Backup...", "", 0, 0, self)
            backup_dialog.setWindowTitle("Backup wird erstellt")
            backup_dialog.setCancelButton(None)
            backup_dialog.setMinimumDuration(0)
            backup_dialog.show()
            QApplication.processEvents()
            # Backup-Thread erstellen und starten
            self.backup_thread = BackupThread(camera, registry_paths)
            backup_path = ""
            backup_success = False

            # Signals verbinden
            def on_backup_completed(path):
                nonlocal backup_path, backup_success
                backup_path = path
                backup_success = True
                backup_dialog.setLabelText("Backup erfolgreich erstellt!")
                QTimer.singleShot(500, backup_dialog.close)

            def on_backup_failed(error_msg):
                nonlocal backup_success
                backup_success = False
                backup_dialog.setLabelText(f"Backup-Fehler: {error_msg}")
                QTimer.singleShot(1500, backup_dialog.close)

            def on_backup_progress(status):
                backup_dialog.setLabelText(status)
                self.statusBar().showMessage(status)


            self.backup_thread.backup_completed.connect(on_backup_completed)
            self.backup_thread.backup_failed.connect(on_backup_failed)
            self.backup_thread.progress_updated.connect(on_backup_progress)

            # Thread starten und auf Completion warten
            self.backup_thread.start()

            # Event-Loop für Dialog während Backup läuft
            while self.backup_thread.isRunning():
                QApplication.processEvents()
                self.backup_thread.msleep(20)

            # Warten bis Thread beendet ist
            self.backup_thread.wait()
            # Final event processing to ensure signal is handled
            QApplication.processEvents()
            print(backup_path)
            success_count = 0
            total_paths = len(registry_paths)

            # Update each registry path
            for i, registry_path in enumerate(registry_paths):
                self.statusBar().showMessage(f"Updating registry {i+1}/{total_paths}...")

                try:
                    # Try direct registry access first
                    with winreg.OpenKey(
                            winreg.HKEY_LOCAL_MACHINE,
                            registry_path,
                            0,
                            winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY
                    ) as key:
                        # Set FriendlyName
                        winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, new_name)
                        success_count += 1

                except (FileNotFoundError, PermissionError, OSError):
                    # Try with PowerShell as fallback
                    try:
                        powershell_cmd = f"""
                        $regPath = "HKLM:\\{registry_path}"
                        if (Test-Path $regPath) {{
                            try {{
                                Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "{new_name}" -Force
                                Write-Output "SUCCESS"
                            }} catch {{
                                Write-Output "ERROR: $($_.Exception.Message)"
                            }}
                        }} else {{
                            Write-Output "PATH_NOT_FOUND"
                        }}
                        """

                        result = subprocess.run(
                            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd],
                            startupinfo=startupinfo,
                            capture_output=True,
                            text=True,
                            encoding='utf-8',
                            errors='replace',
                            timeout=5
                        )

                        if result.returncode == 0 and "SUCCESS" in result.stdout:
                            success_count += 1

                    except Exception:
                        continue


            # Show backup information if successful
            if success_count > 0 and backup_path:
                QMessageBox.information(
                    self, "✅ Success with Backup",
                    f"Camera was successfully renamed to '{new_name}'!\n\n"
                    f"Updated {success_count} of {total_paths} registry locations.\n\n"
                    f"💾 Backup created: {os.path.basename(backup_path)}\n"
                    f"📁 Backup folder: {os.path.dirname(backup_path)}\n\n"
                    "💡 A system restart may be required for the changes "
                    "to take effect in all applications."
                )
            elif success_count > 0:
                QMessageBox.information(
                    self, "✅ Partial Success",
                    f"Camera was partially renamed to '{new_name}'!\n\n"
                    f"Updated {success_count} of {total_paths} registry locations.\n\n"
                    "💡 A system restart may be required for the changes "
                    "to take effect in all applications."
                )

            return success_count > 0

        except Exception as e:
            print(f"Error during registry update: {e}")
            return False

    def create_backup_folder(self):
        """Creates the backup folder if it doesn't exist"""
        backup_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "CamRenamer_Backups")
        os.makedirs(backup_folder, exist_ok=True)
        return backup_folder

    def create_registry_backup(self, camera: CameraDevice, registry_paths: List[str]) -> str:
        """Creates a backup .reg file for the camera registry entries - optimized version"""
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_device_id = camera.device_id.replace("\\", "_").replace("/", "_").replace(":", "_")
            safe_hardware_id = camera.hardware_id[:20].replace("\\", "_").replace("/", "_").replace(":",
                                                                                                    "_") if camera.hardware_id else "unknown"

            backup_filename = f"CamRenamer_Backup_{safe_hardware_id}_{safe_device_id}_{timestamp}.reg"
            backup_folder = self.create_backup_folder()
            backup_path = os.path.join(backup_folder, backup_filename)

            # Create single PowerShell script that processes all paths at once
            paths_array = "', '".join([path.replace('"', '').strip() for path in registry_paths])

            powershell_cmd = f"""
            $paths = @('{paths_array}')

            foreach ($regPath in $paths) {{
                $fullPath = "HKLM:\\$regPath"
                if (Test-Path $fullPath) {{
                    try {{
                        $key = Get-Item $fullPath -ErrorAction SilentlyContinue
                        if ($key) {{
                            Write-Output "[HKEY_LOCAL_MACHINE\\$regPath]"

                            $key.GetValueNames() | ForEach-Object {{
                                $valueName = $_
                                $value = $key.GetValue($valueName)
                                $valueType = $key.GetValueKind($valueName)

                                if ($valueName -eq "") {{
                                    $regValueName = "@"
                                }} else {{
                                    $regValueName = "`"$valueName`""
                                }}

                                switch ($valueType) {{
                                    "String" {{
                                        $escapedValue = $value -replace '\\\\', '\\\\\\\\' -replace '"', '\\"'
                                        Write-Output "$regValueName=`"$escapedValue`""
                                    }}
                                    "DWord" {{
                                        $hexValue = [System.Convert]::ToString([int]$value, 16).PadLeft(8, '0')
                                        Write-Output "$regValueName=dword:$hexValue"
                                    }}
                                    "QWord" {{
                                        $hexValue = [System.Convert]::ToString([long]$value, 16).PadLeft(16, '0')
                                        Write-Output "$regValueName=qword:$hexValue"
                                    }}
                                    "Binary" {{
                                        $hexString = ($value | ForEach-Object {{ [System.Convert]::ToString($_, 16).PadLeft(2, '0') }}) -join ','
                                        Write-Output "$regValueName=hex:$hexString"
                                    }}
                                    "MultiString" {{
                                        $hexBytes = [System.Text.Encoding]::Unicode.GetBytes(($value -join "`0") + "`0`0")
                                        $hexString = ($hexBytes | ForEach-Object {{ [System.Convert]::ToString($_, 16).PadLeft(2, '0') }}) -join ','
                                        Write-Output "$regValueName=hex(7):$hexString"
                                    }}
                                    "ExpandString" {{
                                        $hexBytes = [System.Text.Encoding]::Unicode.GetBytes($value + "`0")
                                        $hexString = ($hexBytes | ForEach-Object {{ [System.Convert]::ToString($_, 16).PadLeft(2, '0') }}) -join ','
                                        Write-Output "$regValueName=hex(2):$hexString"
                                    }}
                                }}
                            }}
                            Write-Output ""
                        }}
                    }} catch {{
                        Write-Output "; Error accessing: $regPath"
                        Write-Output ""
                    }}
                }} else {{
                    Write-Output "; Registry key not found: $regPath"
                    Write-Output ""
                }}
            }}
            """

            # Single PowerShell execution for all paths
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd],
                startupinfo=startupinfo,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                timeout=60
            )

            # Write to file
            with open(backup_path, 'w', encoding='utf-16le') as f:
                f.write('\ufeffWindows Registry Editor Version 5.00\n\n')
                f.write(f'; CamRenamer Backup created on {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
                f.write(f'; Camera: {camera.friendly_name}\n')
                f.write(f'; Device ID: {camera.device_id}\n')
                f.write(f'; Hardware ID: {camera.hardware_id}\n\n')

                if result.returncode == 0 and result.stdout.strip():
                    f.write(result.stdout)
                else:
                    f.write(f'; Error during backup: {result.stderr}\n')

            return backup_path
        except Exception as e:
            print(f"Error creating backup: {e}")
            return ""

    def update_camera_name_in_registry(self, camera: CameraDevice, new_name: str) -> bool:
        """Updates the camera name in the registry with optimized threaded search"""
        try:
            # Use the optimized registry search thread for backwards compatibility
            registry_search = RegistrySearchThread(camera)
            registry_paths = []

            # Connect to get results
            def on_search_completed(paths):
                nonlocal registry_paths
                registry_paths = paths

            registry_search.search_completed.connect(on_search_completed)
            registry_search.progress_updated.connect(self.statusBar().showMessage)

            # Run synchronously and wait for completion
            registry_search.start()
            registry_search.wait()  # Wait for thread to complete

            if not registry_paths:
                return False

            # Use the new function with found paths
            return self.update_camera_name_in_registry_with_paths(camera, new_name, registry_paths)

        except Exception as e:
            print(f"Error during registry update: {e}")
            return False

    def closeEvent(self, event):
        """Called when the application is closed"""
        if self.scanner_thread and self.scanner_thread.isRunning():
            self.scanner_thread.terminate()
            self.scanner_thread.wait(3000)

        # If any successful rename occurred, show the custom exit dialog
        if self.successful_rename_occurred:
            dialog = ExitDialog(self)
            dialog.exec()

        event.accept()


def main():
    """Main function"""
    app = QApplication(sys.argv)

    # Set app properties
    app.setApplicationName("CamRenamer")
    app.setApplicationVersion("1.1")
    app.setOrganizationName("Retroverse")
    app.setOrganizationDomain("retroverse.de")

    icon = QIcon(":/img/icon.png")
    app.setWindowIcon(icon)

    # Enable high-DPI support and improve text rendering
    print("CamRenamer is starting...")
    #app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    #app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    app.setAttribute(Qt.ApplicationAttribute.AA_SynthesizeMouseForUnhandledTabletEvents, False)

    # Set default font with antialiasing
    font = QFont("Segoe UI", 10)
    font.setHintingPreference(QFont.HintingPreference.PreferFullHinting)
    font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
    app.setFont(font)

    # Splash Screen
    splash_pixmap = QPixmap(450, 300)
    splash_pixmap.fill(QColor(43, 43, 43))

    painter = QPainter(splash_pixmap)
    painter.setPen(QColor(255, 255, 255))
    painter.setFont(QFont("Arial", 26, QFont.Weight.Bold))

    # Draw title
    title_rect = splash_pixmap.rect()
    title_rect.setTop(80)
    title_rect.setHeight(50)
    painter.drawText(title_rect, Qt.AlignmentFlag.AlignCenter, "🎥 CamRenamer")

    # Draw subtitle
    painter.setFont(QFont("Arial", 14))
    subtitle_rect = splash_pixmap.rect()
    subtitle_rect.setTop(140)
    subtitle_rect.setHeight(30)
    painter.drawText(subtitle_rect, Qt.AlignmentFlag.AlignCenter, "USB Camera Manager v1.1")

    # Draw status
    painter.setFont(QFont("Arial", 12))
    status_rect = splash_pixmap.rect()
    status_rect.setTop(200)
    status_rect.setHeight(30)
    painter.drawText(status_rect, Qt.AlignmentFlag.AlignCenter, "Loading...")

    painter.end()

    splash = QSplashScreen(splash_pixmap)
    splash.show()

    # Short wait for splash
    for i in range(10):
        time.sleep(0.1)
        app.processEvents()

    # Create main window
    window = CamRenamerMainWindow()

    # Close splash and show main window
    splash.finish(window)
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())










