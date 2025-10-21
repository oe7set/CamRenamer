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
    QSplashScreen, QToolBar, QDialog, QTextEdit
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QFont, QColor, QPixmap, QPainter, QAction


@dataclass
class CameraDevice:
    """Data class for USB camera information"""
    name: str
    device_id: str
    registry_path: str
    friendly_name: str
    hardware_id: str
    is_connected: bool = True


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

        self.scanner_thread = CameraScanner()
        self.scanner_thread.cameras_found.connect(self.on_cameras_found)
        self.scanner_thread.progress_updated.connect(self.progress_bar.setValue)
        self.scanner_thread.status_updated.connect(self.statusBar().showMessage)
        self.scanner_thread.finished.connect(self.on_scan_finished)
        self.scanner_thread.start()

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
        """Renames the selected camera"""
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
            # Block UI during renaming
            self.setEnabled(False)
            self.statusBar().showMessage("Renaming camera...")

            success = self.update_camera_name_in_registry(camera, new_name)

            self.setEnabled(True)

            if success:
                # Success message is now handled in update_camera_name_in_registry
                # Update table
                camera.friendly_name = new_name
                camera.name = new_name
                self.update_camera_table()
                self.statusBar().showMessage(f"Camera successfully renamed to: {new_name}")
                self.successful_rename_occurred = True  # Mark that a successful rename happened
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

    def create_backup_folder(self):
        """Creates the backup folder if it doesn't exist"""
        backup_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "CamRenamer_Backups")
        os.makedirs(backup_folder, exist_ok=True)
        return backup_folder

    def create_registry_backup(self, camera: CameraDevice, registry_paths: List[str]) -> str:
        """Creates a backup .reg file for the camera registry entries"""
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_device_id = camera.device_id.replace("\\", "_").replace("/", "_").replace(":", "_")
            safe_hardware_id = camera.hardware_id[:20].replace("\\", "_").replace("/", "_").replace(":", "_") if camera.hardware_id else "unknown"

            backup_filename = f"CamRenamer_Backup_{safe_hardware_id}_{safe_device_id}_{timestamp}.reg"
            backup_folder = self.create_backup_folder()
            backup_path = os.path.join(backup_folder, backup_filename)

            with open(backup_path, 'w', encoding='utf-16le') as f:
                # Write REG file header
                f.write('\ufeffWindows Registry Editor Version 5.00\n\n')
                f.write(f'; CamRenamer Backup created on {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
                f.write(f'; Camera: {camera.friendly_name}\n')
                f.write(f'; Device ID: {camera.device_id}\n')
                f.write(f'; Hardware ID: {camera.hardware_id}\n\n')

                # Export each registry path
                for reg_path in registry_paths:
                    try:
                        # Use reg export command to get the current values
                        export_cmd = f'reg export "HKLM\\{reg_path}" - /y'
                        result = subprocess.run(
                            export_cmd,
                            shell=True,
                            capture_output=True,
                            text=True,
                            encoding='utf-16le',
                            errors='replace'
                        )

                        if result.returncode == 0 and result.stdout:
                            # Skip the header from reg export and add our own comment
                            lines = result.stdout.split('\n')
                            f.write(f'; Registry path: {reg_path}\n')
                            for line in lines[1:]:  # Skip first line (header)
                                if line.strip():
                                    f.write(line + '\n')
                            f.write('\n')
                    except Exception as e:
                        f.write(f'; Error exporting {reg_path}: {str(e)}\n\n')

            return backup_path
        except Exception as e:
            print(f"Error creating backup: {e}")
            return ""

    def find_all_camera_registry_paths(self, camera: CameraDevice) -> List[str]:
        """Finds all registry paths for the camera including all Control entries"""
        registry_paths = []

        # Standard device path
        device_path = f"SYSTEM\\CurrentControlSet\\Enum\\{camera.device_id}"
        registry_paths.append(device_path)

        # Search for ALL entries in Control using PowerShell
        powershell_cmd = f"""
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $deviceID = "{camera.device_id}"
        $hardwareID = "{camera.hardware_id}"
        
        # Extract main parts from hardware ID for broader matching
        $hwIDParts = @()
        if ($hardwareID -ne "") {{
            $hwIDParts += $hardwareID -split '&' | Where-Object {{ $_ -ne "" }}
            $hwIDParts += ($hardwareID -split '\\\\')[0]  # Get main part before backslash
        }}
        
        # Extract device ID parts for matching
        $deviceIDParts = @()
        if ($deviceID -ne "") {{
            $deviceIDParts += $deviceID -split '\\\\' | Where-Object {{ $_ -ne "" }}
        }}
        
        # Extract VID and PID from hardware ID for specific DeviceClasses matching
        $vidPid = ""
        if ($hardwareID -match "VID_([0-9A-F]{{4}})&PID_([0-9A-F]{{4}})") {{
            $vidPid = "VID_$($matches[1])&PID_$($matches[2])"
        }}
        
        $foundPaths = @()
        
        # Search in entire Control directory
        $controlPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Control"
        if (Test-Path $controlPath) {{
            try {{
                # Get all subdirectories in Control
                $controlSubDirs = Get-ChildItem $controlPath -Recurse -ErrorAction SilentlyContinue | Where-Object {{ $_.PSIsContainer }}
                
                foreach ($subDir in $controlSubDirs) {{
                    $keyPath = $subDir.PSPath
                    $keyName = $subDir.Name
                    
                    # Check if this key contains our device ID or hardware ID
                    $matchFound = $false
                    
                    # Direct device ID match
                    if ($keyName -like "*$deviceID*") {{
                        $matchFound = $true
                    }}
                    
                    # Hardware ID parts match
                    foreach ($hwPart in $hwIDParts) {{
                        if ($hwPart.Length -gt 3 -and $keyName -like "*$hwPart*") {{
                            $matchFound = $true
                            break
                        }}
                    }}
                    
                    # Device ID parts match  
                    foreach ($devPart in $deviceIDParts) {{
                        if ($devPart.Length -gt 3 -and $keyName -like "*$devPart*") {{
                            $matchFound = $true
                            break
                        }}
                    }}
                    
                    # VID/PID match for DeviceClasses
                    if ($vidPid -ne "" -and $keyName -like "*$vidPid*") {{
                        $matchFound = $true
                    }}
                    
                    if ($matchFound) {{
                        $relativePath = $keyName -replace "HKEY_LOCAL_MACHINE\\\\", ""
                        $foundPaths += $relativePath
                        
                        # Also check for common subkeys that might contain FriendlyName
                        $commonSubKeys = @("Device Parameters", "#GLOBAL", "Properties", "Control")
                        foreach ($subKeyName in $commonSubKeys) {{
                            $subKeyPath = Join-Path $keyPath $subKeyName
                            if (Test-Path $subKeyPath) {{
                                $relativeSubPath = ($subKeyPath -replace "Microsoft.PowerShell.Core\\\\Registry::", "") -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                $foundPaths += $relativeSubPath
                                
                                # Check Device Parameters under these subkeys too
                                $deviceParamsPath = Join-Path $subKeyPath "Device Parameters"
                                if (Test-Path $deviceParamsPath) {{
                                    $relativeDeviceParams = ($deviceParamsPath -replace "Microsoft.PowerShell.Core\\\\Registry::", "") -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                    $foundPaths += $relativeDeviceParams
                                }}
                            }}
                        }}
                    }}
                }}
                
                # SPECIFIC SEARCH for DeviceClasses with ##?# structure
                Write-Host "Searching specifically for DeviceClasses with ##?# structure..."
                $deviceClassesPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Control\\DeviceClasses"
                if (Test-Path $deviceClassesPath) {{
                    try {{
                        # Get all GUID folders in DeviceClasses
                        $guidFolders = Get-ChildItem $deviceClassesPath -ErrorAction SilentlyContinue | Where-Object {{ $_.PSIsContainer }}
                        
                        foreach ($guidFolder in $guidFolders) {{
                            try {{
                                # Look for subkeys that match our device pattern
                                $deviceKeys = Get-ChildItem $guidFolder.PSPath -ErrorAction SilentlyContinue | Where-Object {{ $_.PSIsContainer }}
                                
                                foreach ($deviceKey in $deviceKeys) {{
                                    $deviceKeyName = $deviceKey.Name
                                    $deviceKeyBaseName = Split-Path $deviceKeyName -Leaf
                                    
                                    # Check if this device key matches our criteria
                                    $deviceMatchFound = $false
                                    
                                    # Check for VID/PID match in ##?# structure
                                    if ($vidPid -ne "" -and $deviceKeyBaseName -like "*$vidPid*") {{
                                        $deviceMatchFound = $true
                                    }}
                                    
                                    # Check for device ID parts in the key name
                                    foreach ($devPart in $deviceIDParts) {{
                                        if ($devPart.Length -gt 3 -and $deviceKeyBaseName -like "*$devPart*") {{
                                            $deviceMatchFound = $true
                                            break
                                        }}
                                    }}
                                    
                                    # Check for hardware ID parts
                                    foreach ($hwPart in $hwIDParts) {{
                                        if ($hwPart.Length -gt 3 -and $deviceKeyBaseName -like "*$hwPart*") {{
                                            $deviceMatchFound = $true
                                            break
                                        }}
                                    }}
                                    
                                    if ($deviceMatchFound) {{
                                        Write-Host "Found matching DeviceClass: $deviceKeyName"
                                        $relativeDeviceKeyPath = $deviceKeyName -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                        $foundPaths += $relativeDeviceKeyPath
                                        
                                        # Check for #GLOBAL subkey
                                        $globalPath = Join-Path $deviceKey.PSPath "#GLOBAL"
                                        if (Test-Path $globalPath) {{
                                            $relativeGlobalPath = ($globalPath -replace "Microsoft.PowerShell.Core\\\\Registry::", "") -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                            $foundPaths += $relativeGlobalPath
                                            Write-Host "Found #GLOBAL: $relativeGlobalPath"
                                            
                                            # Check Device Parameters under #GLOBAL
                                            $globalDeviceParams = Join-Path $globalPath "Device Parameters"
                                            if (Test-Path $globalDeviceParams) {{
                                                $relativeGlobalDeviceParams = ($globalDeviceParams -replace "Microsoft.PowerShell.Core\\\\Registry::", "") -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                                $foundPaths += $relativeGlobalDeviceParams
                                                Write-Host "Found Device Parameters: $relativeGlobalDeviceParams"
                                            }}
                                        }}
                                        
                                        # Also check for Device Parameters directly under the device key
                                        $directDeviceParams = Join-Path $deviceKey.PSPath "Device Parameters"
                                        if (Test-Path $directDeviceParams) {{
                                            $relativeDirectDeviceParams = ($directDeviceParams -replace "Microsoft.PowerShell.Core\\\\Registry::", "") -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                            $foundPaths += $relativeDirectDeviceParams
                                            Write-Host "Found direct Device Parameters: $relativeDirectDeviceParams"
                                        }}
                                    }}
                                }}
                            }} catch {{
                                # Continue on access errors for individual GUID folders
                                Write-Host "Access denied to GUID folder: $($guidFolder.Name)"
                            }}
                        }}
                    }} catch {{
                        Write-Host "Access denied to DeviceClasses"
                    }}
                }}
                
                # Also search specifically in known camera-related paths
                $specificPaths = @(
                    "Class",
                    "CoDeviceInstallers", 
                    "MediaCategories",
                    "MediaInterfaces",
                    "MediaResources",
                    "MediaSets"
                )
                
                foreach ($specificPath in $specificPaths) {{
                    $fullSpecificPath = Join-Path $controlPath $specificPath
                    if (Test-Path $fullSpecificPath) {{
                        try {{
                            $specificSubDirs = Get-ChildItem $fullSpecificPath -Recurse -ErrorAction SilentlyContinue | Where-Object {{ $_.PSIsContainer }}
                            
                            foreach ($specificSubDir in $specificSubDirs) {{
                                $specificKeyName = $specificSubDir.Name
                                $specificMatchFound = $false
                                
                                # Check device ID match
                                if ($specificKeyName -like "*$deviceID*") {{
                                    $specificMatchFound = $true
                                }}
                                
                                # Check VID/PID match
                                if ($vidPid -ne "" -and $specificKeyName -like "*$vidPid*") {{
                                    $specificMatchFound = $true
                                }}
                                
                                # Check hardware ID parts match
                                foreach ($hwPart in $hwIDParts) {{
                                    if ($hwPart.Length -gt 3 -and $specificKeyName -like "*$hwPart*") {{
                                        $specificMatchFound = $true
                                        break
                                    }}
                                }}
                                
                                if ($specificMatchFound) {{
                                    $specificRelativePath = $specificKeyName -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                    $foundPaths += $specificRelativePath
                                    
                                    # Check for subkeys
                                    $specificCommonSubKeys = @("Device Parameters", "#GLOBAL", "Properties")
                                    foreach ($specificSubKeyName in $specificCommonSubKeys) {{
                                        $specificSubKeyPath = Join-Path $specificSubDir.PSPath $specificSubKeyName
                                        if (Test-Path $specificSubKeyPath) {{
                                            $specificRelativeSubPath = ($specificSubKeyPath -replace "Microsoft.PowerShell.Core\\\\Registry::", "") -replace "HKEY_LOCAL_MACHINE\\\\", ""
                                            $foundPaths += $specificRelativeSubPath
                                        }}
                                    }}
                                }}
                            }}
                        }} catch {{
                            # Continue on access errors
                        }}
                    }}
                }}
            }} catch {{
                # Continue on access errors
            }}
        }}
        
        # Remove duplicates and output found paths
        $foundPaths | Select-Object -Unique | ForEach-Object {{ Write-Output $_ }}
        """

        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                timeout=60  # Increased timeout for more extensive search
            )

            if result.returncode == 0 and result.stdout.strip():
                for line in result.stdout.strip().split('\n'):
                    path = line.strip()
                    if path and path not in registry_paths and not path.startswith('Searching') and not path.startswith('Found') and not path.startswith('Access denied'):
                        registry_paths.append(path)

        except Exception as e:
            print(f"Error searching Control registry: {e}")

        return registry_paths

    def update_camera_name_in_registry(self, camera: CameraDevice, new_name: str) -> bool:
        """Updates the camera name in the registry with optimized threaded search"""
        try:
            # Use the optimized registry search thread
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

            # Create backup before making changes
            self.statusBar().showMessage("Creating backup...")
            backup_path = self.create_registry_backup(camera, registry_paths)

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
                            capture_output=True,
                            text=True,
                            encoding='utf-8',
                            errors='replace',
                            timeout=5  # Reduced timeout
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
