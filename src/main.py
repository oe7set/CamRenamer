import sys
import winreg
import subprocess
import json
import time
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


class AboutDialog(QDialog):
    """About dialog"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About CamRenamer")
        self.setFixedSize(450, 350)
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

        # OK Button
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        ok_button.setMinimumHeight(35)
        layout.addWidget(ok_button)

        self.setLayout(layout)


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

        self.setWindowTitle("CamRenamer - USB Camera Manager v1.1")
        self.setMinimumSize(950, 650)
        self.resize(1150, 750)

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
                background-color: #404040;
                color: #ffffff;
                padding: 12px;
                border: none;
                border-right: 1px solid #555555;
                font-weight: bold;
                font-size: 11px;
            }

            QHeaderView::section:hover {
                background-color: #4a90e2;
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
        self.camera_table.setMinimumHeight(300)

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
                QMessageBox.information(
                    self, "✅ Success",
                    f"Camera was successfully renamed to '{new_name}'!\n\n"
                    "💡 A system restart may be required for the changes "
                    "to take effect in all applications."
                )
                # Update table
                camera.friendly_name = new_name
                camera.name = new_name
                self.update_camera_table()
                self.statusBar().showMessage(f"Camera successfully renamed to: {new_name}")
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

    def update_camera_name_in_registry(self, camera: CameraDevice, new_name: str) -> bool:
        """Updates the camera name in the registry"""
        try:
            # Try multiple possible registry paths
            possible_keys = [
                f"SYSTEM\\CurrentControlSet\\Enum\\{camera.device_id}",
                f"SYSTEM\\CurrentControlSet\\Control\\DeviceClasses\\{{6994AD05-93EF-11D0-A3CC-00A0C9223196}}\\#{camera.device_id}#{{6994AD05-93EF-11D0-A3CC-00A0C9223196}}\\Control",
            ]

            success = False

            for registry_path in possible_keys:
                try:
                    # Open registry key
                    with winreg.OpenKey(
                            winreg.HKEY_LOCAL_MACHINE,
                            registry_path,
                            0,
                            winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY
                    ) as key:
                        # Set FriendlyName
                        winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, new_name)
                        success = True
                        break

                except (FileNotFoundError, PermissionError, OSError):
                    continue

            # If direct registry change fails, use PowerShell
            if not success:
                powershell_cmd = f"""
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                $deviceID = "{camera.device_id}"
                $newName = "{new_name}"

                # Try registry update via PowerShell
                try {{
                    $regPath = "HKLM:\\SYSTEM\\CurrentControlSet\\Enum\\$deviceID"
                    if (Test-Path $regPath) {{
                        Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value $newName -Force
                        Write-Output "SUCCESS"
                    }} else {{
                        Write-Output "PATH_NOT_FOUND"
                    }}
                }} catch {{
                    Write-Output "ERROR: $($_.Exception.Message)"
                }}
                """

                result = subprocess.run(
                    ["powershell", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd],
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    timeout=15
                )

                if result.returncode == 0 and "SUCCESS" in result.stdout:
                    success = True

            return success

        except subprocess.TimeoutExpired:
            print("Timeout during registry update")
            return False
        except Exception as e:
            print(f"Error during registry update: {e}")
            return False

    def closeEvent(self, event):
        """Called when the application is closed"""
        if self.scanner_thread and self.scanner_thread.isRunning():
            self.scanner_thread.terminate()
            self.scanner_thread.wait(3000)
        event.accept()


def main():
    """Main function"""
    app = QApplication(sys.argv)

    # Set app properties
    app.setApplicationName("CamRenamer")
    app.setApplicationVersion("1.1")
    app.setOrganizationName("Retroverse")
    app.setOrganizationDomain("retroverse.de")

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
