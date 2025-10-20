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
    """Datenklasse für USB-Kamera-Informationen"""
    name: str
    device_id: str
    registry_path: str
    friendly_name: str
    hardware_id: str
    is_connected: bool = True


class CameraScanner(QThread):
    """Thread für das Scannen von USB-Kameras"""
    cameras_found = Signal(list)
    progress_updated = Signal(int)
    status_updated = Signal(str)

    def run(self):
        """Scannt nach angeschlossenen USB-Kameras"""
        try:
            self.status_updated.emit("Scanne nach USB-Kameras...")
            cameras = []

            # PowerShell-Befehl für Kamera-Erkennung mit UTF-8 Output
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

            # PowerShell ausführen mit expliziter UTF-8 Encoding
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
                            friendly_name = device.get('FriendlyName', 'Unbekannt')
                            instance_id = device.get('InstanceId', '')
                            hardware_id = device.get('HardwareID', [''])[0] if device.get('HardwareID') else ''
                            status = device.get('Status', 'Unknown')

                            # Registry-Pfad konstruieren
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
                    self.status_updated.emit(f"{len(cameras)} Kamera(s) gefunden")

                except json.JSONDecodeError as e:
                    self.status_updated.emit(f"Fehler beim Parsen der Kamera-Daten: {str(e)}")
            else:
                if result.stderr:
                    self.status_updated.emit(f"PowerShell Fehler: {result.stderr}")
                else:
                    self.status_updated.emit("Keine Kameras gefunden")

            self.cameras_found.emit(cameras)

        except subprocess.TimeoutExpired:
            self.status_updated.emit("Timeout beim Scannen der Kameras")
            self.cameras_found.emit([])
        except Exception as e:
            self.status_updated.emit(f"Fehler beim Scannen: {str(e)}")
            self.cameras_found.emit([])


class AboutDialog(QDialog):
    """Über-Dialog"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Über CamRenamer")
        self.setFixedSize(450, 350)
        self.setModal(True)

        layout = QVBoxLayout()

        # Titel
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
        author = QLabel("© 2024 Retroverse - Erwin Spitaler")
        author.setAlignment(Qt.AlignmentFlag.AlignCenter)
        author.setFont(QFont("Arial", 10))
        layout.addWidget(author)

        # Beschreibung
        desc = QTextEdit()
        desc.setReadOnly(True)
        desc.setMaximumHeight(180)
        desc.setPlainText(
            "CamRenamer ermöglicht es, USB-Kameras zu verwalten und deren Namen "
            "in der Windows-Registry zu ändern. Dies ist besonders nützlich für "
            "Streaming-Anwendungen wie OBS, wo eindeutige Kamera-Namen erforderlich sind.\n\n"
            "Funktionen:\n"
            "• Automatisches Erkennen von USB-Kameras\n"
            "• Umbenennen von Kameras in der Registry\n"
            "• Modernes Dark Theme Design\n"
            "• PowerShell und Registry Integration"
        )
        layout.addWidget(desc)

        # OK Button
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        ok_button.setMinimumHeight(35)
        layout.addWidget(ok_button)

        self.setLayout(layout)


class CamRenamerMainWindow(QMainWindow):
    """Hauptfenster der Anwendung"""

    def __init__(self):
        super().__init__()
        self.cameras: List[CameraDevice] = []
        self.scanner_thread: Optional[CameraScanner] = None

        self.setWindowTitle("CamRenamer - USB Kamera Manager v1.1")
        self.setMinimumSize(950, 650)
        self.resize(1150, 750)

        # Modernes Design anwenden
        self.apply_modern_style()

        # UI komponenten erstellen
        self.setup_ui()

        # Menu und Toolbar erstellen
        self.setup_menu_and_toolbar()

        # Enter-Taste für Umbenennung aktivieren
        self.new_name_edit.returnPressed.connect(self.rename_selected_camera)

        # Initial scan nach kurzer Verzögerung
        QTimer.singleShot(500, self.scan_cameras)

    def apply_modern_style(self):
        """Wendet ein modernes Dark Theme an"""
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
        """Erstellt Menü und Toolbar"""
        # Menübar
        menubar = self.menuBar()

        # Datei Menü
        file_menu = menubar.addMenu("&Datei")

        refresh_action = QAction("🔄 &Aktualisieren", self)
        refresh_action.setShortcut("F5")
        refresh_action.setStatusTip("Kamera-Liste neu laden")
        refresh_action.triggered.connect(self.scan_cameras)
        file_menu.addAction(refresh_action)

        file_menu.addSeparator()

        exit_action = QAction("❌ &Beenden", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Anwendung beenden")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Extras Menü
        tools_menu = menubar.addMenu("&Extras")

        clear_action = QAction("🧹 Tabelle &leeren", self)
        clear_action.setShortcut("Ctrl+L")
        clear_action.setStatusTip("Kamera-Tabelle leeren")
        clear_action.triggered.connect(self.clear_table)
        tools_menu.addAction(clear_action)

        # Hilfe Menü
        help_menu = menubar.addMenu("&Hilfe")

        about_action = QAction("ℹ️ &Über", self)
        about_action.setShortcut("F1")
        about_action.setStatusTip("Über diese Anwendung")
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        # Toolbar
        toolbar = QToolBar("Haupt-Toolbar")
        self.addToolBar(toolbar)

        toolbar.addAction(refresh_action)
        toolbar.addSeparator()
        toolbar.addAction(clear_action)
        toolbar.addSeparator()
        toolbar.addAction(about_action)

        # Status Bar
        self.statusBar().showMessage("Bereit - Drücken Sie F5 zum Aktualisieren")

    def setup_ui(self):
        """Erstellt die Benutzeroberfläche"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(25, 25, 25, 25)

        # Header
        header_layout = QHBoxLayout()

        title_label = QLabel("🎥 USB Kamera Manager")
        title_label.setFont(QFont("Arial", 20, QFont.Weight.Bold))
        header_layout.addWidget(title_label)

        header_layout.addStretch()

        self.scan_button = QPushButton("🔄 Kameras scannen")
        self.scan_button.clicked.connect(self.scan_cameras)
        self.scan_button.setMinimumHeight(40)
        header_layout.addWidget(self.scan_button)

        main_layout.addLayout(header_layout)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(25)
        main_layout.addWidget(self.progress_bar)

        # Kamera Tabelle
        camera_group = QGroupBox("📋 Gefundene USB Kameras")
        camera_layout = QVBoxLayout(camera_group)

        self.camera_table = QTableWidget()
        self.camera_table.setColumnCount(5)
        self.camera_table.setHorizontalHeaderLabels([
            "🎥 Kamera Name", "🔧 Gerät ID", "💾 Hardware ID", "🔌 Status", "⚙️ Aktionen"
        ])

        # Tabellen-Styling
        header = self.camera_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)

        self.camera_table.setAlternatingRowColors(True)
        self.camera_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.camera_table.setMinimumHeight(300)

        camera_layout.addWidget(self.camera_table)
        main_layout.addWidget(camera_group)

        # Umbenennen Bereich
        rename_group = QGroupBox("✏️ Kamera umbenennen")
        rename_layout = QVBoxLayout(rename_group)

        rename_form_layout = QHBoxLayout()

        rename_form_layout.addWidget(QLabel("📝 Neuer Name:"))

        self.new_name_edit = QLineEdit()
        self.new_name_edit.setPlaceholderText("Geben Sie den neuen Namen ein...")
        self.new_name_edit.setMinimumHeight(35)
        rename_form_layout.addWidget(self.new_name_edit)

        self.rename_button = QPushButton("✅ Umbenennen")
        self.rename_button.clicked.connect(self.rename_selected_camera)
        self.rename_button.setEnabled(False)
        self.rename_button.setMinimumHeight(35)
        rename_form_layout.addWidget(self.rename_button)

        rename_layout.addLayout(rename_form_layout)

        # Hinweis-Label
        hint_label = QLabel("💡 Tipp: Wählen Sie eine Kamera aus der Tabelle aus und geben Sie einen neuen Namen ein.")
        hint_label.setStyleSheet("color: #aaaaaa; font-size: 11px; font-style: italic;")
        rename_layout.addWidget(hint_label)

        main_layout.addWidget(rename_group)

        # Tabellen-Auswahl Ereignis
        self.camera_table.itemSelectionChanged.connect(self.on_camera_selection_changed)

    def clear_table(self):
        """Leert die Kamera-Tabelle"""
        self.cameras.clear()
        self.camera_table.setRowCount(0)
        self.new_name_edit.clear()
        self.rename_button.setEnabled(False)
        self.statusBar().showMessage("Tabelle geleert")

    def show_about(self):
        """Zeigt den Über-Dialog"""
        dialog = AboutDialog(self)
        dialog.exec()

    def on_camera_selection_changed(self):
        """Wird aufgerufen wenn eine Kamera ausgewählt wird"""
        selected_rows = self.camera_table.selectionModel().selectedRows()

        if selected_rows:
            row = selected_rows[0].row()
            if row < len(self.cameras):
                camera = self.cameras[row]
                self.new_name_edit.setText(camera.friendly_name)
                self.rename_button.setEnabled(True)
                self.statusBar().showMessage(f"Kamera ausgewählt: {camera.friendly_name}")
                self.new_name_edit.setFocus()
                self.new_name_edit.selectAll()
            else:
                self.rename_button.setEnabled(False)
        else:
            self.rename_button.setEnabled(False)
            self.new_name_edit.clear()
            self.statusBar().showMessage("Bereit")

    def scan_cameras(self):
        """Startet den Kamera-Scan"""
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
        """Wird aufgerufen wenn der Scan beendet ist"""
        self.scan_button.setEnabled(True)
        self.progress_bar.setVisible(False)

    def on_cameras_found(self, cameras: List[CameraDevice]):
        """Wird aufgerufen wenn Kameras gefunden wurden"""
        self.cameras = cameras
        self.update_camera_table()

    def update_camera_table(self):
        """Aktualisiert die Kamera-Tabelle"""
        self.camera_table.setRowCount(len(self.cameras))

        for row, camera in enumerate(self.cameras):
            # Name
            name_item = QTableWidgetItem(camera.friendly_name)
            name_item.setToolTip(f"Vollständiger Name: {camera.friendly_name}")
            self.camera_table.setItem(row, 0, name_item)

            # Device ID (gekürzt)
            device_id = camera.device_id[:50] + "..." if len(camera.device_id) > 50 else camera.device_id
            device_item = QTableWidgetItem(device_id)
            device_item.setToolTip(f"Vollständige Device ID: {camera.device_id}")
            self.camera_table.setItem(row, 1, device_item)

            # Hardware ID (gekürzt)
            hardware_id = camera.hardware_id[:30] + "..." if len(camera.hardware_id) > 30 else camera.hardware_id
            hardware_item = QTableWidgetItem(hardware_id)
            hardware_item.setToolTip(f"Vollständige Hardware ID: {camera.hardware_id}")
            self.camera_table.setItem(row, 2, hardware_item)

            # Status
            status_text = "🟢 Verbunden" if camera.is_connected else "🔴 Getrennt"
            status_item = QTableWidgetItem(status_text)
            status_item.setToolTip(f"Status: {'Aktiv und betriebsbereit' if camera.is_connected else 'Nicht verfügbar'}")
            self.camera_table.setItem(row, 3, status_item)

            # Aktionen
            action_item = QTableWidgetItem("👆 Zum Umbenennen auswählen")
            action_item.setToolTip("Klicken Sie auf diese Zeile, um die Kamera auszuwählen")
            self.camera_table.setItem(row, 4, action_item)

    def rename_selected_camera(self):
        """Benennt die ausgewählte Kamera um"""
        selected_rows = self.camera_table.selectionModel().selectedRows()

        if not selected_rows:
            QMessageBox.warning(self, "⚠️ Warnung", "Bitte wählen Sie eine Kamera aus der Tabelle aus.")
            return

        new_name = self.new_name_edit.text().strip()
        if not new_name:
            QMessageBox.warning(self, "⚠️ Warnung", "Bitte geben Sie einen neuen Namen ein.")
            self.new_name_edit.setFocus()
            return

        if len(new_name) > 255:
            QMessageBox.warning(self, "⚠️ Warnung", "Der Name ist zu lang. Maximal 255 Zeichen erlaubt.")
            return

        row = selected_rows[0].row()
        camera = self.cameras[row]

        if new_name == camera.friendly_name:
            QMessageBox.information(self, "ℹ️ Information", "Der neue Name ist identisch mit dem aktuellen Namen.")
            return

        # Bestätigung
        reply = QMessageBox.question(
            self, "❓ Bestätigung",
            f"Möchten Sie '{camera.friendly_name}' wirklich zu '{new_name}' umbenennen?\n\n"
            "⚠️ Hinweis: Dies erfordert Administratorrechte und kann einen Neustart erfordern.\n"
            "Die Änderung wird in der Windows-Registry gespeichert.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # UI während Umbenennung blockieren
            self.setEnabled(False)
            self.statusBar().showMessage("Benenne Kamera um...")

            success = self.update_camera_name_in_registry(camera, new_name)

            self.setEnabled(True)

            if success:
                QMessageBox.information(
                    self, "✅ Erfolg",
                    f"Kamera wurde erfolgreich zu '{new_name}' umbenannt!\n\n"
                    "💡 Ein Neustart des Systems könnte erforderlich sein, damit die Änderungen "
                    "in allen Anwendungen wirksam werden."
                )
                # Tabelle aktualisieren
                camera.friendly_name = new_name
                camera.name = new_name
                self.update_camera_table()
                self.statusBar().showMessage(f"Kamera erfolgreich umbenannt zu: {new_name}")
            else:
                QMessageBox.critical(
                    self, "❌ Fehler",
                    "Fehler beim Umbenennen der Kamera!\n\n"
                    "Mögliche Ursachen:\n"
                    "• Keine Administratorrechte\n"
                    "• Kamera wird gerade verwendet\n"
                    "• Registry-Zugriff verweigert\n\n"
                    "Versuchen Sie, die Anwendung als Administrator zu starten."
                )
                self.statusBar().showMessage("Fehler beim Umbenennen")

    def update_camera_name_in_registry(self, camera: CameraDevice, new_name: str) -> bool:
        """Aktualisiert den Kamera-Namen in der Registry"""
        try:
            # Mehrere mögliche Registry-Pfade versuchen
            possible_keys = [
                f"SYSTEM\\CurrentControlSet\\Enum\\{camera.device_id}",
                f"SYSTEM\\CurrentControlSet\\Control\\DeviceClasses\\{{6994AD05-93EF-11D0-A3CC-00A0C9223196}}\\#{camera.device_id}#{{6994AD05-93EF-11D0-A3CC-00A0C9223196}}\\Control",
            ]

            success = False

            for registry_path in possible_keys:
                try:
                    # Registry-Schlüssel öffnen
                    with winreg.OpenKey(
                            winreg.HKEY_LOCAL_MACHINE,
                            registry_path,
                            0,
                            winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY
                    ) as key:
                        # FriendlyName setzen
                        winreg.SetValueEx(key, "FriendlyName", 0, winreg.REG_SZ, new_name)
                        success = True
                        break

                except (FileNotFoundError, PermissionError, OSError):
                    continue

            # Falls direkte Registry-Änderung fehlschlägt, PowerShell verwenden
            if not success:
                powershell_cmd = f"""
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                $deviceID = "{camera.device_id}"
                $newName = "{new_name}"

                # Versuche Registry-Update über PowerShell
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
            print("Timeout beim Registry-Update")
            return False
        except Exception as e:
            print(f"Fehler beim Registry-Update: {e}")
            return False

    def closeEvent(self, event):
        """Wird beim Schließen der Anwendung aufgerufen"""
        if self.scanner_thread and self.scanner_thread.isRunning():
            self.scanner_thread.terminate()
            self.scanner_thread.wait(3000)
        event.accept()


def main():
    """Hauptfunktion"""
    app = QApplication(sys.argv)

    # App-Eigenschaften setzen
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

    # Titel zeichnen
    title_rect = splash_pixmap.rect()
    title_rect.setTop(80)
    title_rect.setHeight(50)
    painter.drawText(title_rect, Qt.AlignmentFlag.AlignCenter, "🎥 CamRenamer")

    # Untertitel zeichnen
    painter.setFont(QFont("Arial", 14))
    subtitle_rect = splash_pixmap.rect()
    subtitle_rect.setTop(140)
    subtitle_rect.setHeight(30)
    painter.drawText(subtitle_rect, Qt.AlignmentFlag.AlignCenter, "USB Kamera Manager v1.1")

    # Status zeichnen
    painter.setFont(QFont("Arial", 12))
    status_rect = splash_pixmap.rect()
    status_rect.setTop(200)
    status_rect.setHeight(30)
    painter.drawText(status_rect, Qt.AlignmentFlag.AlignCenter, "Lädt...")

    painter.end()

    splash = QSplashScreen(splash_pixmap)
    splash.show()

    # Kurze Wartezeit für Splash
    for i in range(10):
        time.sleep(0.1)
        app.processEvents()

    # Hauptfenster erstellen
    window = CamRenamerMainWindow()

    # Splash schließen und Hauptfenster anzeigen
    splash.finish(window)
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
