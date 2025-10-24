# 🎥 CamRenamer - USB Camera Manager

A Windows application for renaming USB cameras in the Registry for better identification in streaming software like OBS.

![CamRenamer Logo](https://via.placeholder.com/400x100/2b2b2b/ffffff?text=🎥+CamRenamer+v1.1)

## 📋 Overview

CamRenamer allows you to easily rename USB cameras by changing their names directly in the Windows Registry. This is particularly useful for:

- **Streaming**: Unique camera names in OBS, XSplit, and other applications
- **Multi-Camera Setups**: Better organization of multiple cameras
- **Professional Production**: Clear identification of cameras by location/function

## ✨ Features

- 🔍 **Automatic Camera Detection**: Finds all connected USB cameras
- ✏️ **Easy Renaming**: Intuitive user interface for changing names
- 🔒 **Registry Integration**: Direct changes in the Windows Registry
- 🌐 **Unicode Support**: Full support for international characters
- 🎨 **Modern Design**: Dark theme with user-friendly interface
- ⚡ **PowerShell Fallback**: Alternative method for registry issues
- 🛡️ **Error Handling**: Robust handling of permission errors

## 📸 Screenshots

### Main Window
```
┌─────────────────────────────────────────────────────────────┐
│ 🎥 USB Camera Manager                      🔄 Scan Cameras │
├─────────────────────────────────────────────────────────────┤
│ 📋 Found USB Cameras                                       │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ 🎥 Name    │ 🔧 Device ID │ 💾 HW ID  │ 🔌 Status    │ │
│ │ USB Camera │ USB\VID_1234 │ USB\...   │ 🟢 Connected │ │
│ └─────────────────────────────────────────────────────────┘ │
│                                                             │
│ ✏️ Rename Camera                                            │
│ ┌─────────────────────────────────────────────────────────┐ │
│ │ 📝 New Name: [Streaming Camera Left     ] ✅ Rename     │ │
│ └─────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
```

## 🚀 Installation

### Prerequisites

- **Windows 10/11** (required for Registry access)
- **Python 3.12+** (for development)
- **Administrator rights** (for Registry changes)

### Option 1: Executable File (Recommended)

1. Download the latest version from [Releases](https://github.com/retroverse/camrenamer/releases)
2. Extract the ZIP file
3. Run `CamRenamer.exe` as Administrator

### Option 2: With uv (Development)

```powershell
# Clone repository
git clone https://github.com/retroverse/camrenamer.git
cd camrenamer

# Install dependencies
uv sync

# Start application
uv run python src/main_complete.py
```

### Option 3: With pip

```powershell
# Clone repository
git clone https://github.com/retroverse/camrenamer.git
cd camrenamer

# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -e .[dev]

# Start application
python src/main_complete.py
```

## 🎮 Usage

### Step-by-Step Guide

1. **Run as Administrator**: Right-click on the application → "Run as administrator"

2. **Scan Cameras**:
   - Click on "🔄 Scan Cameras" or press F5
   - All connected USB cameras will be listed

3. **Select Camera**:
   - Click on a row in the table
   - The current name will be displayed in the input field

4. **Rename**:
   - Enter the new name
   - Click on "✅ Rename" or press Enter
   - Confirm the change

5. **Done**:
   - The change is immediately saved in the Registry
   - A restart may be required for full effectiveness

### Keyboard Shortcuts

| Keyboard Combination | Action |
|---------------------|--------|
| `F5` | Refresh camera list |
| `Ctrl+L` | Clear table |
| `Ctrl+Q` | Exit application |
| `F1` | Show about dialog |
| `Enter` | Rename selected camera |

## 🔧 Development

### Project Setup

```powershell
# Clone repository
git clone https://github.com/retroverse/camrenamer.git
cd camrenamer

# Development environment with uv
uv sync --group dev

# Or with pip
pip install -e .[dev]
```

### Running Tests

```powershell
# All tests
uv run pytest

# With Coverage
uv run pytest --cov=src --cov-report=html

# Unit tests only
uv run pytest -m "unit"

# Integration tests only
uv run pytest -m "integration"
```

### Code Quality

```powershell
# Linting with Ruff
uv run ruff check src/

# Check formatting
uv run ruff format --check src/

# Type checking with mypy
uv run mypy src/
```

### Create Build

```powershell
# Executable with PyInstaller
uv run pyinstaller --windowed --onefile --name CamRenamer src/main_complete.py

# Or with auto-py-to-exe (GUI)
uv run auto-py-to-exe
```

## 📁 Project Structure

```
CamRenamer/
├── src/
│   ├── main.py              # Original main file (incomplete)
│   └── main_complete.py     # Improved main file
├── tests/
│   ├── __init__.py
│   ├── test_camera_scanner.py
│   └── test_registry_update.py
├── docs/
│   └── README.md
├── pyproject.toml           # Original
├── pyproject_improved.toml # Improved version
├── uv.lock
├── LICENSE
└── README.md
```

## 🛠️ Technical Details

### Architecture

- **GUI Framework**: PySide6 (Qt6)
- **Threading**: QThread for asynchronous operations
- **Registry Access**: winreg + PowerShell Fallback
- **Camera Detection**: PowerShell Get-PnpDevice
- **Design**: Dark Theme with modern UI

### Registry Paths

The application searches in the following Registry paths:

```
HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Enum\{DeviceID}
HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\DeviceClasses\{GUID}
```

### Unicode Support

- UTF-8 Encoding for PowerShell output
- Error handling for invalid characters
- Support for international camera names

## ⚠️ Security Notes

- **Administrator rights required**: Registry changes require elevated permissions
- **Backup recommended**: Create a Registry backup before major changes
- **Antivirus**: Some antivirus programs may block Registry access

## 🐛 Known Issues

### Issue: "Access Denied"
**Solution**: Run application as Administrator

### Issue: Changes not visible
**Solution**: Restart system or reconnect camera

### Issue: PowerShell Error
**Solution**: Check PowerShell execution policy:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 📞 Support

- 🐛 **Bug Reports**: [GitHub Issues](https://github.com/retroverse/camrenamer/issues)
- 💡 **Feature Requests**: [GitHub Discussions](https://github.com/retroverse/camrenamer/discussions)
- 📧 **Email**: erwin@retroverse.de

## 🤝 Contributing

Contributions are welcome! Please note:

1. Fork the repository
2. Create feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'Add amazing feature'`)
4. Branch pushen (`git push origin feature/amazing-feature`)
5. Create Pull Request

### Code Guidelines

- Use Python 3.12+
- Follow PEP 8 style (automatic with Ruff)
- Write tests for new features
- Docstrings for public APIs

## 📜 License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## 🙏 Acknowledgments

- **Qt/PySide6**: For the excellent GUI framework
- **Python Community**: For the great tools and libraries
- **OBS Community**: For the inspiration for this tool

## 📊 Statistics

![GitHub release (latest by date)](https://img.shields.io/github/v/release/retroverse/camrenamer)
![GitHub](https://img.shields.io/github/license/retroverse/camrenamer)
![Python](https://img.shields.io/badge/python-3.12+-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)

---

**Developed with ❤️ by [Erwin Spitaler OE7SET](https://github.com/oe7set)**

pyinstaller --onefile --windowed --icon=src\img\icon.ico src\main.py
