# V-EComm

## Overview
V-EComm is a portable Emergency Communications application designed for digital radio and EmComm workflows.  
The application is intended to operate offline, be field‑deployable, and run without complex installation.

## Features
- Offline-capable operation
- Structured EmComm forms
- Winlink-compatible workflows
- Portable configuration storage
- Cross-station operator usability
- Dark theme optimized UI

## Running (Developer Mode)
Run directly with Python:

    python v_emcomm.py

## Building Windows EXE

1. Install PyInstaller:
    pip install pyinstaller

2. Double-click:
    build_exe.bat

3. Output will appear in:
    dist/V-EComm.exe

The EXE includes the embedded application icon and required resources.

## Folder Structure
templates/  → message & form templates  
assets/     → UI graphics and icons  
config/     → user configuration files

## Distribution
Share only:
    V-EComm.exe

No Python installation required for operators.

## Notes
- Designed for portable and field deployment.
- Config files remain local to the application directory.

---
V-EComm — Emergency Communications Software
