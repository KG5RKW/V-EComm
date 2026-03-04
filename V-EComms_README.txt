# V-EComm

V-EComm is a lightweight, offline-friendly **EmComm forms + messaging workspace** built for radio operators who need fast, consistent documentation during nets, incidents, exercises, and field ops.

It provides quick access to:
- **Custom Forms** (MAGNET/ICS-style TXT forms)
- **Winlink/ARC reference forms** (as provided in the included folders)
- A simple workflow that stays **portable** (USB-ready) and **works offline**

---

## Quick Start (Windows / Linux)

### Run from source (requires Python)
1. Install Python 3.10+  
2. Open a terminal in this folder (the one that contains `V-EComm.py`)
3. Install dependencies:
   
   pip install -r requirements.txt
   
4. Launch:
   
   python V-EComm.py
   

---

## Folder Layout

Keep these folders next to the app so templates are always found:

- `Custom Forms/`
- `Winlink Forms/`


---

## Build a Windows EXE (recommended for sharing)

This repo includes a build helper that creates a portable EXE.

### One-time setup

pip install pyinstaller


### Build
- Double-click: `build_exe.bat`  
**or**

build_exe.bat


### Output
Your EXE will be created here:
- `dist\V-EComm.exe`

**Icon:** The build script embeds the V-EComm logo as the EXE icon.

---

## Notes for Operators

- The app is designed to run **offline**
- Keep the folders intact so forms/templates load correctly
- If you move the app, move the entire folder together

---

## Versioning

App Name: **V-EComm**  
(Former internal naming: V-EmComms)

