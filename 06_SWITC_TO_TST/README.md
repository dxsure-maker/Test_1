# SWITC_TO_TST v1.7

Windows executable packaging support for `SWITC_TO_TST_v1_7.py`.

## Build the executable

Run this from this folder on Windows:

```bat
build_exe.bat
```

The build script creates a local `.venv`, installs the build dependencies, and writes the executable to:

```text
dist\SWITC_TO_TST_v1_7.exe
```

If Python is not on `PATH`, set `PYTHON_EXE` first:

```bat
set PYTHON_EXE=C:\Path\To\python.exe
build_exe.bat
```

You can also call the PowerShell script directly:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Use `-Clean` to remove prior `build` and `dist` folders before rebuilding:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build_exe.ps1 -Clean
```
