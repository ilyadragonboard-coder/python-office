# PythonOffice

A Microsoft Office-inspired desktop suite made with **Python + Tkinter**.

## Apps

- **Writer** — text editor with basic formatting, open/save and word counter.
- **Calc** — 10-column spreadsheet-like editor with CSV import/export and JSON persistence.
- **Impress** — slide editor with slide management, background colors, save/open and fullscreen-style presentation mode.

## Run in PyCharm

1. Open the `PythonOffice` folder in PyCharm.
2. Select a Python 3.10+ interpreter.
3. Open `main.py`.
4. Press **Run**.

The main application has **no mandatory third-party runtime dependencies**.

## Build the Windows EXE

The `.exe` must be built on Windows because PyInstaller does not cross-compile Windows executables from Linux/macOS.

On Windows:

```bat
python -m pip install -r requirements.txt
build_exe.bat
```

The result will be:

```text
dist\PythonOffice.exe
```

## Project structure

```text
PythonOffice/
├── main.py
├── README.md
├── LICENSE
├── requirements.txt
├── requirements-dev.txt
├── pyproject.toml
├── build_exe.bat
├── .gitignore
├── .github/
│   └── workflows/
│       └── build-windows.yml
├── assets/
└── data/
```

## License

MIT
