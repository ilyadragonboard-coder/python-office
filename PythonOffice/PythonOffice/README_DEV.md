# Development notes

## PyCharm
Recommended interpreter: CPython 3.12 on Windows.

## Architecture
The app intentionally keeps all UI in `main.py` for easy study/editing in PyCharm.
As the project grows, split the classes into:

- `app/`
- `writer/`
- `calc/`
- `impress/`
- `common/`

without changing the user-facing file formats.

## Why Tkinter?
Tkinter ships with regular desktop Python installations and keeps the project simple to run from PyCharm.
