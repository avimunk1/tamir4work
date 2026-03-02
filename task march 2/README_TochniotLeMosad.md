# Tochniot Le Mosad – Windows Exe

Creates the **"תוכניות למוסד"** sheet in Excel files that have **מוסדות** and **תוכניות** sheets.

## Build the Windows Exe

**You must run the build on Windows** (PyInstaller builds for the current OS).

1. On Windows, open Command Prompt or PowerShell in the project folder.
2. Run:
   ```bash
   uv run python scripts/build_tochniot_exe.py
   ```
   Or directly: `uv run pyinstaller TochniotLeMosad_win.spec`
3. The exe will be at: `dist/TochniotLeMosad.exe`

### Build options

- **Default**: No console window (double-click to run, file dialog opens).
- **With console**: `uv run python scripts/build_tochniot_exe.py --console` – keeps a console window for debugging.

## Using the Exe

1. Copy `TochniotLeMosad.exe` to any folder.
2. Double-click it (or run from command line).
3. Select your Excel file in the file dialog.
4. The file is updated in place with the new sheet. A success message appears when done.

### Command line

You can also pass the file path:

```
TochniotLeMosad.exe "C:\path\to\your\file.xlsx"
```

## Run as Python script (without building)

```bash
uv run python -m excel_extractor.tochniot_le_mosad
# or with file path:
uv run python -m excel_extractor.tochniot_le_mosad "task march 2/תוכניות למוסדות דרוזים.xlsx"
```

## Requirements

The Excel file must have:

- Sheet **"מוסדות"** with columns: שם להקמה, סמל להקמה, רשות להקמה
- Sheet **"תוכניות"** with column: תוכנית

The new sheet **"תוכניות למוסד"** will have one row per institution × program combination.
