# Automatic CATPart-to-CATDrawing Batch Converter

A Python + Tkinter automation tool for CATIA that batch-generates CATDrawing draft files from CATPart files using a BOM Excel sheet.

## Overview

This tool helps you generate CATDrawing files in bulk with a simple UI.

It reads part metadata from an Excel BOM, opens each target CATPart in CATIA, creates drawing views automatically, optionally runs a CATScript for drawing frame generation, and saves the output to a selected folder.

## Key Features

- Batch process CATPart files to CATDrawing files
- Simple Tkinter UI (no command-line setup required for normal use)
- Row range control (`Start row` and `End row`) for partial batch runs
- Automatic view generation (isometric, front, and unfolded view when needed)
- Scale auto-selection based on CATIA parameter `L`
- Per-row error handling (failed rows are skipped, batch continues)
- Real-time status log in UI

## Required Files

You need these three file types:

1. BOM Excel file (`.xlsx`)
2. Target CATPart file(s) (`.CATPart`)
3. CATDrawing CATScript file (`.CATScript`) for automatic frame generation

## Project Structure

```text
Auto CATDrawing Generation/
|-- drawing generation_final.py
|-- BOM.xlsx
|-- CATDrawing_Template.CATScript
|-- CATfile/
```

## Requirements

- Windows OS
- CATIA installed and accessible via COM (`CATIA.Application`)
- Python 3.9+ (recommended)
- Python packages:
  - `pandas`
  - `pywin32`

Install dependencies:

```powershell
pip install pandas pywin32
```

## BOM Format Notes

The script reads the BOM using fixed column positions.

- Column C (index 2): `partname`
- Column D (index 3): `quantity`
- Column E (index 4): process keyword (`翻滚` or `折弯` enables unfolded view)

Make sure your sheet name matches the value entered in the UI.

## How To Use

1. Run CATIA.
2. Start the script:

```powershell
python "drawing generation_final.py"
```

3. In the UI, fill in:
   - BOM Excel file
   - Sheet name
   - Part folder
   - Drawing output folder
   - CATScript file
   - Start row (inclusive)
   - End row (exclusive)
4. Click `Run`.
5. Monitor progress in the status panel.

## Error Handling Behavior

- If one row fails, the script logs the error and continues with the next row.
- Open CATIA drawing/part documents are closed safely per row.
- Final summary is shown in status log:
  - succeeded count
  - failed count

## Output

Generated drawing files are saved to the selected `Drawing output folder`.

## Troubleshooting

### CATIA does not open from script

- Verify CATIA is installed correctly.
- Ensure CATIA COM automation is available on your machine.

### `Excel file not found` / folder errors

- Re-check paths selected in the UI.
- Avoid network paths with permission restrictions during first test.

### Row processing errors

- Confirm BOM data type and required columns.
- Confirm corresponding `.CATPart` files exist in selected part folder.

## Publish To GitHub

Run these commands in your project folder (`Auto CATDrawing Generation`):

```powershell
git init
git add .
git commit -m "Initial commit: CATPart to CATDrawing batch converter"
```

Create a new empty repository on GitHub, then connect and push:

```powershell
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

If your repository already exists locally, skip `git init` and just add/commit/push.

## Recommended Next Improvements

- Add a `requirements.txt`
- Add persistent config save/load in UI
- Add optional error log file export (`.txt`)
- Add cancellation support for long runs

## Disclaimer

Use at your own risk in production environments. Always test with a small row range first.
