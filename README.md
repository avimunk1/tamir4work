# Excel Data Extractor

Extract data from multiple Excel files and create a consolidated output file.

## Requirements

- Python 3.11+
- [uv](https://docs.astral.sh/uv/) (recommended) or pip

## Setup

```bash
uv sync
```

## Usage

1. Place your Excel files (`.xlsx`) in the `input/` folder.
2. Run the extractor:

```bash
uv run excel-extractor
```

3. Find the output in `output/extracted_data_YYYYMMDD_HHMMSS.xlsx`.

## Output

The output Excel has 3 columns:

| Column           | Source                          | Description                                      |
| ---------------- | ------------------------------- | ------------------------------------------------ |
| source file name | Filename                        | Base name of the input file (without .xlsx)      |
| smel-mosad       | Sheet "פרטים כלליים" C6 + C7   | Combined values from cells C6 and C7 with space |
| total hours      | Sheet "תוכנית עבודה" G21       | Value from cell G21                              |

## Input File Structure

Each input Excel must have:

- Sheet **"פרטים כלליים"** with cells C6 and C7
- Sheet **"תוכנית עבודה"** with cell G21

## Folders

- `input/` – Created automatically on first run. Put your source Excel files here.
- `output/` – Created automatically. Extracted data is written here.
