# ppt-to-pdf-converter

I had a folder of PowerPoint decks I needed as one PDF. PowerPoint can export to PDF one file at a time; this does the whole folder.

![CI](https://github.com/Rey-EL/ppt-to-pdf-converter/actions/workflows/ci.yml/badge.svg)
![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)

## Features

- Finds every `.ppt` and `.pptx` in a folder, including subfolders
- Converts each deck to PDF through PowerPoint itself, so the output matches what PowerPoint would produce
- Merges all the PDFs into one file
- Converts into a temp folder that deletes itself when the run finishes
- Progress bar and status line while it works

## Install

Windows only. You need Python 3 and Microsoft PowerPoint installed.

```bash
git clone https://github.com/Rey-EL/ppt-to-pdf-converter.git
cd ppt-to-pdf-converter
pip install -r requirements.txt
```

## Usage

```bash
python ppt_to_pdf_converter.py
```

1. Click "1. Select Folder with Presentations" and pick the folder.
2. Click "2. Convert and Save As PDF..." and choose where the final file goes.

## How it works

`convert_ppt_to_pdf` drives PowerPoint through COM automation (`SaveAs` with format 32 = PDF). `main_process` converts every deck into a temporary directory, merges the results with pypdf, and writes the final file. CI runs a compile smoke test on Python 3.10–3.12 (COM automation itself needs real PowerPoint on Windows, so it is not exercised in CI).

## Project structure

```
ppt-to-pdf-converter/
├── ppt_to_pdf_converter.py   # the tool (COM conversion + tkinter GUI)
├── requirements.txt          # pywin32 and pypdf (Windows)
├── tests/                    # pytest smoke tests (compile + structure)
└── .github/workflows/ci.yml  # CI workflow
```

## License

MIT — see [LICENSE.md](LICENSE.md).
