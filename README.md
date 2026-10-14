# PDF Aracı

A local PDF toolkit for Windows: an offline alternative to online services like iLovePDF. Your files never leave your computer.

> The app's interface is in Turkish.

## Features

- **Edit pages:** reorder, rotate, delete, insert pages from other PDFs or images
- **Edit text:** click any text in a PDF and change it; rewritten with the same font when it's installed on your system
- **Merge / Split**
- **Compress:** three levels, with an estimated output size before saving
- **JPG → PDF / PDF → JPG**
- **PDF → Word / Word → PDF** (Word → PDF requires Microsoft Word or LibreOffice)

## Installation

1. Install [Python](https://www.python.org/downloads/) (check "Add python.exe to PATH").
2. Double-click `kurulum.bat`. It installs the dependencies, builds an executable with PyInstaller, and adds desktop and Start menu shortcuts. Installation needs an internet connection; the app itself runs fully offline.

To run without installing:

```
pip install pywebview pymupdf pdf2docx docx2pdf pywin32 pillow fonttools
```

then double-click `PDF Aracı.pyw`.

To update, run `kurulum.bat` again from the new version's folder. To uninstall, run `kaldir.bat`.

## Project structure

| File | Purpose |
| --- | --- |
| `app.py` | Window (pywebview) and every PDF operation the UI calls (PyMuPDF) |
| `metin.py` | Text editing: locating, removing and rewriting text with a matching font |
| `arayuz.html`, `stil.css` | User interface (single page, vanilla JS) |
| `kurulum.ps1` | Installs dependencies, builds the executable, creates shortcuts |

## License

[GNU AGPL v3](LICENSE). The project uses PyMuPDF, which is AGPL-licensed, so it is released under the same license.
