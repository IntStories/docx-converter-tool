# DOCX Conversion & Zipping Tool

A Windows-friendly tool to convert Microsoft Word `.docx` files into multiple formats — HTML, ODT, EPUB, PDF, and a custom TXT with paragraph spacing — and bundle all outputs into a single ZIP archive.

The tool features a simple GUI for selecting input and output locations and uses an internally bundled Pandoc executable to minimise setup hassle.

---

## Features

- Converts `.docx` to **HTML**, **ODT**, **EPUB**, **PDF**, and **TXT**  
- Custom TXT output preserves paragraph spacing for readability  
- Pandoc is bundled *inside* the Windows executable, so no separate installation is required  
- Uses Microsoft Word (via `docx2pdf`) to generate PDF while preserving style  
- User-friendly file open and save dialogs for seamless workflow  
- Packages all converted files into one ZIP archive at your chosen location  

---

## Available Versions

- **Source code** (`.py` script) — requires Python 3.6+, and dependencies installed via `pip`  
- **Standalone Windows executable** (`.exe`) — runs without Python installed, includes Pandoc bundled internally  

---

## Requirements

### For source code

- Python 3.6 or newer  
- Microsoft Word installed (for PDF conversion)  
- Pandoc installed and added to your system PATH (if not using bundled executable)  
- Python packages: `pypandoc`, `python-docx`, `docx2pdf`, `tkinter` (usually included with Python)  

Install Python dependencies via:

```bash
pip install pypandoc python-docx docx2pdf
