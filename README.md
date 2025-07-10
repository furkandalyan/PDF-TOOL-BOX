# PDF Toolbox

A modern, feature-rich, and visually appealing PDF Toolbox built with Python and ttkbootstrap.

## Features
- Merge multiple PDFs with reordering
- Delete specific pages from a PDF
- Split PDFs by range, every N pages, equal parts, or bookmarks
- Batch encrypt, decrypt, and rotate PDFs
- OCR (text extraction) from scanned PDF pages
- Edit PDF metadata (title, author, subject, keywords)
- Undo/Redo for file operations
- Modern dark UI with ttkbootstrap

## Requirements
- Python 3.8+
- pip install -r requirements.txt (see below)

### Main dependencies
- PyPDF2
- fitz (PyMuPDF)
- Pillow
- pytesseract
- ttkbootstrap

## Usage
1. Install dependencies:
   ```sh
   pip install PyPDF2 pymupdf pillow pytesseract ttkbootstrap
   ```
2. (Optional) Install Tesseract OCR for OCR features:
   - Windows: Download from https://github.com/tesseract-ocr/tesseract
   - Add Tesseract to your PATH
3. Run the app:
   ```sh
   python pdf_toolbox.py
   ```
4. Enjoy a modern PDF toolbox for all your PDF needs!

---

**Note:** Drag-and-drop is not supported in the modern UI version. Use the file selectors for all operations.
