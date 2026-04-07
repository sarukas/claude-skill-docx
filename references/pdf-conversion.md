# PDF Conversion

Convert DOCX files to PDF format.

## Method 1: Word COM Automation (Windows)

On Windows with Microsoft Word installed, use COM automation:

```python
import subprocess
import sys
from pathlib import Path

def docx_to_pdf(input_path, output_path=None):
    """Convert DOCX to PDF using Word COM automation via docx2pdf."""
    try:
        import docx2pdf
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "docx2pdf"])
        import docx2pdf

    input_path = Path(input_path)
    if output_path is None:
        output_path = input_path.with_suffix(".pdf")
    docx2pdf.convert(str(input_path), str(output_path))
    return output_path
```

**Requirements**: Windows + Microsoft Word installed.

## Method 2: LibreOffice (Cross-platform)

```bash
soffice --headless --convert-to pdf document.docx
```

**Requirements**: LibreOffice installed
- Linux: `sudo apt-get install libreoffice`
- Windows/Mac: Download from libreoffice.org

**Pros**: Cross-platform, free, no Word dependency.
**Cons**: May have slight formatting differences from Word.

## When to Use

- User specifically requests PDF output
- After generating a DOCX via `md_to_docx_py.py`

## User Confirmation Template

Before using PDF conversion, ask:

> "I can convert this document to PDF. This requires:
> - Windows with Microsoft Word installed (via docx2pdf), OR
> - LibreOffice installed (cross-platform alternative)
>
> Which method should I use?"

## Troubleshooting

**Error: "Microsoft Word is not installed"**
- Use LibreOffice method instead

**Error: "docx2pdf failed"**
- Check that Word is not already open with the file
- Try LibreOffice method instead

**LibreOffice not found**
- Install from libreoffice.org or package manager
