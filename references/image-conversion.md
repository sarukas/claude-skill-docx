# Document to Image Conversion

Convert Word documents to images for visual analysis.

## Use Cases

- Visual inspection of document layout
- Creating thumbnails or previews
- Comparing document versions visually
- Extracting specific pages as images

## Two-Step Process

### Step 1: Convert DOCX to PDF

```bash
soffice --headless --convert-to pdf document.docx
```

**Requirements**: LibreOffice installed
- Linux: `sudo apt-get install libreoffice`
- Mac: Download from libreoffice.org
- Windows: Download from libreoffice.org

### Step 2: Convert PDF Pages to Images

```bash
pdftoppm -jpeg -r 150 document.pdf page
```

This creates files like `page-1.jpg`, `page-2.jpg`, etc.

**Requirements**: Poppler utilities
- Linux: `sudo apt-get install poppler-utils`
- Mac: `brew install poppler`
- Windows: Download from poppler.freedesktop.org

## Options and Customization

### Resolution

```bash
-r 150    # 150 DPI (default, good for screen viewing)
-r 300    # 300 DPI (better quality, larger files)
-r 72     # 72 DPI (lower quality, smaller files)
```

### Output Format

```bash
-jpeg     # JPEG format (smaller files)
-png      # PNG format (better quality, larger files)
```

### Page Range

```bash
-f N      # First page to convert (e.g., -f 2 starts from page 2)
-l N      # Last page to convert (e.g., -l 5 stops at page 5)
```

### Examples

**Convert all pages to JPEG**:
```bash
pdftoppm -jpeg -r 150 document.pdf page
```

**Convert specific range (pages 2-5)**:
```bash
pdftoppm -jpeg -r 150 -f 2 -l 5 document.pdf page
```

**High-quality PNG**:
```bash
pdftoppm -png -r 300 document.pdf page
```

**Single page only**:
```bash
pdftoppm -jpeg -r 150 -f 3 -l 3 document.pdf page
```

## Output Files

The `page` prefix in the command determines the output filename:

```bash
pdftoppm -jpeg -r 150 document.pdf output
# Creates: output-1.jpg, output-2.jpg, output-3.jpg, etc.
```

## Alternative: Direct DOCX to Image

For simple cases, you can combine both steps:

```bash
# Convert to PDF
soffice --headless --convert-to pdf document.docx

# Convert PDF to images
pdftoppm -jpeg -r 150 document.pdf page
```

## User Confirmation Template

Before converting to images, ask:

> "I can convert this document to images (JPEG/PNG). This requires:
> - LibreOffice (for DOCX to PDF conversion)
> - Poppler utils (for PDF to image conversion)
>
> Would you like me to proceed? What resolution do you prefer?
> - 150 DPI (standard, good for viewing)
> - 300 DPI (high quality, larger files)"

## Troubleshooting

**Error: "soffice: command not found"**
- Solution: Install LibreOffice

**Error: "pdftoppm: command not found"**
- Solution: Install Poppler utilities

**Images are too large/small**
- Solution: Adjust resolution with `-r` flag (72, 150, 300 DPI)

**Need specific pages only**
- Solution: Use `-f` (first) and `-l` (last) flags
