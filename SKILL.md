---
name: docx
description: Full DOCX lifecycle - create from Markdown, read/inspect, edit with formatting preservation, add comments, validate, and export. No Pandoc dependency. Supports tracked changes, company templates, Mermaid diagrams, and 16 configurable style settings.
---

# DOCX - Document Lifecycle Management

## Decision Tree

```
User Request
|
+-- READ document
|   +-- Quick text extraction --> Use /markdown skill
|   +-- Structured analysis --> scripts/docx_inspect.py
|       +-- --text: paragraphs with indices
|       +-- --headings: heading outline
|       +-- --tables: tables as Markdown
|       +-- --comments: comments with metadata
|       +-- --tracked-changes: insertions/deletions
|       +-- --structure: document stats (default)
|
+-- CREATE new document from Markdown
|   +-- Basic: scripts/md_to_docx_py.py input.md output.docx
|   +-- With template: ... --template your-template.docx
|   +-- With title page: ... --title "Title" --date "2026-02-08"
|   +-- With TOC: ... --toc
|   +-- With style: ... --style scripts/example.style
|   +-- Node.js alternative: scripts/md_to_docx_js.mjs input.md output.docx
|   +-- To PDF: See PDF/Image Conversion section
|
+-- EDIT existing document
|   +-- Simple find/replace --> scripts/docx_find_replace.py
|   +-- Find/replace with tracked changes --> ... --track-changes
|   +-- Full rewrite --> Re-generate via md_to_docx_py.py
|   +-- Complex tracked changes --> Load references/redlining-workflow.md
|   +-- Bespoke edits --> python-docx script, load references/docx-editing-patterns.md
|   +-- NEVER: Markdown round-trip for editing existing documents
|
+-- ADD COMMENTS --> scripts/docx_add_comments.py
|
+-- REVIEW/QA
|   +-- Proofread text --> docx_inspect --text, Claude analyzes
|   +-- Check calculations --> docx_inspect --tables, Claude verifies
|   +-- Check section numbering --> docx_inspect --headings
|   +-- Find stale references --> docx_find_replace --dry-run
|
+-- VALIDATE --> scripts/docx_validate.py
```

## 1. CREATE - New Documents from Markdown

### Python Converter (Primary)

```bash
python scripts/md_to_docx_py.py input.md output.docx
python scripts/md_to_docx_py.py input.md output.docx --template your-template.docx
python scripts/md_to_docx_py.py input.md output.docx --title "Proposal" --date "2026-02-08" --toc
python scripts/md_to_docx_py.py input.md output.docx --style scripts/example.style --copyright "Your Company"
```

**Features**: Title page (auto-detected from H1 + HR), TOC, template cover page preservation, header/footer preservation, Mermaid diagrams, .dotx/.dotm support, 16 style settings.

**Title page auto-detection**: First `# H1` becomes title, paragraphs between H1 and first `---` become preamble, `---` acts as page break.

**Template cover page (Python only)**: When template has a "Title" style paragraph, the converter replaces it with the document title and preserves all template formatting, images, and layout.

### Node.js Converter (Alternative)

```bash
node scripts/md_to_docx_js.mjs input.md output.docx
node scripts/md_to_docx_js.mjs input.md output.docx --title "Proposal" --date "2026-02-08" --toc
```

Same CLI interface. Does NOT preserve template cover pages/headers/footers (extracts styles only).

### Style Configuration (16 Settings)

| Key | Description | Default |
|-----|-------------|---------|
| `font_body` | Body text font | `Arial` |
| `font_heading` | Heading font | `Arial` |
| `font_code` | Code/monospace font | `Consolas` |
| `font_size` | Body text size (pt) | `10.5` |
| `color_heading` | Heading color (hex) | `2D3B4D` |
| `color_body` | Body text color (hex) | `333333` |
| `table_header_bg` | Table header background | `D5E8F0` |
| `table_header_text` | Table header text color | `2D3B4D` |
| `table_alt_row` | Alternating row background | `F2F2F2` |
| `table_border` | Table border color | `CCCCCC` |
| `table_border_size` | Border width (eighth-points) | `4` |
| `table_cell_margin` | Cell margin (twips) | `28` |
| `table_font_size` | Table text size (pt) | `9.5` |
| `table_banded_rows` | Banded rows on/off | `true` |
| `code_bg` | Code block background | `F5F5F5` |
| `code_font_size` | Code text size (pt) | `9` |

**Three configuration methods** (priority: defaults < template < style file < inline comment < CLI flags):

1. **Style file** (`--style example.style`): `key: value` pairs, `#` comments
2. **Inline comment** in Markdown (invisible to Obsidian):
   ```markdown
   <!-- docx-style
   font_body: Georgia
   font_size: 12
   -->
   ```
3. **CLI flags**: `--font-body "Times New Roman" --color-heading 1A3D5C`

### Mermaid Diagram Rendering

Fenced `mermaid` code blocks are rendered as images via mermaid.ink (JPEG). Falls back to local `mmdc` CLI. Both converters detect actual image format from file header bytes.

## 2. READ - Document Inspection

### docx_inspect.py

```bash
python scripts/docx_inspect.py input.docx                    # Structure summary
python scripts/docx_inspect.py input.docx --text              # Paragraphs with indices
python scripts/docx_inspect.py input.docx --headings          # Heading outline
python scripts/docx_inspect.py input.docx --tables            # Tables as Markdown
python scripts/docx_inspect.py input.docx --comments          # Comments + metadata
python scripts/docx_inspect.py input.docx --tracked-changes   # Insertions/deletions
python scripts/docx_inspect.py input.docx --text --headings   # Multiple modes
```

Outputs Markdown to stdout. For quick text extraction, use the `/markdown` skill instead.

## 3. EDIT - Modify Existing Documents

**Key principle: NEVER use Markdown round-trip for editing existing documents.** It destroys formatting, comments, tracked changes, and template structure.

### docx_find_replace.py

```bash
# Simple replacement (preserves formatting)
python scripts/docx_find_replace.py input.docx output.docx --find "Old Co" --replace "New Co"

# With tracked changes
python scripts/docx_find_replace.py input.docx output.docx --find "Old Co" --replace "New Co" \
    --track-changes --author "RIO AI"

# Dry run (report matches only)
python scripts/docx_find_replace.py input.docx output.docx --find "Old Co" --replace "New Co" --dry-run

# Scope control
python scripts/docx_find_replace.py input.docx output.docx --find "Old" --replace "New" --scope tables

# Case-insensitive, whole word
python scripts/docx_find_replace.py input.docx output.docx --find "word" --replace "term" \
    --no-case-sensitive --whole-word
```

**Scopes**: `body`, `headers`, `footers`, `tables`, `all` (default).

### Complex Tracked Changes

For complex redlining (structural changes, multiple precise edits), load `references/redlining-workflow.md` and follow the unpack -> edit -> pack workflow using the Document class.

### Bespoke Editing Scripts

For formatting changes or complex modifications, write a python-docx script. Load `references/docx-editing-patterns.md` for patterns. Key points:
- python-docx drops unknown XML parts (comments, etc.) on save
- Always inject custom XML parts AFTER python-docx save, at ZIP level
- Use high comment IDs (100+) to avoid collisions

## 4. COMMENT - Add Review Comments

### docx_add_comments.py

```bash
python scripts/docx_add_comments.py input.docx output.docx --comments comments.json
python scripts/docx_add_comments.py input.docx output.docx --comments comments.json --author "Reviewer"
```

**JSON manifest format** (`comments.json`):
```json
[
    {"anchor_text": "text to comment on", "text": "This needs revision"},
    {"anchor_text": "another phrase", "text": "Consider rewording", "resolved": true},
    {"anchor_text": "reply target", "text": "I agree", "reply_to": 0}
]
```

Fields: `anchor_text` (required), `text` (required), `resolved` (optional, default false), `reply_to` (optional, 0-based index of parent comment).

## 5. VALIDATE - Check Document Integrity

### docx_validate.py

```bash
python scripts/docx_validate.py input.docx
python scripts/docx_validate.py input.docx --check structure,comments,headings
python scripts/docx_validate.py input.docx --verbose
```

**Checks**: `structure` (ZIP, required parts, XML well-formedness), `comments` (marker pairing, orphaned markers), `headings` (hierarchy consistency), `content-types` (completeness), `schema` (OOXML XSD via existing validators).

## 6. REVIEW/QA Workflows

### Proofread Document Text
```bash
python scripts/docx_inspect.py proposal.docx --text
```
Then analyze the output for typos, grammar, consistency.

### Verify Table Calculations
```bash
python scripts/docx_inspect.py proposal.docx --tables
```
Then check arithmetic in pricing/cost tables.

### Check Section Numbering
```bash
python scripts/docx_inspect.py proposal.docx --headings
```
Verify heading hierarchy (no H1->H3 jumps) and sequential numbering.

### Find Stale References
```bash
python scripts/docx_find_replace.py proposal.docx _ --find "2024" --replace "2025" --dry-run
```
Report occurrences without modifying.

## 7. PDF/Image Conversion

### DOCX to PDF

**Windows (Word installed)**: Use docx2pdf COM automation. See `references/pdf-conversion.md`.

**Cross-platform**: `soffice --headless --convert-to pdf document.docx`

### DOCX to Images

See `references/image-conversion.md`. Two-step: DOCX -> PDF (LibreOffice) -> images (Poppler).

## 8. Templates

- **`scripts/example.style`** - Example style configuration (customize for your brand)

## 9. Dependencies

### Python (for all tools)
```bash
pip install -r requirements.txt
```
Core: `python-docx`, `lxml`, `defusedxml`, `mistune>=3.0.0`, `Pillow`, `requests`

### Node.js (for JS converter only)
```bash
npm install -g marked docx adm-zip
```

Both converters auto-install missing dependencies on first run.

## 10. Advanced Methods

Load these references on-demand for complex operations:

| Reference | When | Size |
|-----------|------|------|
| `references/ooxml-manipulation.md` | OOXML editing, Document class API | ~600 lines |
| `references/redlining-workflow.md` | Complex tracked changes workflow | ~140 lines |
| `references/docx-editing-patterns.md` | python-docx patterns, comment injection | ~170 lines |
| `references/docx-js-creation.md` | Node.js docx npm patterns | ~500 lines |
| `references/pdf-conversion.md` | DOCX to PDF methods | ~70 lines |
| `references/image-conversion.md` | DOCX to images (LibreOffice + Poppler) | ~130 lines |

### OOXML Pack/Unpack Tools

For low-level XML editing:
```bash
python ooxml/scripts/unpack.py document.docx unpacked_dir/    # Unpack + pretty-print XML
python ooxml/scripts/pack.py unpacked_dir/ output.docx        # Repack (condenses XML)
python ooxml/scripts/validate.py unpacked_dir/ --original document.docx  # Schema validation
```

## 11. Troubleshooting

**"Missing dependency"**: Run `pip install -r requirements.txt`

**Template not applied**: Ensure `--template` path points to your `.docx` or `.dotx` template file

**Mermaid diagram failed**: Install `mmdc` CLI: `npm install -g @mermaid-js/mermaid-cli`

**Comments not visible in Word**: Verify `word/comments.xml` exists in the ZIP and relationships are set

**Tracked changes not showing**: Check that `w:ins`/`w:del` elements have `w:author` and `w:date` attributes

**docx_find_replace misses text**: Text may be split across multiple `<w:r>` elements. Use `docx_inspect --text` to verify text is present, then try without `--whole-word`

**Heading hierarchy warning**: Fix source Markdown to avoid H1->H3 jumps (use H2 in between)
