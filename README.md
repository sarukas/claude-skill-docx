# claude-skill-docx

A Claude Code skill for the full DOCX document lifecycle: create from Markdown, read/inspect, edit with formatting preservation, add comments, validate, and export. No Pandoc dependency.

## Installation

Clone into your project's `.claude/skills/` directory:

```bash
git clone https://github.com/sarukas/claude-skill-docx .claude/skills/docx
```

Or using the [skill](https://github.com/anthropics/skill) CLI:

```bash
npx skill install sarukas/claude-skill-docx
```

Claude Code automatically discovers skills from `SKILL.md` files in `.claude/skills/`.

### Dependencies

**Python** (required for all tools):

```bash
pip install -r requirements.txt
```

Core: `python-docx`, `lxml`, `defusedxml`, `mistune>=3.0.0`, `Pillow`, `requests`

**Node.js** (optional, for JS converter only):

```bash
npm install -g marked docx adm-zip
```

## Features

### Create - Markdown to DOCX

```bash
python scripts/md_to_docx_py.py input.md output.docx
python scripts/md_to_docx_py.py input.md output.docx --template template.docx
python scripts/md_to_docx_py.py input.md output.docx --title "Proposal" --date "2026-02-08" --toc
python scripts/md_to_docx_py.py input.md output.docx --style scripts/example.style --copyright "ACME Inc"
```

- Title page auto-detection (H1 + preamble + `---`)
- Table of contents generation
- Template support (.docx, .dotx, .dotm) with cover page preservation
- Mermaid diagram rendering (via mermaid.ink or local `mmdc`)
- 16 configurable style settings via file, inline comment, or CLI flags

### Read - Document Inspection

```bash
python scripts/docx_inspect.py input.docx                    # Structure summary
python scripts/docx_inspect.py input.docx --text              # Paragraphs with indices
python scripts/docx_inspect.py input.docx --headings          # Heading outline
python scripts/docx_inspect.py input.docx --tables            # Tables as Markdown
python scripts/docx_inspect.py input.docx --comments          # Comments + metadata
python scripts/docx_inspect.py input.docx --tracked-changes   # Insertions/deletions
```

### Edit - Find and Replace

```bash
python scripts/docx_find_replace.py input.docx output.docx --find "Old" --replace "New"
python scripts/docx_find_replace.py input.docx output.docx --find "Old" --replace "New" \
    --track-changes --author "Reviewer"
python scripts/docx_find_replace.py input.docx output.docx --find "Old" --dry-run
```

Options: `--scope` (body/headers/footers/tables/all), `--no-case-sensitive`, `--whole-word`

### Comment - Add Review Comments

```bash
python scripts/docx_add_comments.py input.docx output.docx --comments comments.json
```

JSON format:
```json
[
    {"anchor_text": "text to comment on", "text": "This needs revision"},
    {"anchor_text": "another phrase", "text": "Consider rewording", "resolved": true},
    {"anchor_text": "reply target", "text": "I agree", "reply_to": 0}
]
```

### Validate - Check Document Integrity

```bash
python scripts/docx_validate.py input.docx
python scripts/docx_validate.py input.docx --check structure,comments,headings
python scripts/docx_validate.py input.docx --verbose
```

## Style Configuration

16 settings configurable via three methods (priority: defaults < style file < inline comment < CLI flags):

| Key | Description | Default |
|-----|-------------|---------|
| `font_body` | Body text font | Arial |
| `font_heading` | Heading font | Arial |
| `font_code` | Code/monospace font | Consolas |
| `font_size` | Body text size (pt) | 10.5 |
| `color_heading` | Heading color (hex) | 2D3B4D |
| `color_body` | Body text color (hex) | 333333 |
| `table_header_bg` | Table header background | D5E8F0 |
| `table_header_text` | Table header text color | 2D3B4D |
| `table_alt_row` | Alternating row background | F2F2F2 |
| `table_border` | Table border color | CCCCCC |
| `table_border_size` | Border width (eighth-points) | 4 |
| `table_cell_margin` | Cell margin (twips) | 28 |
| `table_font_size` | Table text size (pt) | 9.5 |
| `table_banded_rows` | Banded rows on/off | true |
| `code_bg` | Code block background | F5F5F5 |
| `code_font_size` | Code text size (pt) | 9 |

**Style file** (`--style example.style`): simple `key: value` pairs with `#` comments.

**Inline comment** in Markdown (invisible to renderers):
```markdown
<!-- docx-style
font_body: Georgia
font_size: 12
-->
```

**CLI flags**: `--font-body "Times New Roman" --color-heading 1A3D5C`

See `scripts/example.style` for a complete example.

## Node.js Converter (Alternative)

```bash
node scripts/md_to_docx_js.mjs input.md output.docx
node scripts/md_to_docx_js.mjs input.md output.docx --title "Proposal" --date "2026-02-08" --toc
```

Same CLI interface as the Python converter. Does not preserve template cover pages/headers/footers.

## Advanced Tools

### OOXML Pack/Unpack

For low-level XML editing of DOCX internals:

```bash
python ooxml/scripts/unpack.py document.docx unpacked_dir/
python ooxml/scripts/pack.py unpacked_dir/ output.docx
python ooxml/scripts/validate.py unpacked_dir/ --original document.docx
```

### Reference Documentation

| Reference | Topic |
|-----------|-------|
| `references/ooxml-manipulation.md` | OOXML editing, Document class API |
| `references/redlining-workflow.md` | Complex tracked changes workflow |
| `references/docx-editing-patterns.md` | python-docx patterns, comment injection |
| `references/docx-js-creation.md` | Node.js docx npm patterns |
| `references/pdf-conversion.md` | DOCX to PDF methods |
| `references/image-conversion.md` | DOCX to images |

## Tests

Run the 65-test suite:

```bash
python tests/run_tests.py
```

Covers all tools across 7 modules: creation (20 tests), inspection (7), find/replace (8), comments (6), validation (7), markdown conversion (14), and Node.js converter (3).

## Architecture

```
SKILL.md                    # Claude Code skill definition
requirements.txt            # Python dependencies
scripts/
  md_to_docx_py.py          # Markdown to DOCX (Python, primary)
  md_to_docx_js.mjs         # Markdown to DOCX (Node.js, alternative)
  docx_inspect.py           # Document inspection/analysis
  docx_find_replace.py      # Find and replace with tracked changes
  docx_add_comments.py      # Add comments from JSON manifest
  docx_validate.py          # Document integrity validation
  document.py               # Document class for OOXML manipulation
  utilities.py              # Shared utilities
  example.style             # Example style configuration
  templates/                # XML templates for comment injection
references/                 # On-demand reference documentation
ooxml/                      # OOXML schemas and pack/unpack tools
tests/
  run_tests.py              # Comprehensive test suite (65 tests)
```

## License

MIT
