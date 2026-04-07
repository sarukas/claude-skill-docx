# DOCX Editing Patterns Reference

Consolidated patterns for programmatic DOCX editing. Load this reference when building bespoke editing scripts.

## Core Principle: python-docx Drops Unknown Parts

python-docx only preserves parts it recognizes. When you `Document.save()`, any non-standard parts (like `comments.xml`, `commentsExtended.xml`) are silently dropped.

**Rule**: Always inject custom XML parts AFTER the python-docx save, at ZIP level.

**Correct pipeline**:
1. Open with `Document()`, modify paragraph XML (add comment markers, edit text, etc.)
2. Save with `doc.save()` -- preserves styles, headers, images, etc.
3. Patch the ZIP to add `comments.xml`, `commentsExtended.xml`, relationships, content types

## Adding Comments to DOCX

### Three XML pieces required:

**1. Comment markers in `word/document.xml`** (inside the target paragraph):
```xml
<w:commentRangeStart w:id="100"/>
  <w:r>...(runs to be commented)...</w:r>
<w:commentRangeEnd w:id="100"/>
<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="100"/>
</w:r>
```

**2. Comment content in `word/comments.xml`**:
```xml
<w:comments xmlns:w="..." xmlns:r="...">
  <w:comment w:id="100" w:author="Name" w:date="2026-01-01T00:00:00Z" w:initials="N">
    <w:p w14:paraId="4A6F7B8C">
      <w:r><w:t>Comment text here</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>
```

**3. Resolved state in `word/commentsExtended.xml`** (Word 2013+):
```xml
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                mc:Ignorable="w15">
  <w15:commentEx w15:paraId="4A6F7B8C" w15:done="1"/>
</w15:commentsEx>
```

### Key: Resolved state is NOT on `w:comment`

- `w15:done="1"` on `w:comment` element is **ignored by Word**
- Resolved state lives in a **separate part**: `word/commentsExtended.xml`
- Linked by `w14:paraId` on the comment's `<w:p>` element matching `w15:paraId` in commentsExtended
- `done="0"` = open, `done="1"` = resolved

### ZIP-level additions needed:

**Relationship** in `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rIdN"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
  Target="comments.xml"/>
<Relationship Id="rIdN+1"
  Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
  Target="commentsExtended.xml"/>
```

**Content types** in `[Content_Types].xml`:
```xml
<Override PartName="/word/comments.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
<Override PartName="/word/commentsExtended.xml"
  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
```

## Comment Marker Placement

### Correct: wrap the paragraph content runs
```
<w:pPr>...</w:pPr>           -- paragraph properties (keep outside range)
<w:commentRangeStart/>        -- BEFORE first run
<w:r>..text runs..</w:r>      -- content being commented
<w:commentRangeEnd/>          -- AFTER last run
<w:r><w:commentReference/></w:r>  -- reference run (after rangeEnd)
```

### Wrong: empty range or misplaced markers
- `commentRangeStart` immediately followed by `commentRangeEnd` with no runs between = comment attaches to wrong location
- Using `etree.SubElement(para, ...)` appends at end -- fine for rangeEnd but NOT for rangeStart which must go before the first run

### Correct insertion approach:
```python
children = list(para)
runs = [c for c in children if c.tag == f"{{{W}}}r"]
first_run_idx = children.index(runs[0])
para.insert(first_run_idx, range_start)  # Before first run

children = list(para)  # Refresh after insert!
last_run_idx = children.index(runs[-1])
para.insert(last_run_idx + 1, range_end)  # After last run
```

**Always refresh `children = list(para)` after each insert** -- indices shift.

## Editing Paragraph Text

### Preserving formatting while replacing text:
1. Capture `rPr` (run properties) from first text-bearing run via `deepcopy`
2. Remove ALL `<w:r>` elements from the paragraph
3. Create one new `<w:r>` with captured `rPr` and new text
4. Non-run elements (`pPr`, `commentRangeStart`, etc.) are preserved

### Adding bullet points by cloning:
- `deepcopy` of an existing bullet paragraph copies ALL runs (multiple text fragments)
- **Must remove all runs from clone**, then create one fresh run with new text
- Only keep `pPr` (paragraph properties) from the clone for formatting/indentation

## Adding Page Break + New Content
```python
from docx.enum.text import WD_BREAK
p_break = doc.add_paragraph()
run_break = p_break.add_run()
run_break.add_break(WD_BREAK.PAGE)
p_new = doc.add_paragraph('New page text')
```

## Document Structure Facts
- Template-generated proposals: ~320-340 paragraphs, ~15 tables, 1 section
- `doc.paragraphs` gives flat list of all paragraphs in document body
- Tables are separate from paragraphs in the element tree
- Headers/footers are in separate parts (`header1.xml`, `footer1.xml`)
- python-docx preserves all template parts (headers, footers, images, styles) on save

## Use High Comment IDs
- Use IDs like 100+ to avoid collisions with orphaned markers from previous edits
- Always clean orphaned markers before adding new ones
- Check for existing `comments.xml` in ZIP and skip/replace it

## Namespaces Reference
- `W` = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
- `R` = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
- `W14` = `http://schemas.microsoft.com/office/word/2010/wordml` (paraId)
- `W15` = `http://schemas.microsoft.com/office/word/2012/wordml` (commentsExtended)
- `MC` = `http://schemas.openxmlformats.org/markup-compatibility/2006`

## Tracked Changes: Minimal, Precise Edits

When implementing tracked changes, only mark text that actually changes. Repeating unchanged text makes edits harder to review.

Break replacements into: [unchanged text] + [deletion] + [insertion] + [unchanged text]

Preserve the original run's RSID for unchanged text by extracting the `<w:r>` element from the original and reusing it.

### Example

Changing "30 days" to "60 days" in a sentence:

```xml
<!-- BAD - Replaces entire sentence -->
<w:del><w:r><w:delText>The term is 30 days.</w:delText></w:r></w:del>
<w:ins><w:r><w:t>The term is 60 days.</w:t></w:r></w:ins>

<!-- GOOD - Only marks what changed -->
<w:r w:rsidR="00AB12CD"><w:t>The term is </w:t></w:r>
<w:del><w:r><w:delText>30</w:delText></w:r></w:del>
<w:ins><w:r><w:t>60</w:t></w:r></w:ins>
<w:r w:rsidR="00AB12CD"><w:t> days.</w:t></w:r>
```
