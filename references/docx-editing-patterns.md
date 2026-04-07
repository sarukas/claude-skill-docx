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


---

## Inline Editing of Existing DOCX Documents

When revising a finished proposal, **always edit inline** rather than regenerating from Markdown. Inline editing preserves the cover page, headers, footers, and embedded images. Regenerating from Markdown destroys all of these.

- Fresh generation: use `wordtemplate2025.docx` as the template
- Revisions to an existing document: open and edit the DOCX directly with python-docx

### 1. `find_para` Must Skip TOC Entries

`doc.paragraphs` iterates top to bottom. TOC entries (style `toc 2`, `toc 3`) appear in the paragraph list **before** the actual body headings. A naive `find_para(fragment)` will hit the TOC entry and update it instead of the real heading.

**Fix**: skip paragraphs whose style name starts with `toc`:

```python
def find_para(fragment, skip_styles=('toc',)):
    for p in doc.paragraphs:
        if any(s in p.style.name.lower() for s in skip_styles):
            continue
        if fragment in p.text:
            return p
    return None
```

After saving, remind the user to press **Ctrl+A \u2192 F9** in Word to refresh the TOC.

### 2. Heading Text Is Often Split Across Multiple Runs

Heading paragraphs frequently have text spread across several `<w:r>` elements, e.g.:

```
runs: ['3. Iteration 1 \u2014 ', 'Data Analysis', ' Focus']
```

`docx_find_replace.py` looks for a contiguous string and silently reports 0 matches. **Fix**: filter by style name, then clear all runs and write one new run:

```python
def set_heading_text(doc, fragment, new_text):
    from docx.oxml.ns import qn
    for p in doc.paragraphs:
        if 'Heading' in p.style.name and fragment in p.text:
            for r in list(p._element.findall(qn('w:r'))):
                p._element.remove(r)
            p.add_run(new_text)
            return True
    return False
```

### 3. Bullet Items: Use `List Paragraph` Style, Never `Normal` + Manual Bullet

Inserting bullet paragraphs with `style='Normal'` and a leading `\u2022` character produces plain paragraphs. The user must reformat every bullet manually in Word.

```python
# Wrong
new_para("\u2022 NCR data ingestion...")          # Normal style, manual bullet

# Correct
new_para("NCR data ingestion...", style='List Paragraph')  # proper list style, no manual bullet
```

### 4. `insert_row_before_last` \u2014 Iterate Forward, Never `reversed()`

When inserting multiple rows before the last row of a table, each call places the new row immediately before the current last row. Forward iteration produces the correct order:

```
Insert M1 \u2192 [header][M1][Total]
Insert M2 \u2192 [header][M1][M2][Total]
Insert M3 \u2192 [header][M1][M2][M3][Total]
```

Using `reversed()` produces them backwards. Always iterate the row list forward.

### 5. `set_cell` Must Clear Existing Runs Before Writing

If the cell already contains text, simply adding a new run appends to it. Remove all `<w:r>` children first:

```python
def set_cell(table, row, col, text, bold=False):
    from docx.oxml.ns import qn
    cell = table.rows[row].cells[col]
    para = cell.paragraphs[0]
    for r in list(para._element.findall(qn('w:r'))):
        para._element.remove(r)
    if text:
        run = para.add_run(text)
        run.bold = bold
```

### 6. Chaining `addnext` \u2014 Advance the Anchor Each Time

`lxml`'s `addnext` moves the element to immediately after the anchor. When inserting a sequence of paragraphs, advance the anchor after each call or all items will be inserted in reverse order:

```python
last = ref_para._element
for text in paragraphs_to_insert:
    elem = make_elem(text)
    last.addnext(elem)
    last = elem   # advance \u2014 without this, each insert goes right after ref_para
```

### 7. Element vs Paragraph Object in Helper Functions

`new_para()` and similar helpers that return `OxmlElement` objects are **not** `Paragraph` objects \u2014 they have no `._element` attribute. Guard insertion helpers against both types:

```python
def insert_after(ref, elem):
    ref_el = ref._element if hasattr(ref, '_element') else ref
    ref_el.addnext(elem)
    return elem

def insert_before(ref, elem):
    ref_el = ref._element if hasattr(ref, '_element') else ref
    ref_el.addprevious(elem)
    return elem
```

### Inline Editing Quick Checklist

- [ ] `find_para` skips `toc` style entries
- [ ] Heading updates: filter by `'Heading' in style.name`, clear all runs, write one run
- [ ] Bullet items: `style='List Paragraph'`, never `Normal` + manual `\u2022`
- [ ] Table row insertion: iterate **forward** (no `reversed()`)
- [ ] `set_cell`: clear existing runs before writing new text
- [ ] Chain `addnext`: advance `last` anchor after each insertion
- [ ] Remind user: **Ctrl+A \u2192 F9** in Word to refresh TOC after saving
