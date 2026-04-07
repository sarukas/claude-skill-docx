# Redlining Workflow for Document Review

This workflow allows you to plan comprehensive tracked changes using markdown before implementing them in OOXML.

**CRITICAL**: For complete tracked changes, you must implement ALL changes systematically.

## When to Use This Workflow

Use this method when you need to:
- Add tracked changes (insertions/deletions) to a document
- Implement document review comments with change tracking
- Make edits that need to be visible as "suggestions" rather than direct changes

**⚠️ Do NOT use for simple content edits** - use the primary Markdown workflow instead.

## Batching Strategy

Group related changes into batches of 3-10 changes. This makes debugging manageable while maintaining efficiency. Test each batch before moving to the next.

## Principle: Minimal, Precise Edits

When implementing tracked changes, only mark text that actually changes. Repeating unchanged text makes edits harder to review and appears unprofessional.

Break replacements into: [unchanged text] + [deletion] + [insertion] + [unchanged text]

Preserve the original run's RSID for unchanged text by extracting the `<w:r>` element from the original and reusing it.

### Example

Changing "30 days" to "60 days" in a sentence:

```python
# BAD - Replaces entire sentence
'<w:del><w:r><w:delText>The term is 30 days.</w:delText></w:r></w:del><w:ins><w:r><w:t>The term is 60 days.</w:t></w:r></w:ins>'

# GOOD - Only marks what changed, preserves original <w:r> for unchanged text
'<w:r w:rsidR="00AB12CD"><w:t>The term is </w:t></w:r><w:del><w:r><w:delText>30</w:delText></w:r></w:del><w:ins><w:r><w:t>60</w:t></w:r></w:ins><w:r w:rsidR="00AB12CD"><w:t> days.</w:t></w:r>'
```

## Tracked Changes Workflow

### Step 1: Inspect the Document

Analyze document structure and content:

```bash
python scripts/docx_inspect.py path-to-file.docx --text --headings
```

For a quick text overview, use the /markdown skill to convert to Markdown.

### Step 2: Identify and Group Changes

Review the document and identify ALL changes needed, organizing them into logical batches.

**Location Methods** (for finding changes in XML):
- Section/heading numbers (e.g., "Section 3.2", "Article IV")
- Paragraph identifiers if numbered
- Grep patterns with unique surrounding text
- Document structure (e.g., "first paragraph", "signature block")
- **DO NOT use markdown line numbers** - they don't map to XML structure

**Batch Organization** (group 3-10 related changes per batch):
- By section: "Batch 1: Section 2 amendments", "Batch 2: Section 5 updates"
- By type: "Batch 1: Date corrections", "Batch 2: Party name changes"
- By complexity: Start with simple text replacements, then tackle complex structural changes
- Sequential: "Batch 1: Pages 1-3", "Batch 2: Pages 4-6"

### Step 3: Read Documentation and Unpack

1. **Read documentation**:
   - **MANDATORY - READ ENTIRE FILE**: Read `references/ooxml-manipulation.md` (~600 lines) completely from start to finish
   - **NEVER set any range limits** when reading this file
   - Pay special attention to the "Document Library" and "Tracked Change Patterns" sections

2. **Unpack the document**:
   ```bash
   python ooxml/scripts/unpack.py <file.docx> <output_dir>
   ```

3. **Note the suggested RSID**: The unpack script will suggest an RSID to use for your tracked changes. Copy this RSID for use in step 4.

### Step 4: Implement Changes in Batches

Group changes logically (by section, by type, or by proximity) and implement them together in a single script.

This approach:
- Makes debugging easier (smaller batch = easier to isolate errors)
- Allows incremental progress
- Maintains efficiency (batch size of 3-10 changes works well)

**Suggested Batch Groupings:**
- By document section (e.g., "Section 3 changes", "Definitions", "Termination clause")
- By change type (e.g., "Date changes", "Party name updates", "Legal term replacements")
- By proximity (e.g., "Changes on pages 1-3", "Changes in first half of document")

**For Each Batch:**

a. **Map text to XML**: Grep for text in `word/document.xml` to verify how text is split across `<w:r>` elements

b. **Create and run script**: Use `get_node` to find nodes, implement changes, then `doc.save()`
   - See **"Document Library"** section in ooxml-manipulation.md for patterns

**Note**: Always grep `word/document.xml` immediately before writing a script to get current line numbers and verify text content. Line numbers change after each script run.

### Step 5: Pack the Document

After all batches are complete, convert the unpacked directory back to .docx:

```bash
python ooxml/scripts/pack.py unpacked_dir reviewed-document.docx
```

### Step 6: Final Verification

Do a comprehensive check of the complete document:

1. **Inspect the final document for tracked changes**:
   ```bash
   python scripts/docx_inspect.py reviewed-document.docx --tracked-changes --text
   ```

2. **Validate document integrity**:
   ```bash
   python scripts/docx_validate.py reviewed-document.docx
   ```

3. **Check that no unintended changes were introduced**

## Tips

- Test each batch before moving to the next
- Keep batch size between 3-10 changes
- Use grep to verify text location before each script
- Preserve original RSIDs for unchanged text
- Document your batch groupings for tracking
- Verify comprehensively at the end

## Related Documentation

- `ooxml-manipulation.md` - Complete OOXML editing guide with Document library
- OOXML schemas in `ooxml/schemas/` - For advanced XML structure reference
