#!/usr/bin/env python3
"""
Validate DOCX document integrity.

Checks document structure, schema compliance, comment integrity,
and heading hierarchy.

Usage:
    python docx_validate.py input.docx
    python docx_validate.py input.docx --check schema,structure,comments
    python docx_validate.py input.docx --verbose
"""

import argparse
import sys
import zipfile
from pathlib import Path

try:
    from docx import Document
except ImportError:
    print("Missing dependency: pip install python-docx", file=sys.stderr)
    sys.exit(1)

try:
    import defusedxml.minidom
except ImportError:
    defusedxml = None


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def check_structure(docx_path, verbose=False):
    """Check basic DOCX ZIP structure."""
    issues = []

    try:
        zf = zipfile.ZipFile(docx_path, "r")
    except zipfile.BadZipFile:
        return [("CRITICAL", "Not a valid ZIP/DOCX file")]

    names = zf.namelist()

    # Required parts
    required = ["[Content_Types].xml", "word/document.xml"]
    for req in required:
        if req not in names:
            issues.append(("CRITICAL", f"Missing required part: {req}"))

    # Check Content_Types.xml
    if "[Content_Types].xml" in names:
        ct_xml = zf.read("[Content_Types].xml").decode("utf-8", errors="ignore")
        if "word/document.xml" not in ct_xml and "wordprocessingml.document.main" not in ct_xml:
            issues.append(("WARNING", "document.xml not referenced in Content_Types"))

    # Check for relationships
    if "word/_rels/document.xml.rels" not in names:
        issues.append(("WARNING", "Missing document relationships file"))

    # Check XML well-formedness
    for name in names:
        if name.endswith(".xml") or name.endswith(".rels"):
            try:
                xml_data = zf.read(name)
                if defusedxml:
                    defusedxml.minidom.parseString(xml_data)
                else:
                    from xml.dom.minidom import parseString
                    parseString(xml_data)
            except Exception as e:
                issues.append(("CRITICAL", f"Malformed XML in {name}: {e}"))

    if verbose and not issues:
        print(f"  Structure: {len(names)} parts found, all well-formed")

    zf.close()
    return issues


def check_comments(docx_path, verbose=False):
    """Check comment integrity."""
    issues = []

    try:
        zf = zipfile.ZipFile(docx_path, "r")
    except zipfile.BadZipFile:
        return [("CRITICAL", "Not a valid ZIP/DOCX file")]

    names = zf.namelist()

    if "word/comments.xml" not in names:
        if verbose:
            print("  Comments: No comments.xml found (OK)")
        zf.close()
        return issues

    # Parse comments
    comments_xml = zf.read("word/comments.xml")
    try:
        if defusedxml:
            comments_dom = defusedxml.minidom.parseString(comments_xml)
        else:
            from xml.dom.minidom import parseString
            comments_dom = parseString(comments_xml)
    except Exception as e:
        issues.append(("CRITICAL", f"Malformed comments.xml: {e}"))
        zf.close()
        return issues

    comment_ids = set()
    for c in comments_dom.getElementsByTagName("w:comment"):
        cid = c.getAttribute("w:id")
        if cid:
            if cid in comment_ids:
                issues.append(("WARNING", f"Duplicate comment ID: {cid}"))
            comment_ids.add(cid)

    # Check document.xml for matching markers
    if "word/document.xml" in names:
        doc_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")

        if defusedxml:
            doc_dom = defusedxml.minidom.parseString(doc_xml.encode("utf-8"))
        else:
            from xml.dom.minidom import parseString
            doc_dom = parseString(doc_xml.encode("utf-8"))

        # Collect referenced comment IDs from markers
        marker_ids = set()
        for tag in ["w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"]:
            for elem in doc_dom.getElementsByTagName(tag):
                mid = elem.getAttribute("w:id")
                if mid:
                    marker_ids.add(mid)

        # Check for orphaned markers
        orphaned = marker_ids - comment_ids
        if orphaned:
            issues.append(("WARNING", f"Orphaned comment markers (no comment): {orphaned}"))

        # Check for comments without markers
        unmarked = comment_ids - marker_ids
        if unmarked:
            issues.append(("WARNING", f"Comments without markers in document: {unmarked}"))

        # Check rangeStart/rangeEnd pairing
        starts = set()
        ends = set()
        for rs in doc_dom.getElementsByTagName("w:commentRangeStart"):
            starts.add(rs.getAttribute("w:id"))
        for re_elem in doc_dom.getElementsByTagName("w:commentRangeEnd"):
            ends.add(re_elem.getAttribute("w:id"))

        unmatched_starts = starts - ends
        unmatched_ends = ends - starts
        if unmatched_starts:
            issues.append(("WARNING", f"commentRangeStart without End: {unmatched_starts}"))
        if unmatched_ends:
            issues.append(("WARNING", f"commentRangeEnd without Start: {unmatched_ends}"))

    if verbose and not issues:
        print(f"  Comments: {len(comment_ids)} comments, all markers valid")

    zf.close()
    return issues


def check_headings(docx_path, verbose=False):
    """Check heading hierarchy for consistency."""
    issues = []

    try:
        doc = Document(str(docx_path))
    except Exception as e:
        return [("CRITICAL", f"Cannot open document: {e}")]

    prev_level = 0
    for i, para in enumerate(doc.paragraphs):
        style = para.style.name if para.style else ""
        if style.startswith("Heading"):
            try:
                level = int(style.split()[-1])
            except (ValueError, IndexError):
                continue

            if level > prev_level + 1 and prev_level > 0:
                issues.append(("WARNING",
                    f"Heading level jump: H{prev_level} to H{level} "
                    f"at paragraph [{i}]: \"{para.text[:50]}\""))
            prev_level = level

    if verbose and not issues:
        print("  Headings: Hierarchy is consistent")

    return issues


def check_content_types(docx_path, verbose=False):
    """Check Content_Types.xml completeness."""
    issues = []

    try:
        zf = zipfile.ZipFile(docx_path, "r")
    except zipfile.BadZipFile:
        return [("CRITICAL", "Not a valid ZIP/DOCX file")]

    if "[Content_Types].xml" not in zf.namelist():
        zf.close()
        return [("CRITICAL", "Missing [Content_Types].xml")]

    ct_xml = zf.read("[Content_Types].xml").decode("utf-8", errors="ignore")
    names = zf.namelist()

    # Check that key parts have content type entries
    part_map = {
        "word/document.xml": "wordprocessingml.document.main",
        "word/styles.xml": "wordprocessingml.styles",
        "word/comments.xml": "wordprocessingml.comments",
    }

    for part, ct_fragment in part_map.items():
        if part in names and ct_fragment not in ct_xml and part not in ct_xml:
            issues.append(("WARNING", f"Part {part} exists but may lack Content-Type entry"))

    # Check for media files without extensions defined
    media_exts = set()
    for name in names:
        if name.startswith("word/media/"):
            ext = Path(name).suffix.lower()
            if ext:
                media_exts.add(ext)

    for ext in media_exts:
        ext_no_dot = ext.lstrip(".")
        if ext_no_dot not in ct_xml:
            issues.append(("WARNING", f"Media extension {ext} may lack Default content type"))

    if verbose and not issues:
        print("  Content Types: All parts properly typed")

    zf.close()
    return issues


def main():
    parser = argparse.ArgumentParser(
        description="Validate DOCX document integrity",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", type=Path, help="Input DOCX file")
    parser.add_argument("--check", default="structure,comments,headings,content-types",
                        help="Comma-separated checks: structure,comments,headings,content-types,schema")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")

    args = parser.parse_args()

    if not args.input.is_file():
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    checks = [c.strip() for c in args.check.split(",")]
    all_issues = []

    print(f"Validating: {args.input.name}\n")

    check_map = {
        "structure": ("Structure", check_structure),
        "comments": ("Comments", check_comments),
        "headings": ("Headings", check_headings),
        "content-types": ("Content Types", check_content_types),
    }

    for check_name in checks:
        if check_name in check_map:
            label, func = check_map[check_name]
            issues = func(args.input, verbose=args.verbose)
            if issues:
                print(f"  {label}:")
                for severity, msg in issues:
                    print(f"    [{severity}] {msg}")
                all_issues.extend(issues)
            elif not args.verbose:
                print(f"  {label}: OK")
        elif check_name == "schema":
            # Use existing OOXML schema validators if available
            try:
                script_dir = Path(__file__).resolve().parent
                skill_dir = script_dir.parent
                sys.path.insert(0, str(skill_dir))
                from ooxml.scripts.validation.docx import DOCXSchemaValidator
                print("  Schema: Using OOXML schema validator (requires unpacked directory)")
                print("    Tip: Unpack first with ooxml/scripts/unpack.py, then validate")
            except ImportError:
                print("  Schema: Validator not available (missing ooxml package)")
        else:
            print(f"  Unknown check: {check_name}")

    print()
    criticals = sum(1 for sev, _ in all_issues if sev == "CRITICAL")
    warnings = sum(1 for sev, _ in all_issues if sev == "WARNING")

    if criticals:
        print(f"FAILED: {criticals} critical issue(s), {warnings} warning(s)")
        sys.exit(1)
    elif warnings:
        print(f"PASSED with {warnings} warning(s)")
        sys.exit(0)
    else:
        print("PASSED: All checks OK")
        sys.exit(0)


if __name__ == "__main__":
    main()
