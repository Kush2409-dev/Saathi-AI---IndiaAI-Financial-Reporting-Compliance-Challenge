"""
ADI Annual Report Extraction Pipeline (v3)
===========================================
1. Run ADI prebuilt-layout on PDF
2. Find TOC → extract section names + printed page numbers
3. Detect page offset (printed page vs PDF page) dynamically
4. Build page ranges using corrected PDF page numbers
5. For each range, dump all paragraphs + tables as clean markdown
6. Output: { "section_name": "clean markdown", ... }

Requires a TOC in the report. Raises error if none found.
"""

import os
import re
import json
import shutil
from pathlib import Path
from collections import defaultdict, Counter
from dotenv import load_dotenv, find_dotenv
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential


# ═══════════════════════════════════════════════════════════════════
# 1. ENV & CONFIG
# ═══════════════════════════════════════════════════════════════════

BASE_PATH = Path(__file__).resolve().parent
ENV_PATH = BASE_PATH / ".env"

if ENV_PATH.exists():
    load_dotenv(dotenv_path=ENV_PATH)
else:
    load_dotenv(find_dotenv())

ADI_ENDPOINT = os.getenv("ADI_ENDPOINT")
ADI_KEY = os.getenv("ADI_KEY")

if not ADI_KEY or not ADI_ENDPOINT:
    raise ValueError(
        f"Environment variables missing.\n"
        f"  Checked: {ENV_PATH}\n"
        f"  ADI_KEY present:      {bool(ADI_KEY)}\n"
        f"  ADI_ENDPOINT present: {bool(ADI_ENDPOINT)}\n"
    )

PROJECT_ROOT = BASE_PATH.parent
IN_DIR = PROJECT_ROOT / "Reports_to_test"
OUT_DIR = PROJECT_ROOT / "Extracted_Reports"
COMPLETED_DIR = IN_DIR / "COMPLETED"

OUT_DIR.mkdir(parents=True, exist_ok=True)
COMPLETED_DIR.mkdir(parents=True, exist_ok=True)

CLIENT = DocumentAnalysisClient(
    endpoint=ADI_ENDPOINT,
    credential=AzureKeyCredential(ADI_KEY),
)


# ═══════════════════════════════════════════════════════════════════
# 2. KNOWN SECTION PATTERNS (for canonicalizing TOC entry names)
# ═══════════════════════════════════════════════════════════════════

KNOWN_SECTIONS = [
    (r"balance\s*sheet",                                    "Balance Sheet"),
    (r"statement\s*of\s*profit\s*and\s*loss",               "Statement of Profit and Loss"),
    (r"profit\s*(?:&|and)\s*loss\s*(?:statement|account)",  "Statement of Profit and Loss"),
    (r"cash\s*flow\s*statement",                            "Cash Flow Statement"),
    (r"statement\s*of\s*changes?\s*in\s*equity",            "Statement of Changes in Equity"),
    (r"notes\s*(?:to|on)\s*(?:the\s*)?financial\s*statements?", "Notes to Financial Statements"),
    (r"significant\s*accounting\s*polic",                   "Significant Accounting Policies"),
    (r"director'?s?\s*report",                              "Directors' Report"),
    (r"management\s*discussion\s*(?:&|and)\s*analysis",     "Management Discussion and Analysis"),
    (r"corporate\s*governance\s*report",                    "Corporate Governance Report"),
    (r"secretarial\s*audit\s*report",                       "Secretarial Audit Report"),
    (r"independent\s*auditor'?s?\s*report",                 "Independent Auditor's Report"),
    (r"auditor'?s?\s*report",                               "Auditor's Report"),
    (r"notice\s*of\s*(?:annual\s*)?general\s*meeting",      "Notice of Annual General Meeting"),
    (r"standalone\s*financial\s*statements?",               "Standalone Financial Statements"),
    (r"consolidated\s*financial\s*statements?",             "Consolidated Financial Statements"),
    (r"mgt[\s\-]*9",                                        "MGT-9"),
]

COMPILED_PATTERNS = [(re.compile(pat, re.IGNORECASE), label) for pat, label in KNOWN_SECTIONS]


def _canonicalize(raw_name: str) -> str:
    """Match a TOC entry to a known canonical name, else return cleaned original."""
    for pattern, label in COMPILED_PATTERNS:
        if pattern.search(raw_name):
            return label
    return re.sub(r"[.\s]+$", "", raw_name).strip()


# ═══════════════════════════════════════════════════════════════════
# 3. PAGE OFFSET DETECTION
# ═══════════════════════════════════════════════════════════════════

def detect_page_offset(paragraphs: list, total_pdf_pages: int) -> int:
    """
    Detect the offset between printed page numbers and PDF page numbers.

    Strategy: Look for standalone numbers near the bottom of pages (page footers).
    These are the printed page numbers. Compare with the PDF page number to
    find the offset.

    Returns: offset such that  pdf_page = printed_page - offset
             (e.g., if printed page 80 is PDF page 76, offset = 4)
    """
    # Collect candidate footer page numbers:
    # A footer is a standalone 1-3 digit number near the bottom of the page (high Y)
    candidates: list[tuple[int, int]] = []  # (pdf_page, printed_number)

    for para in paragraphs:
        text = para.content.strip()
        # Must be a standalone number (1-3 digits), nothing else
        if not re.fullmatch(r"\d{1,3}", text):
            continue

        printed_num = int(text)
        if printed_num < 1 or printed_num > 500:
            continue

        pg = para.bounding_regions[0].page_number
        poly = para.bounding_regions[0].polygon

        if not poly:
            continue

        # Check if near the bottom of the page (Y > 9.0 on a typical ~11-inch page)
        # or near the top (some reports put page numbers at top)
        max_y = max(pt.y for pt in poly)
        min_y = min(pt.y for pt in poly)

        is_footer = max_y > 9.5   # bottom ~15% of page
        is_header = min_y < 1.5   # top ~15% of page

        if is_footer or is_header:
            candidates.append((pg, printed_num))

    if not candidates:
        print("  Page offset: No footer page numbers found → assuming offset = 0")
        return 0

    # Compute offset for each candidate: offset = printed_num - pdf_page
    offsets = [printed - pdf for pdf, printed in candidates]

    # The most common offset wins (robust against occasional misreads)
    offset_counts = Counter(offsets)
    best_offset, count = offset_counts.most_common(1)[0]

    print(f"  Page offset: detected {count}/{len(candidates)} footers agree → offset = {best_offset}")
    print(f"    (printed page = PDF page + {best_offset})")

    # Sanity check: offset should be small and non-negative typically
    if best_offset < -10 or best_offset > 20:
        print(f"  WARNING: Unusual offset {best_offset}, falling back to 0")
        return 0

    return best_offset


# ═══════════════════════════════════════════════════════════════════
# 4. TOC DETECTION
# ═══════════════════════════════════════════════════════════════════

def _name_looks_like_financial_line(name: str) -> bool:
    """Reject names containing amounts like 63,406.03 — these are financial tables."""
    if re.search(r"\d{1,3}(?:,\d{2,3})+\.\d{2}", name):
        return True
    if re.match(r"^[a-z]\)", name.strip()):
        return True
    return False


def _find_toc_in_tables(tables: list, total_pages: int) -> list[dict] | None:
    """
    Find the TOC table. Strict validation:
      - Names must not contain financial amounts
      - Page numbers must span ≥30% of the document
      - Prefer tables with TOC-like header keywords
    """
    TOC_KEYWORDS = {"contents", "index", "table of contents", "particulars",
                    "page no", "page no.", "page number"}
    best = None
    best_score = 0

    for table in tables:
        # Check for TOC header
        header_text = " ".join(
            c.content.strip().lower() for c in table.cells if c.row_index == 0
        )
        has_toc_header = any(kw in header_text for kw in TOC_KEYWORDS)

        cells_by_row = defaultdict(list)
        for cell in table.cells:
            cells_by_row[cell.row_index].append(cell)

        entries = []
        for row_idx in sorted(cells_by_row):
            row_cells = sorted(cells_by_row[row_idx], key=lambda c: c.column_index)
            texts = [c.content.strip() for c in row_cells]

            # Skip header rows
            combined_lower = " ".join(texts).lower()
            skip_kws = ("contents", "page no", "particulars", "s. no", "sr. no", "index")
            if any(kw in combined_lower for kw in skip_kws):
                if not any(re.fullmatch(r"\d{2,3}", t) for t in texts):
                    continue

            # Parse row
            name_parts = []
            page_num = None
            for t in texts:
                if re.fullmatch(r"\d{1,3}", t):
                    page_num = int(t)
                elif re.fullmatch(r"\d+\.", t):
                    continue
                elif t and not re.fullmatch(r"[\d.]+", t):
                    name_parts.append(t)

            name = " ".join(name_parts).strip()
            if _name_looks_like_financial_line(name):
                continue
            if name and page_num and page_num > 0 and len(name) > 2:
                entries.append({"name": name, "start_page": page_num})

        if len(entries) < 4:
            continue

        # Validate: page range should span ≥30% of document
        pages = [e["start_page"] for e in entries]
        coverage = (max(pages) - min(pages)) / max(total_pages, 1)
        if coverage < 0.3:
            continue

        score = len(entries) + (50 if has_toc_header else 0)
        if score > best_score:
            best_score = score
            best = entries

    if best:
        best.sort(key=lambda e: e["start_page"])
        return best
    return None


def _find_toc_in_paragraphs(paragraphs: list, total_pages: int) -> list[dict] | None:
    """Fallback: scan first/last pages for 'Section Name ... 42' patterns."""
    entries = []
    scan_pages = set(range(1, min(6, total_pages + 1)))
    scan_pages |= set(range(max(1, total_pages - 3), total_pages + 1))

    for para in paragraphs:
        pg = para.bounding_regions[0].page_number
        if pg not in scan_pages:
            continue
        text = para.content.strip()
        m = re.match(r"^(.+?)[\s.…·]+(\d{1,3})\s*$", text)
        if m:
            name = m.group(1).strip()
            page = int(m.group(2))
            if len(name) > 3 and page > 0:
                entries.append({"name": name, "start_page": page})

    if len(entries) >= 4:
        seen = set()
        unique = []
        for e in entries:
            key = (e["name"].lower(), e["start_page"])
            if key not in seen:
                seen.add(key)
                unique.append(e)
        unique.sort(key=lambda e: e["start_page"])
        return unique
    return None


def build_page_map(toc: list[dict], total_pdf_pages: int, offset: int) -> tuple[dict[int, str], list[dict]]:
    """
    Build { pdf_page_number: "Section Name" } for every page.

    TOC entries have PRINTED page numbers.
    We convert: pdf_page = printed_page - offset
    """
    page_map: dict[int, str] = {}

    # Convert TOC printed pages → PDF pages
    toc_pdf = []
    for entry in toc:
        pdf_pg = entry["start_page"] - offset
        if pdf_pg < 1:
            pdf_pg = 1  # clamp
        toc_pdf.append({"name": entry["name"], "pdf_page": pdf_pg, "printed_page": entry["start_page"]})

    # Pages before the first TOC section
    first_pdf_page = toc_pdf[0]["pdf_page"]
    for pg in range(1, first_pdf_page):
        page_map[pg] = "Preamble"

    # Assign each TOC section its page range
    for i, entry in enumerate(toc_pdf):
        start = entry["pdf_page"]
        end = toc_pdf[i + 1]["pdf_page"] - 1 if i + 1 < len(toc_pdf) else total_pdf_pages
        label = _canonicalize(entry["name"])
        for pg in range(start, end + 1):
            page_map[pg] = label

    return page_map, toc_pdf


# ═══════════════════════════════════════════════════════════════════
# 5. TABLE → MARKDOWN
# ═══════════════════════════════════════════════════════════════════

def table_to_markdown(table) -> str:
    """Convert an ADI table object to a clean markdown table."""
    grid: dict[tuple[int, int], str] = {}

    for cell in table.cells:
        r, c = cell.row_index, cell.column_index
        text = cell.content.strip().replace("\n", " ").replace("|", "\\|")
        r_span = getattr(cell, "row_span", 1) or 1
        c_span = getattr(cell, "column_span", 1) or 1
        for dr in range(r_span):
            for dc in range(c_span):
                key = (r + dr, c + dc)
                if key not in grid:
                    grid[key] = text if (dr == 0 and dc == 0) else ""

    lines = []
    for r in range(table.row_count):
        row_cells = [grid.get((r, c), "") for c in range(table.column_count)]
        lines.append("| " + " | ".join(row_cells) + " |")
        if r == 0:
            lines.append("| " + " | ".join(["---"] * table.column_count) + " |")

    return "\n".join(lines)


def _table_page(table) -> int:
    if table.bounding_regions:
        return table.bounding_regions[0].page_number
    for cell in table.cells:
        if cell.bounding_regions:
            return cell.bounding_regions[0].page_number
    return 0


def _table_top_y(table) -> float:
    if table.bounding_regions:
        poly = table.bounding_regions[0].polygon
        if poly:
            return min(pt.y for pt in poly)
    return 0.0


# ═══════════════════════════════════════════════════════════════════
# 6. MARKDOWN CLEANUP
# ═══════════════════════════════════════════════════════════════════

def _clean_paragraph_text(text: str) -> str:
    """
    Clean ADI paragraph text for proper markdown rendering:
      - Convert literal \\n to actual newlines
      - Normalize whitespace
      - Strip trailing spaces
    """
    # ADI sometimes returns literal \n in content
    cleaned = text.replace("\\n", "\n")
    # Collapse multiple spaces (but not newlines)
    cleaned = re.sub(r"[^\S\n]+", " ", cleaned)
    # Strip each line
    cleaned = "\n".join(line.strip() for line in cleaned.split("\n"))
    return cleaned.strip()


def _is_page_footer_number(text: str, para) -> bool:
    """Check if a paragraph is just a standalone page number at the bottom of the page."""
    if not re.fullmatch(r"\d{1,3}", text.strip()):
        return False
    poly = para.bounding_regions[0].polygon if para.bounding_regions else None
    if poly:
        max_y = max(pt.y for pt in poly)
        if max_y > 9.5:
            return True
    return False


# ═══════════════════════════════════════════════════════════════════
# 7. CORE EXTRACTION
# ═══════════════════════════════════════════════════════════════════

def extract_report(result) -> dict[str, str]:
    paragraphs = list(result.paragraphs) if result.paragraphs else []
    tables = list(result.tables) if result.tables else []
    total_pdf_pages = result.pages[-1].page_number if result.pages else 1

    # ── Step 1: Detect page offset ──
    offset = detect_page_offset(paragraphs, total_pdf_pages)

    # ── Step 2: Find TOC ──
    # Note: TOC page numbers need offset adjustment for the coverage check,
    # but the raw TOC entries still hold printed page numbers
    toc = _find_toc_in_tables(tables, total_pdf_pages) or \
          _find_toc_in_paragraphs(paragraphs, total_pdf_pages)

    if not toc:
        print(f"  WARNING: No valid TOC found.")
        print(f"  Tables scanned: {len(tables)}, Pages: {total_pdf_pages}")
        raise ValueError(
            "NO TABLE OF CONTENTS FOUND.\n"
            "This pipeline requires a TOC to map sections to page ranges."
        )

    # ── Step 3: Build page map with offset correction ──
    page_map, toc_pdf = build_page_map(toc, total_pdf_pages, offset)

    # Diagnostics
    print(f"  TOC: {len(toc)} sections detected (offset = {offset})")
    for entry in toc_pdf:
        label = _canonicalize(entry["name"])
        # Find end page
        idx = toc_pdf.index(entry)
        end_pg = toc_pdf[idx + 1]["pdf_page"] - 1 if idx + 1 < len(toc_pdf) else total_pdf_pages
        print(f"    printed pg {entry['printed_page']:>3} → PDF pg {entry['pdf_page']:>3}-{end_pg:<3} : {label}")

    # ── Step 4: Build table bounding boxes (to de-duplicate para vs table) ──
    # ADI returns table content BOTH as table cells AND as paragraphs.
    # We keep only the table version and skip overlapping paragraphs.
    table_zones: list[tuple[int, float, float]] = []  # (page, top_y, bottom_y)
    for table in tables:
        pg = _table_page(table)
        top_y = _table_top_y(table)
        # Get bottom Y of table
        bot_y = top_y  # fallback
        if table.bounding_regions:
            poly = table.bounding_regions[0].polygon
            if poly:
                bot_y = max(pt.y for pt in poly)
        # Add a small margin (0.15 inches) to catch paragraphs just outside the box
        table_zones.append((pg, top_y - 0.15, bot_y + 0.15))

    def _para_overlaps_table(pg: int, y: float) -> bool:
        """Check if a paragraph's position falls within any table's bounding box."""
        for t_pg, t_top, t_bot in table_zones:
            if pg == t_pg and t_top <= y <= t_bot:
                return True
        return False

    # ── Step 5: Build content blocks in reading order ──
    blocks = []

    skipped_footer = 0
    skipped_table_overlap = 0

    for para in paragraphs:
        pg = para.bounding_regions[0].page_number
        y = min(pt.y for pt in para.bounding_regions[0].polygon) if para.bounding_regions[0].polygon else 0.0
        text = para.content.strip()

        # Skip page footer numbers
        if _is_page_footer_number(text, para):
            skipped_footer += 1
            continue

        # Skip paragraphs that overlap with a table (already captured in table markdown)
        if _para_overlaps_table(pg, y):
            skipped_table_overlap += 1
            continue

        blocks.append({
            "type": "para",
            "page": pg,
            "y": y,
            "text": text,
        })

    print(f"  Paragraphs: {len(paragraphs)} total → {len(blocks)} kept "
          f"({skipped_table_overlap} table-overlaps removed, {skipped_footer} footers removed)")

    for table in tables:
        pg = _table_page(table)
        y = _table_top_y(table)
        blocks.append({
            "type": "table",
            "page": pg,
            "y": y,
            "md": table_to_markdown(table),
        })

    blocks.sort(key=lambda b: (b["page"], b["y"]))

    # ── Step 6: Identify TOC table pages to skip ──
    toc_table_pages = set()
    for table in tables:
        cells_text = " ".join(c.content for c in table.cells).lower()
        if any(kw in cells_text for kw in ("contents", "table of contents", "index")):
            if "page" in cells_text:
                toc_table_pages.add(_table_page(table))

    # ── Step 7: Assign blocks to sections ──
    sections: dict[str, list[str]] = {}

    for block in blocks:
        pg = block["page"]
        if pg in toc_table_pages:
            continue

        section = page_map.get(pg, "Other")

        if section not in sections:
            sections[section] = []

        if block["type"] == "para":
            cleaned = _clean_paragraph_text(block["text"])
            if cleaned:
                sections[section].append(cleaned)
        else:
            sections[section].append(f"\n{block['md']}\n")

    # ── Step 8: Join into clean markdown ──
    output: dict[str, str] = {}
    for name, chunks in sections.items():
        md = "\n\n".join(chunks).strip()
        if md:
            output[name] = md

    return output


# ═══════════════════════════════════════════════════════════════════
# 8. OUTPUT: MARKDOWN FILES + JSON
# ═══════════════════════════════════════════════════════════════════

def save_outputs(sections: dict[str, str], result, pdf_name: str, out_dir: Path):
    """
    Saves THREE outputs per report:

    1. <name>.md          — Single combined markdown file (feed this to the LLM)
                            Each section starts with ## heading, content follows.

    2. <name>_sections/   — Individual .md file per section (for targeted LLM queries)
         Balance Sheet.md
         Statement of Profit and Loss.md
         ...

    3. <name>.json        — Structured JSON with meta + sections (for programmatic access)
    """
    stem = Path(pdf_name).stem
    total_pages = result.pages[-1].page_number if result.pages else 0
    num_tables = len(result.tables) if result.tables else 0

    # ── 1. Combined markdown ──
    combined_lines = [
        f"# {stem}",
        f"",
        f"**Pages:** {total_pages} | **Tables detected:** {num_tables} | **Sections:** {len(sections)}",
        f"",
        f"---",
        f"",
    ]
    for sec_name, content in sections.items():
        combined_lines.append(f"## {sec_name}")
        combined_lines.append("")
        combined_lines.append(content)
        combined_lines.append("")
        combined_lines.append("---")
        combined_lines.append("")

    combined_md = "\n".join(combined_lines)
    md_path = out_dir / f"{stem}.md"
    md_path.write_text(combined_md, encoding="utf-8")
    print(f"  Saved: {md_path}")

    # ── 2. Individual section .md files ──
    sections_dir = out_dir / f"{stem}_sections"
    sections_dir.mkdir(parents=True, exist_ok=True)

    for sec_name, content in sections.items():
        # Sanitize filename
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', sec_name)[:80]
        sec_path = sections_dir / f"{safe_name}.md"
        sec_md = f"## {sec_name}\n\n{content}\n"
        sec_path.write_text(sec_md, encoding="utf-8")

    print(f"  Saved: {sections_dir}/ ({len(sections)} files)")

    # ── 3. JSON (for programmatic access) ──
    json_output = {
        "meta": {
            "source_file": pdf_name,
            "total_pages": total_pages,
            "sections_found": list(sections.keys()),
            "num_tables_detected": num_tables,
        },
        "sections": sections,
    }
    json_path = out_dir / f"{stem}.json"
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump(json_output, jf, indent=2, ensure_ascii=False)
    print(f"  Saved: {json_path}")

    return json_output


# ═══════════════════════════════════════════════════════════════════
# 9. PIPELINE RUNNER
# ═══════════════════════════════════════════════════════════════════

def run_pipeline():
    pdf_files = list(IN_DIR.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDFs found in {IN_DIR}")
        return

    print(f"Found {len(pdf_files)} PDF(s) to process.\n")

    for pdf_file in pdf_files:
        print(f"{'═' * 60}")
        print(f"Processing: {pdf_file.name}")
        print(f"{'═' * 60}")

        with open(pdf_file, "rb") as f:
            poller = CLIENT.begin_analyze_document("prebuilt-layout", document=f)
            result = poller.result()

        try:
            sections = extract_report(result)
        except ValueError as e:
            print(f"  ERROR: {e}")
            continue

        output = save_outputs(sections, result, pdf_file.name, OUT_DIR)

        print(f"\n  Final sections: {output['meta']['sections_found']}")
        print(f"  Tables detected: {output['meta']['num_tables_detected']}")

        shutil.move(str(pdf_file), str(COMPLETED_DIR / pdf_file.name))
        print()

    print("All done.")


if __name__ == "__main__":
    run_pipeline()