"""
NFRA Audit Compliance — LangGraph Orchestration Engine
========================================================
Multi-step agentic reasoning over extracted annual report sections
to evaluate compliance against accounting standards and audit rules.

Folder structure expected:
  Extracted_Reports/
    CompanyA_sections/
      Balance Sheet.md
      Statement of Profit and Loss.md
      Independent Auditor's Report.md
      Notes to Financial Statements.md
      ...
    CompanyB_sections/
      ...

Usage:
  python audit_orchestration.py \
      --reports  "Extracted_Reports" \
      --rules    "Accounting Standard - Rules.xlsx" \
      --output   "Compliance Matrix.xlsx"
"""

import os
import re
import json
import shutil
import argparse
from typing import TypedDict, List, Dict, Optional, Annotated
from pathlib import Path
from collections import OrderedDict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from dotenv import load_dotenv
from langgraph.graph import StateGraph, END
from langchain_openai import AzureChatOpenAI
from langchain_core.messages import SystemMessage

load_dotenv()  # searches current dir, then parent dirs automatically

# Also try explicit paths common in this project structure
for candidate in [Path(".env"), Path("Python/.env"), Path(__file__).resolve().parent / ".env"]:
    if candidate.exists():
        load_dotenv(dotenv_path=candidate, override=True)
        break

# ── Validate env vars loaded ──
_key = os.getenv("AZURE_OPENAI_API_KEY")
_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
if not _key or not _endpoint:
    raise ValueError(
        f"Azure OpenAI credentials not found.\n"
        f"  AZURE_OPENAI_API_KEY set:  {bool(_key)}\n"
        f"  AZURE_OPENAI_ENDPOINT set: {bool(_endpoint)}\n"
        f"  Looked for .env in: {Path('.env').resolve()}, {Path('Python/.env').resolve()}\n"
        f"  Tip: Place .env in the directory you run the script from."
    )

LLM = AzureChatOpenAI(
    azure_endpoint=_endpoint,
    api_key=_key,
    api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-12-01-preview"),
    azure_deployment=os.getenv("AZURE_OPENAI_DEPLOYMENT"),
    temperature=0,
)


# ═══════════════════════════════════════════════════════════════════
# 2. STATE
# ═══════════════════════════════════════════════════════════════════

class AuditState(TypedDict):
    # ── Identity ──
    company_name: str
    rule_id: str
    rule_details: str
    standard_ref: str
    audit_requirement: str

    # ── Workflow steps parsed from rules Excel ──
    # OrderedDict keyed by step ID: {"1": {step_id, question, target_sections, ...}, "2.1": {...}, ...}
    workflow_steps: Dict[str, Dict]
    step_order: List[str]          # ["1", "2.1", "2.2", "2.3", "3", "4", "4.1", "5", "6"]
    current_step_id: str           # e.g. "1", "2.2", "3" — NOT an integer index
    steps_executed: int            # safety counter to prevent infinite loops
    skipped_steps: List[str]       # step IDs that were skipped by conditional routing

    # ── Report content (loaded from section .md files) ──
    available_sections: Dict[str, str]
    section_filenames: Dict[str, str]

    # ── Accumulated results ──
    question_responses: List[Dict]

    # ── Failure logic ──
    failure_trigger_text: str
    raw_workflow: str
    failure_trigger_met: bool

    # ── Final output (→ Compliance Matrix columns) ──
    compliance_status: Optional[str]
    summary_finding: Optional[str]
    auditor_oversight: Optional[str]
    reasoning_path: Optional[str]
    evidence_snippet: Optional[str]
    page_ref: Optional[str]
    confidence_score: Optional[int]


# ═══════════════════════════════════════════════════════════════════
# 3. REPORT LOADER — reads section .md files from a company folder
# ═══════════════════════════════════════════════════════════════════

def load_company_sections(folder_path: Path) -> Dict[str, str]:
    """
    Load all .md files from a company's extracted sections folder.
    Returns: {"Section Name": "markdown content", ...}
    """
    sections = {}
    for md_file in sorted(folder_path.glob("*.md")):
        section_name = md_file.stem  # filename without .md
        content = md_file.read_text(encoding="utf-8").strip()
        if content:
            sections[section_name] = content
    return sections


def discover_companies(reports_dir: Path) -> List[Dict]:
    """
    Discover all company section folders under Extracted_Reports/.
    Skips COMPLETED folder. Any subfolder containing .md files is a company.
    """
    companies = []
    for item in sorted(reports_dir.iterdir()):
        if not item.is_dir():
            continue
        # Skip COMPLETED and any hidden folders
        if item.name.upper() == "COMPLETED" or item.name.startswith("."):
            continue
        md_files = list(item.glob("*.md"))
        if md_files:
            name = item.name
            if name.endswith("_sections"):
                name = name[: -len("_sections")]
            name = name.replace("_", " ").strip()
            companies.append({
                "name": name,
                "folder": item,
                "section_count": len(md_files),
            })
    return companies


# ═══════════════════════════════════════════════════════════════════
# 4. SECTION MATCHER — fuzzy matches target sections to actual files
# ═══════════════════════════════════════════════════════════════════

# Common aliases in Indian annual reports
SECTION_ALIASES = {
    "independent audit report": ["independent auditor's report", "auditor's report", "auditors report"],
    "independent auditor report": ["independent auditor's report", "auditor's report"],
    "notes to accounts": ["notes to financial statements", "notes on financial statements", "significant accounting policies"],
    "profit and loss account": ["statement of profit and loss", "profit and loss statement", "income statement"],
    "profit and loss": ["statement of profit and loss"],
    "balance sheet": ["balance sheet"],
    "cash flow": ["cash flow statement"],
    "directors report": ["directors' report", "director's report"],
    "corporate governance": ["corporate governance report"],
    "secretarial audit": ["secretarial audit report"],
    "management discussion": ["management discussion and analysis", "mda"],
}

STOP_WORDS = {"the", "of", "and", "to", "a", "in", "for", "on", "at", "as", "by", "an", "its"}


def _normalize(text: str) -> str:
    return re.sub(r"[^a-z0-9\s]", "", text.lower()).strip()


def match_section(target: str, available: Dict[str, str]) -> List[str]:
    """
    Given a target section name from the rules Excel, find the best-matching
    section(s) from the actual extracted report.

    Returns list of matching section names (may return multiple for broad targets).
    """
    target_norm = _normalize(target)
    available_norm = {_normalize(k): k for k in available}

    # 1. Exact match
    if target_norm in available_norm:
        return [available_norm[target_norm]]

    # 2. Alias expansion
    for alias_key, alias_list in SECTION_ALIASES.items():
        if target_norm in _normalize(alias_key) or _normalize(alias_key) in target_norm:
            for alias in alias_list:
                alias_n = _normalize(alias)
                for avail_n, orig_name in available_norm.items():
                    if alias_n in avail_n or avail_n in alias_n:
                        return [orig_name]

    # 3. Substring match
    matches = []
    for avail_n, orig_name in available_norm.items():
        if target_norm in avail_n or avail_n in target_norm:
            matches.append(orig_name)
    if matches:
        return matches

    # 4. Keyword overlap (at least 2 meaningful words)
    target_words = set(target_norm.split()) - STOP_WORDS
    best = []
    best_score = 0
    for avail_n, orig_name in available_norm.items():
        avail_words = set(avail_n.split()) - STOP_WORDS
        overlap = len(target_words & avail_words)
        if overlap > best_score and overlap >= 2:
            best_score = overlap
            best = [orig_name]
        elif overlap == best_score and overlap >= 2:
            best.append(orig_name)

    return best


def retrieve_sections(available: Dict[str, str], targets: List[str], max_chars: int = 30000) -> tuple:
    """
    Fetch content for the requested target sections.
    Includes clear headers and truncation for token management.
    """
    parts = []
    matched = []
    unmatched = []

    for target in targets:
        found = match_section(target, available)
        if found:
            for section_name in found:
                content = available[section_name]
                if len(content) > max_chars:
                    content = content[:max_chars] + "\n\n[... truncated for context window ...]"
                parts.append(f"═══ {section_name} ═══\n\n{content}")
                matched.append(section_name)
        else:
            unmatched.append(target)

    result = "\n\n" + "\n\n".join(parts) if parts else ""

    if unmatched:
        result += f"\n\n⚠ SECTIONS NOT FOUND: {unmatched}"
        result += f"\n  Available sections: {list(available.keys())}"

    return result, matched, unmatched


# ═══════════════════════════════════════════════════════════════════
# 5. RULES PARSER — reads audit rules from Excel
# ═══════════════════════════════════════════════════════════════════

def parse_question_line(block: str) -> Dict:
    """
    Parse workflow step blocks:
      '1. Does the Clause 9...' → step_id="1"
      '2.1 If yes, ...'         → step_id="2.1"
      '4. Has the Company...'   → step_id="4"

    Handles: "N.", "N.N", "N.N " patterns at the start.
    """
    block = block.strip()

    # Try matching "2.1 If yes..." (sub-step with digit after dot)
    m = re.match(r"^(\d+\.\d+)\s+(.*)", block, re.DOTALL)
    if not m:
        # Try matching "1. Does the..." (main step with dot-space separator)
        m = re.match(r"^(\d+)\.\s+(.*)", block, re.DOTALL)
    if not m:
        # Try matching "5 If from..." (step without dot)
        m = re.match(r"^(\d+)\s+(.*)", block, re.DOTALL)
    if not m:
        return None

    step_id = m.group(1)
    rest = m.group(2).strip()

    targets_raw = re.findall(r"\[([^\]]+)\]", rest)
    # Handle comma-separated sections within a single bracket: [P&L, Balance Sheet]
    targets = []
    for t in targets_raw:
        if "," in t:
            targets.extend([s.strip() for s in t.split(",") if s.strip()])
        else:
            targets.append(t.strip())
    question = re.sub(r"\s*:?\s*\[.*", "", rest).rstrip(",").strip()

    return {
        "step_id": step_id,
        "question": question,
        "target_sections": targets,
        "financial_check": None,
        "raw_text": block,
    }


def parse_financial_check(text: str) -> Dict | None:
    """Parse the Numerator/Denominator financial check block."""
    if "numerator" not in text.lower():
        return None
    num = re.search(r"numerator\s*[-–:]\s*(.+?)(?=denominator|\Z)", text, re.I | re.DOTALL)
    den = re.search(r"denominator\s*[-–:]\s*(.+)", text, re.I | re.DOTALL)
    num_secs = re.findall(r"\[([^\]]+)\]", num.group(1)) if num else []
    den_secs = re.findall(r"\[([^\]]+)\]", den.group(1)) if den else []
    return {
        "numerator_desc": num.group(1).strip() if num else "",
        "denominator_desc": den.group(1).strip() if den else "",
        "all_sections": list(set(num_secs + den_secs)),
    }


def parse_rules(excel_path: str, sheet: str = "AS to Validate") -> List[Dict]:
    """Read rules Excel → list of structured rule dicts with ordered workflow steps."""
    df = pd.read_excel(excel_path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    rules = []

    for _, row in df.iterrows():
        rule_id = str(row.get("Rule ID", "")).strip()
        if not rule_id:
            continue

        raw_qs = str(row.get("Audit Questions : Target Sections", ""))
        trigger = str(row.get("Failure Trigger", ""))

        # ── STEP 1: Split raw text into lines ──
        # Excel cells use \n or \r\n for line breaks
        lines = raw_qs.replace("\r\n", "\n").split("\n")

        # ── STEP 2: Merge lines into blocks ──
        # A new block starts when a line begins with a step number like "1.", "2.1", "3."
        # Matches: "1. Does...", "2.1 If...", "3. Check...", "4.1 If..."
        STEP_PATTERN = re.compile(r"^\d+(?:\.\d+)?\s|^\d+\.\s")

        blocks = []
        current_block = ""
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if STEP_PATTERN.match(line):
                if current_block:
                    blocks.append(current_block)
                current_block = line
            elif line.lower().startswith("the financial check"):
                if current_block:
                    blocks.append(current_block)
                current_block = line
            else:
                # Continuation of previous block
                current_block += " " + line
        if current_block:
            blocks.append(current_block)

        # ── STEP 3: Parse each block ──
        workflow_steps = OrderedDict()
        fin_check = None

        for block in blocks:
            block = block.strip()
            if not block:
                continue
            if block.lower().startswith("the financial check"):
                fin_check = parse_financial_check(block)
                continue

            parsed = parse_question_line(block)
            if parsed and parsed["question"]:
                workflow_steps[parsed["step_id"]] = parsed

        # Attach financial check to the relevant step
        if fin_check:
            for sid in reversed(list(workflow_steps.keys())):
                step = workflow_steps[sid]
                if "financial check" in step["question"].lower() or "%" in step["question"]:
                    step["financial_check"] = fin_check
                    step["target_sections"] = list(set(step["target_sections"] + fin_check["all_sections"]))
                    break

        step_order = list(workflow_steps.keys())

        rules.append({
            "rule_id": rule_id,
            "rule_details": str(row.get("Rule Details", "")).strip(),
            "standard_ref": str(row.get("Standard Ref", "")).strip(),
            "audit_requirement": str(row.get("Audit Requirement", "")).strip(),
            "workflow_steps": workflow_steps,
            "step_order": step_order,
            "failure_trigger_text": trigger,
            "raw_workflow": raw_qs,
        })

    return rules


# ═══════════════════════════════════════════════════════════════════
# 6. LLM UTILITY
# ═══════════════════════════════════════════════════════════════════

def llm_json(prompt: str, fallback: Dict) -> Dict:
    """Invoke LLM, extract JSON. Returns fallback on any failure."""
    try:
        resp = LLM.invoke([SystemMessage(content=prompt)])
        text = resp.content.strip()
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
        return json.loads(text)
    except Exception as e:
        print(f"      ⚠ LLM/JSON error: {e}")
        return fallback


# ═══════════════════════════════════════════════════════════════════
# 7. LANGGRAPH NODES
# ═══════════════════════════════════════════════════════════════════

# ─── NODE 1: Initialize ──────────────────────────────────────────

def node_init(state: AuditState) -> AuditState:
    state["current_step_id"] = state["step_order"][0] if state["step_order"] else "DONE"
    state["steps_executed"] = 0
    state["skipped_steps"] = []
    state["question_responses"] = []
    state["failure_trigger_met"] = False
    state["compliance_status"] = None
    state["summary_finding"] = None
    state["auditor_oversight"] = None
    state["confidence_score"] = None
    return state


# ─── NODE 2: Process Question ────────────────────────────────────

def node_process_question(state: AuditState) -> AuditState:
    """
    Core reasoning node. Executes the CURRENT STEP of the audit workflow.
    Uses step_id to look up the step, retrieves sections, includes prior context.
    """
    sid = state["current_step_id"]
    step = state["workflow_steps"][sid]
    total = len(state["step_order"])
    executed = state["steps_executed"] + 1
    state["steps_executed"] = executed

    print(f"    Step {sid} ({executed}/{total}): {step['question'][:80]}...")

    # ── Retrieve target sections ──
    section_content, matched, unmatched = retrieve_sections(
        state["available_sections"], step["target_sections"]
    )

    # ── Cumulative context from prior steps ──
    prior_context = ""
    if state["question_responses"]:
        prior_context = "\n\n─── PRIOR FINDINGS ───\n"
        for r in state["question_responses"]:
            prior_context += (
                f"\nStep {r['step_id']}: {r['question']}\n"
                f"  Answer: {r['answer']}\n"
                f"  Evidence: {r['evidence'][:300]}\n"
            )

    # ── Financial check instructions ──
    fin_block = ""
    if step.get("financial_check"):
        fc = step["financial_check"]
        fin_block = f"""
═══ THIS STEP REQUIRES A FINANCIAL CALCULATION ═══
Numerator: {fc['numerator_desc']}
Denominator: {fc['denominator_desc']}

Procedure:
  1. Extract the relevant finance cost / interest expense figure
  2. Extract borrowings (opening + closing balances)
  3. Average borrowings = (Opening + Closing) / 2
  4. Ratio = (Finance Cost / Average Borrowings) × 100
  5. If ratio < 6%, answer YES (suspiciously low)
  6. Show all numbers and the calculation in your reasoning
═══════════════════════════════════════════════════
"""

    # ── Build full workflow map for context ──
    workflow_map = ""
    skipped = state.get("skipped_steps", [])
    for s_id in state["step_order"]:
        s = state["workflow_steps"][s_id]
        if s_id == sid:
            marker = "→ CURRENT"
        elif any(r["step_id"] == s_id for r in state["question_responses"]):
            marker = "  ✓ done"
        elif s_id in skipped:
            marker = "  ✗ skipped"
        else:
            marker = ""
        workflow_map += f"  [{s_id}] {s['question'][:80]}  {marker}\n"

    # ── Extract search keywords from the question ──
    # Guide the LLM to look for specific terms in large documents
    question_lower = step["question"].lower()
    keyword_hints = set()
    # Add terms directly from the question
    for term in ["npa", "non-performing", "non performing", "default", "interest",
                 "borrowing", "loan", "repayment", "accounted", "not accounted",
                 "provision", "accrual", "derecogn", "impair", "classified",
                 "caro", "clause 9", "clause ix", "finance cost", "amortized",
                 "effective interest", "section 186"]:
        if term in question_lower:
            keyword_hints.add(term)
    # Always add terms from the rule context
    rule_lower = state["rule_details"].lower()
    for term in ["npa", "non-performing", "interest", "default", "loan"]:
        if term in rule_lower:
            keyword_hints.add(term)

    keyword_block = ""
    if keyword_hints:
        keyword_block = f"""
SEARCH GUIDANCE — Look for these keywords/phrases in the report content:
  {', '.join(sorted(keyword_hints))}
  Also search for synonyms and related terms (e.g., "NPA" = "Non-Performing Asset",
  "default" = "overdue", "classified as NPA", "not charged interest", etc.)
  Scan the ENTIRE section — relevant information may be buried in the middle of the document,
  not just in headers or the first few paragraphs.
"""

    prompt = f"""You are an expert Indian financial auditor executing a structured audit workflow.
You must follow the workflow EXACTLY — including conditional branches (if yes → ..., if no → ...).

COMPANY: {state['company_name']}
RULE: {state['rule_id']} — {state['rule_details']}
STANDARD: {state['standard_ref']}
REQUIREMENT: {state['audit_requirement']}

══════════════════════════════════════════════
FULL WORKFLOW MAP (you are at step {sid}):
══════════════════════════════════════════════
{workflow_map}

FAILURE TRIGGER:
{state['failure_trigger_text']}
{prior_context}
══════════════════════════════════════════════
CURRENT STEP [{sid}]:
{step['raw_text']}
══════════════════════════════════════════════
{fin_block}
{keyword_block}
REPORT CONTENT (sections: {matched}):
{section_content}
══════════════════════════════════════════════

INSTRUCTIONS:
1. Execute ONLY step [{sid}]. Do not jump ahead.
2. If this step is conditional (e.g., "If yes...", "If no, go to..."):
   - Check PRIOR FINDINGS to determine which branch applies
   - Apply the condition based on actual evidence found in prior steps
3. SEARCH THOROUGHLY:
   - Read the ENTIRE section content from start to end, not just the beginning
   - Look for the keywords listed above and any related terms
   - Information may appear in notes, sub-notes, disclosure paragraphs, or table footnotes
   - If the section is long, scan for the specific keywords before concluding "Not Found"
4. EVIDENCE DISCIPLINE:
   - Quote EXACT text from the report — copy the relevant sentence(s) verbatim
   - Name the specific section and context (e.g., "Note 12 - Borrowings" or "Para 3 of CARO")
   - "Not Found" means you searched thoroughly and the information genuinely does not exist

CONFIDENCE RUBRIC (apply strictly):
  95-100: Direct explicit statement found
  80-94:  Strong evidence, minor interpretation needed
  60-79:  Indirect/partial evidence
  40-59:  Weak evidence, significant uncertainty
  20-39:  Very little relevant info
  0-19:   Nothing relevant found
  DO NOT default to any fixed number.

Respond ONLY with valid JSON:
{{
    "answer": "Yes / No / Not Found / descriptive answer",
    "evidence": "EXACT quote from report (copy verbatim, or 'None found')",
    "section_ref": "Section name where evidence was found",
    "reasoning": "Step-by-step explanation referencing workflow logic and prior findings",
    "confidence": <integer 0-100>
}}"""

    result = llm_json(prompt, {
        "answer": "Error — unable to process",
        "evidence": "", "section_ref": "", "reasoning": "LLM failed", "confidence": 0,
    })

    result["step_id"] = sid
    result["question"] = step["question"]
    result["target_sections"] = step["target_sections"]
    result["matched_sections"] = matched
    result["unmatched_sections"] = unmatched
    state["question_responses"].append(result)

    print(f"      → {result['answer'][:70]}  [confidence: {result.get('confidence', '?')}]")
    return state


# ─── NODE 3: Validate ────────────────────────────────────────────

def node_validate(state: AuditState) -> AuditState:
    """Senior auditor review — checks evidence quality and recalibrates confidence."""
    r = state["question_responses"][-1]  # always the last response

    prompt = f"""You are a SENIOR AUDITOR reviewing a junior analyst's work. Be critical and honest.

Step [{r['step_id']}]: {r['question']}
Answer: {r['answer']}
Evidence: {r['evidence']}
Reasoning: {r['reasoning']}
Confidence claimed: {r.get('confidence', 'N/A')}

EVALUATE RIGOROUSLY:
1. Is the evidence an exact quote or vague? Exact → strong. Vague → LOWER confidence.
2. Does the reasoning logically hold?
3. "Not found" with high confidence is wrong unless absence IS the finding.

CONFIDENCE RECALIBRATION:
  95-100: Exact quote directly answering the question
  80-94:  Good evidence, minor interpretation
  60-79:  Partial evidence, inference needed
  40-59:  Weak evidence
  20-39:  Barely any info
  0-19:   Nothing found

Respond ONLY with valid JSON:
{{
    "is_valid": true or false,
    "issues": "Specific problems (or 'None')",
    "adjusted_confidence": <integer 0-100>
}}"""

    v = llm_json(prompt, {"is_valid": True, "issues": "Validation skipped", "adjusted_confidence": r.get("confidence", 50)})
    state["question_responses"][-1]["validation"] = v
    if v.get("adjusted_confidence") is not None:
        state["question_responses"][-1]["confidence"] = v["adjusted_confidence"]
    return state


# ─── NODE 4: Route Next Step (LLM-powered workflow router) ───────

def node_route_next_step(state: AuditState) -> AuditState:
    """
    LLM-powered workflow router. Determines next step based on:
      - Conditional logic in the workflow text
      - The last step's answer
      - Which steps are done/skipped
    When jumping (e.g., 1→3), marks intermediate steps as SKIPPED permanently.
    """
    last = state["question_responses"][-1]
    current_sid = state["current_step_id"]
    executed_ids = [r["step_id"] for r in state["question_responses"]]
    skipped = state.get("skipped_steps", [])

    # Build status display — DONE / SKIPPED / PENDING
    step_status = ""
    for sid in state["step_order"]:
        s = state["workflow_steps"][sid]
        if sid in executed_ids:
            ans = next((r["answer"] for r in state["question_responses"] if r["step_id"] == sid), "?")
            step_status += f"  [{sid}] ✓ DONE — Answer: {ans[:60]}\n"
        elif sid in skipped:
            step_status += f"  [{sid}] ✗ SKIPPED (conditional branch not taken)\n"
        else:
            step_status += f"  [{sid}] ○ PENDING — {s['question'][:60]}\n"

    remaining = [s for s in state["step_order"] if s not in executed_ids and s not in skipped]

    prompt = f"""You are a workflow controller. Your ONLY job is to pick the NEXT step ID.

WORKFLOW STATUS:
{step_status}

LAST EXECUTED: Step [{current_sid}]
LAST ANSWER: {last['answer']}

REMAINING STEPS (pick from these ONLY): {remaining}

ROUTING RULES:
1. Read the original workflow text below for conditional logic.
2. "If no, go to point 3" means SKIP to step "3" (intermediate steps are not executed).
3. "If yes, ..." means continue to the next sub-step in order.
4. No conditional → pick the next step in the remaining list.
5. NEVER pick a DONE or SKIPPED step.
6. If no remaining steps, return "DONE".

ORIGINAL WORKFLOW:
{state['raw_workflow']}

Respond ONLY with valid JSON:
{{
    "next_step_id": "<step ID from remaining list, or 'DONE'>",
    "routing_reason": "Brief explanation"
}}"""

    result = llm_json(prompt, {"next_step_id": "DONE", "routing_reason": "Routing failed"})
    next_id = result.get("next_step_id", "DONE").strip()

    # Validate: must be in remaining or DONE
    if next_id != "DONE" and next_id not in remaining:
        # Fallback: pick first remaining
        next_id = remaining[0] if remaining else "DONE"

    # ── Mark intermediate steps as SKIPPED ──
    if next_id != "DONE":
        current_pos = state["step_order"].index(current_sid) if current_sid in state["step_order"] else -1
        next_pos = state["step_order"].index(next_id) if next_id in state["step_order"] else -1
        if next_pos > current_pos + 1:
            for i in range(current_pos + 1, next_pos):
                skip_id = state["step_order"][i]
                if skip_id not in executed_ids and skip_id not in skipped:
                    skipped.append(skip_id)
                    print(f"      ✗ Skipping step [{skip_id}] (conditional branch not taken)")
    state["skipped_steps"] = skipped

    state["current_step_id"] = next_id
    reason = result.get("routing_reason", "")
    print(f"      ↳ Router: [{current_sid}] → [{next_id}]  ({reason[:60]})")
    return state


# ─── NODE 5: Evaluate Failure Triggers ───────────────────────────

def node_evaluate_triggers(state: AuditState) -> AuditState:
    """
    After all questions are answered, apply the failure trigger logic
    from the rules Excel to determine final compliance status.
    """
    print(f"  ⚖ Evaluating failure triggers...")

    qa_block = ""
    executed_ids = [r["step_id"] for r in state["question_responses"]]
    skipped = state.get("skipped_steps", [])
    for sid in state["step_order"]:
        if sid in skipped:
            qa_block += f"Step [{sid}]: SKIPPED (conditional branch not taken)\n\n"
        elif sid in executed_ids:
            r = next(r for r in state["question_responses"] if r["step_id"] == sid)
            qa_block += (
                f"Step [{sid}]: {r['question']}\n"
                f"Answer: {r['answer']}\n"
                f"Evidence: {r['evidence']}\n"
                f"Section: {r.get('section_ref', 'N/A')}\n"
                f"Confidence: {r.get('confidence', 'N/A')}\n\n"
            )

    prompt = f"""You are a senior auditor making the FINAL compliance determination for an Indian company.
You have followed a structured audit workflow step by step. Now synthesize all findings.

COMPANY: {state['company_name']}
RULE: {state['rule_id']} — {state['rule_details']}
STANDARD: {state['standard_ref']}
REQUIREMENT: {state['audit_requirement']}

══════════════════════════════════════════════
AUDIT WORKFLOW THAT WAS FOLLOWED:
══════════════════════════════════════════════
{state.get('raw_workflow', 'N/A')}

══════════════════════════════════════════════
COMPLETE AUDIT TRAIL (results from each workflow step):
══════════════════════════════════════════════
{qa_block}

══════════════════════════════════════════════
FAILURE TRIGGER LOGIC (apply this EXACTLY as written):
══════════════════════════════════════════════
{state['failure_trigger_text']}
══════════════════════════════════════════════

INSTRUCTIONS:
1. Review the workflow steps and their findings holistically
2. Apply EACH condition in the failure trigger against the evidence gathered:
   - If ANY failure condition is met → "Non-Compliant"
   - If a FLAG FOR CHECKING condition is met → "Partial" (needs manual review)
   - If NO failure conditions are met AND evidence confirms compliance → "Compliant"
   - If evidence is insufficient to determine → "Partial"
3. For Auditor Oversight: did the statutory auditor identify this issue in their report?
   If the auditor missed a material non-compliance that the workflow detected, say "Yes".
4. CONFIDENCE: Be honest. If the evidence trail is strong and consistent → high confidence.
   If findings are mixed or evidence was missing for key steps → lower confidence.
   DO NOT default to 85.

Respond ONLY with valid JSON:
{{
    "compliance_status": "Compliant / Non-Compliant / Partial",
    "failure_trigger_met": true or false,
    "summary_finding": "2-3 sentence finding synthesizing the workflow results",
    "auditor_oversight": "Yes — [what was missed] / No — auditor addressed this adequately",
    "confidence": <integer 0-100, calibrated honestly>
}}"""

    result = llm_json(prompt, {
        "compliance_status": "Partial",
        "failure_trigger_met": False,
        "summary_finding": "Unable to determine — processing error",
        "auditor_oversight": "Unknown",
        "confidence": 50,
    })

    state["failure_trigger_met"] = result.get("failure_trigger_met", False)
    state["compliance_status"] = result.get("compliance_status", "Partial")
    state["summary_finding"] = result.get("summary_finding", "")
    state["auditor_oversight"] = result.get("auditor_oversight", "")
    state["confidence_score"] = result.get("confidence", 50)

    status = state["compliance_status"]
    conf = state["confidence_score"]
    print(f"  → {status} (confidence: {conf})")
    return state


# ─── NODE 6: Assemble Output ─────────────────────────────────────

def node_build_output(state: AuditState) -> AuditState:
    """Assemble all fields for the Compliance Matrix row."""

    # Reasoning path: executed steps in order, with skipped steps noted
    rp = []
    executed_ids = [r["step_id"] for r in state["question_responses"]]
    skipped = state.get("skipped_steps", [])

    for sid in state["step_order"]:
        if sid in skipped:
            rp.append(f"Step [{sid}]: SKIPPED (conditional branch not taken)")
        elif sid in executed_ids:
            r = next(r for r in state["question_responses"] if r["step_id"] == sid)
            rp.append(
                f"Step [{sid}]: {r['question']}\n"
                f"  Answer: {r['answer']}\n"
                f"  Reasoning: {r.get('reasoning', 'N/A')}"
            )
    state["reasoning_path"] = "\n\n".join(rp)

    # Evidence: combined from all questions
    ev = []
    for r in state["question_responses"]:
        e = r.get("evidence", "").strip()
        if e and e != "Not found in the provided sections":
            ev.append(e)
    state["evidence_snippet"] = "\n---\n".join(ev) if ev else "No direct evidence found"

    # Section references
    refs = []
    for r in state["question_responses"]:
        ref = r.get("section_ref", "")
        if ref and ref not in refs:
            refs.append(str(ref))
    state["page_ref"] = ", ".join(refs) if refs else "N/A"

    return state


# ═══════════════════════════════════════════════════════════════════
# 8. GRAPH CONSTRUCTION
# ═══════════════════════════════════════════════════════════════════

def route_after_router(state: AuditState) -> str:
    """After the router decides the next step, check if we're done or should continue."""
    if state["current_step_id"] == "DONE":
        return "evaluate_triggers"
    # Safety: prevent infinite loops
    max_steps = len(state["step_order"]) + 2  # allow slight overhead but not infinite
    if state["steps_executed"] >= max_steps:
        print(f"      ⚠ Safety limit reached ({state['steps_executed']} steps), forcing evaluation")
        return "evaluate_triggers"
    return "process_question"


def build_graph():
    """
    Compile the LangGraph workflow with LLM-powered routing:

      init ──► process_question ──► validate ──► route_next_step ──┐
                    ▲                                               │
                    └──── (next step?) ◄────────────────────────────┘
                                │ (DONE)
                                ▼
                     evaluate_triggers ──► build_output ──► END
    """
    g = StateGraph(AuditState)

    g.add_node("init", node_init)
    g.add_node("process_question", node_process_question)
    g.add_node("validate", node_validate)
    g.add_node("route_next_step", node_route_next_step)
    g.add_node("evaluate_triggers", node_evaluate_triggers)
    g.add_node("build_output", node_build_output)

    g.set_entry_point("init")
    g.add_edge("init", "process_question")
    g.add_edge("process_question", "validate")
    g.add_edge("validate", "route_next_step")
    g.add_conditional_edges("route_next_step", route_after_router, {
        "process_question": "process_question",
        "evaluate_triggers": "evaluate_triggers",
    })
    g.add_edge("evaluate_triggers", "build_output")
    g.add_edge("build_output", END)

    return g.compile()


# ═══════════════════════════════════════════════════════════════════
# 9. COMPLIANCE MATRIX WRITER
# ═══════════════════════════════════════════════════════════════════

COLUMNS = [
    "Company Name", "Rule ID", "Compliance Status", "Summary Finding",
    "Auditor Oversight", "Reasoning Path", "Evidence Snippet",
    "Page Ref", "Confidence Score",
]

HDR_FILL = PatternFill("solid", fgColor="2F5496")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11, name="Arial")
BODY_FONT = Font(size=10, name="Arial")
WRAP = Alignment(wrap_text=True, vertical="top")
BORDER = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
STATUS_FILL = {
    "Compliant": PatternFill("solid", fgColor="C6EFCE"),
    "Non-Compliant": PatternFill("solid", fgColor="FFC7CE"),
    "Partial": PatternFill("solid", fgColor="FFEB9C"),
}


def state_to_row(s: AuditState) -> Dict:
    return {
        "Company Name": s.get("company_name", ""),
        "Rule ID": s.get("rule_id", ""),
        "Compliance Status": s.get("compliance_status", "Partial"),
        "Summary Finding": s.get("summary_finding", ""),
        "Auditor Oversight": s.get("auditor_oversight", ""),
        "Reasoning Path": s.get("reasoning_path", ""),
        "Evidence Snippet": s.get("evidence_snippet", ""),
        "Page Ref": s.get("page_ref", ""),
        "Confidence Score": s.get("confidence_score", 0),
    }


def write_matrix(rows: List[Dict], path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Compliance Matrix"
    widths = [25, 14, 18, 50, 35, 60, 50, 22, 15]

    for ci, (col, w) in enumerate(zip(COLUMNS, widths), 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[c.column_letter].width = w

    for ri, row in enumerate(rows, 2):
        for ci, col in enumerate(COLUMNS, 1):
            val = row.get(col, "")
            c = ws.cell(row=ri, column=ci, value=val)
            c.font, c.alignment, c.border = BODY_FONT, WRAP, BORDER
            if col == "Compliance Status" and str(val).strip() in STATUS_FILL:
                c.fill = STATUS_FILL[str(val).strip()]
                c.font = Font(bold=True, size=10, name="Arial")
            if col == "Confidence Score":
                c.alignment = Alignment(horizontal="center", vertical="top")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)
    print(f"\n✅ Compliance Matrix saved → {path}")


# ═══════════════════════════════════════════════════════════════════
# 10. MAIN PIPELINE
# ═══════════════════════════════════════════════════════════════════

def run(reports_dir: str, rules_path: str, output_path: str, rules_sheet: str = "AS to Validate"):
    reports_root = Path(reports_dir)
    completed_dir = reports_root / "COMPLETED"
    completed_dir.mkdir(parents=True, exist_ok=True)

    companies = discover_companies(reports_root)

    if not companies:
        print(f"No company folders found in {reports_root}")
        print(f"  (Skipping 'COMPLETED' folder)")
        print(f"  Expected: subfolders containing .md section files")
        return

    print(f"Found {len(companies)} company folder(s):")
    for c in companies:
        print(f"  • {c['name']} ({c['section_count']} sections) → {c['folder'].name}/")

    rules = parse_rules(rules_path, rules_sheet)
    print(f"\nLoaded {len(rules)} rule(s) from '{rules_path}'")
    for r in rules:
        print(f"  • {r['rule_id']}: {r['rule_details'][:60]}... ({len(r['step_order'])} steps)")

    workflow = build_graph()
    all_rows = []

    for company in companies:
        sections = load_company_sections(company["folder"])
        print(f"\n{'━' * 70}")
        print(f"COMPANY: {company['name']}")
        print(f"  Sections: {list(sections.keys())}")
        print(f"{'━' * 70}")

        for rule in rules:
            if not rule["step_order"]:
                print(f"  ⏭ {rule['rule_id']}: no workflow steps parsed, skipping")
                continue

            print(f"\n  ── Rule {rule['rule_id']}: {rule['rule_details'][:50]}... ──")
            print(f"     Steps: {rule['step_order']}")

            initial: AuditState = {
                "company_name": company["name"],
                "rule_id": rule["rule_id"],
                "rule_details": rule["rule_details"],
                "standard_ref": rule["standard_ref"],
                "audit_requirement": rule["audit_requirement"],
                "workflow_steps": rule["workflow_steps"],
                "step_order": rule["step_order"],
                "current_step_id": rule["step_order"][0],
                "steps_executed": 0,
                "skipped_steps": [],
                "available_sections": sections,
                "section_filenames": {k: f"{k}.md" for k in sections},
                "question_responses": [],
                "failure_trigger_text": rule["failure_trigger_text"],
                "raw_workflow": rule.get("raw_workflow", ""),
                "failure_trigger_met": False,
                "compliance_status": None,
                "summary_finding": None,
                "auditor_oversight": None,
                "reasoning_path": None,
                "evidence_snippet": None,
                "page_ref": None,
                "confidence_score": None,
            }

            final = workflow.invoke(initial)
            all_rows.append(state_to_row(final))

        # ── Move completed company folder to COMPLETED ──
        dest = completed_dir / company["folder"].name
        if dest.exists():
            shutil.rmtree(dest)  # overwrite if re-running
        shutil.move(str(company["folder"]), str(dest))
        print(f"\n  📁 Moved → COMPLETED/{company['folder'].name}")

    write_matrix(all_rows, output_path)

    debug = Path(output_path).with_suffix(".debug.json")
    with open(debug, "w", encoding="utf-8") as f:
        json.dump(all_rows, f, indent=2, ensure_ascii=False)
    print(f"  Debug JSON → {debug}")

    print(f"\n{'═' * 70}")
    print(f"DONE — {len(companies)} companies × {len(rules)} rules = {len(all_rows)} evaluations")
    print(f"{'═' * 70}")


# ═══════════════════════════════════════════════════════════════════
# 11. CLI — Smart defaults, no required arguments
# ═══════════════════════════════════════════════════════════════════

def _auto_detect_paths() -> Dict[str, str]:
    """
    Auto-detect project paths relative to this script's location.
    Expected structure:
      project_root/
        Python/
          audit_orchestration.py   ← this script
        Extracted_Reports/
          CompanyA_sections/
          COMPLETED/
        Accounting Standard - Rules.xlsx
        Compliance Matrix.xlsx
    """
    script_dir = Path(__file__).resolve().parent        # Python/
    project_root = script_dir.parent                     # project_root/

    # Try multiple common locations for each path
    reports_candidates = [
        project_root / "Extracted_Reports",
        script_dir / "Extracted_Reports",
        Path("Extracted_Reports"),
    ]
    rules_candidates = [
        project_root / "Accounting Standard - Rules.xlsx",
        script_dir / "Accounting Standard - Rules.xlsx",
        Path("Accounting Standard - Rules.xlsx"),
    ]

    reports = next((p for p in reports_candidates if p.exists()), reports_candidates[0])
    rules = next((p for p in rules_candidates if p.exists()), rules_candidates[0])
    output = project_root / "Compliance Matrix.xlsx"

    return {
        "reports": str(reports),
        "rules": str(rules),
        "output": str(output),
    }


if __name__ == "__main__":
    defaults = _auto_detect_paths()

    p = argparse.ArgumentParser(
        description="NFRA Audit Compliance — LangGraph Orchestration",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "If no arguments are provided, paths are auto-detected from the project structure.\n"
            f"  --reports  → {defaults['reports']}\n"
            f"  --rules    → {defaults['rules']}\n"
            f"  --output   → {defaults['output']}"
        ),
    )
    p.add_argument("--reports", default=defaults["reports"], help="Extracted_Reports folder")
    p.add_argument("--rules", default=defaults["rules"], help="Accounting Standard Rules Excel")
    p.add_argument("--output", default=defaults["output"], help="Output Compliance Matrix Excel")
    p.add_argument("--sheet", default="AS to Validate", help="Sheet name in rules Excel")
    args = p.parse_args()

    print(f"{'═' * 70}")
    print(f"NFRA Audit Compliance — LangGraph Orchestration Engine")
    print(f"{'═' * 70}")
    print(f"  Reports:  {args.reports}")
    print(f"  Rules:    {args.rules}")
    print(f"  Output:   {args.output}")
    print(f"  Sheet:    {args.sheet}")
    print(f"{'═' * 70}\n")

    run(args.reports, args.rules, args.output, args.sheet)