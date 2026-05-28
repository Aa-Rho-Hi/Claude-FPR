import streamlit as st
import tempfile, os, glob, re, sys, io, json
from datetime import date

st.set_page_config(
    page_title="FAR Extraction Pipeline",
    page_icon="📄",
    layout="wide",
)

try:
    import run_rules as rr
    _import_ok = True
except Exception as _e:
    _import_ok = False
    _import_err = str(_e)

if not _import_ok:
    st.error(f"❌ Failed to load run_rules.py: {_import_err}")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# Helpers — defined before tabs so all sections can call them
# ══════════════════════════════════════════════════════════════════════════════

RULE_TYPES = {
    "Count entries containing keyword":
        "Count each numbered entry (1. 2. 3. …) in a section that contains a specific word.\n"
        "Example: count all patents → section: Patents, keyword: patent",
    "Count all entries in section":
        "Count every numbered entry in a section, regardless of content.\n"
        "Example: count all book chapters → section: Book Chapters",
    "Count entries from a specific year":
        "Count numbered entries in a section that mention a particular year.\n"
        "Example: count 2024 invited talks → section: Invited Talks, year: 2024",
    "Count entries matching ALL keywords":
        "Count entries that contain every keyword listed (AND logic).\n"
        "Example: must contain both 'IEEE' and 'transactions' → keywords: IEEE, transactions",
    "Count entries matching ANY keyword":
        "Count entries that contain at least one of the keywords listed (OR logic).\n"
        "Example: count entries with 'award' or 'prize' → keywords: award, prize",
    "Count entries NOT containing keyword":
        "Count entries in a section that do NOT contain a keyword (exclusion).\n"
        "Example: conference papers excluding workshops → keyword: workshop",
}

def _run_custom_rule(cv_text, cr):
    """Run a custom rule against cv_text and return an integer count."""
    section    = cr.get("section", "").strip().lower()
    rule_type  = cr.get("rule_type", "Count entries containing keyword")
    keywords   = [k.strip().lower() for k in cr.get("keywords", "").split(",") if k.strip()]
    year_str   = str(cr.get("year", "")).strip()

    # Extract the relevant section text
    lines      = cv_text.split("\n")
    in_sec     = not section
    sec_lines  = []
    for line in lines:
        stripped = line.strip()
        ll = stripped.lower()
        if section and section in ll:
            in_sec = True; continue
        if in_sec and section and re.match(r'^[A-Z][A-Z\s]{5,}$', stripped) and stripped.lower() != section:
            in_sec = False
        if in_sec:
            sec_lines.append(stripped)

    sec_text = "\n".join(sec_lines)

    # Split into numbered entries
    entry_pat = re.compile(r'(?m)^\d+[\.\)]\s')
    splits    = [m.start() for m in entry_pat.finditer(sec_text)]
    if splits:
        entries = [sec_text[splits[i]: splits[i+1] if i+1 < len(splits) else len(sec_text)]
                   for i in range(len(splits))]
    else:
        # No numbered entries — fall back to non-blank lines
        entries = [l for l in sec_lines if l]

    count = 0
    for entry in entries:
        el = entry.lower()
        if rule_type == "Count all entries in section":
            count += 1
        elif rule_type == "Count entries containing keyword":
            if keywords and keywords[0] in el:
                count += 1
        elif rule_type == "Count entries from a specific year":
            if year_str and year_str in el:
                count += 1
        elif rule_type == "Count entries matching ALL keywords":
            if keywords and all(k in el for k in keywords):
                count += 1
        elif rule_type == "Count entries matching ANY keyword":
            if keywords and any(k in el for k in keywords):
                count += 1
        elif rule_type == "Count entries NOT containing keyword":
            if keywords and keywords[0] not in el:
                count += 1
    return count


def _parse_custom_rules(text):
    """
    Parse plain-English rules from the Rules Editor text area.
    Looks for blocks of:
        Rule name:  ...
        Look in:    ...   (optional)
        Count:      ...
    Only parses rules that appear AFTER the 'ADD YOUR RULES HERE' marker.
    """
    rules = []
    # Only look at text after the user-rules marker
    marker = "ADD YOUR RULES HERE"
    idx = text.upper().find(marker.upper())
    if idx >= 0:
        text = text[idx + len(marker):]

    # Split into candidate rule blocks by finding "Rule name:" anchors
    blocks = re.split(r'(?im)^\s*rule\s+name\s*:', text)
    for block in blocks[1:]:  # first element is text before first rule
        lines = [l.strip() for l in block.strip().split("\n") if l.strip()]
        if not lines:
            continue

        name    = lines[0].strip()
        section = ""
        count   = ""

        for line in lines[1:]:
            ll = line.lower()
            if ll.startswith("look in:"):
                section = line.split(":", 1)[1].strip()
            elif ll.startswith("count:"):
                count = line.split(":", 1)[1].strip()

        if not name or not count:
            continue

        # Interpret the count instruction
        cl = count.lower()
        if cl.startswith("all entries") or cl == "all":
            rule_type = "Count all entries in section"
            keywords  = ""
            year      = ""
        elif cl.startswith("year:"):
            rule_type = "Count entries from a specific year"
            keywords  = ""
            year      = count.split(":", 1)[1].strip()
        elif cl.startswith("any of:"):
            rule_type = "Count entries matching ANY keyword"
            keywords  = count.split(":", 1)[1].strip()
            year      = ""
        elif cl.startswith("all of:"):
            rule_type = "Count entries matching ALL keywords"
            keywords  = count.split(":", 1)[1].strip()
            year      = ""
        elif cl.startswith("excludes:") or cl.startswith("exclude:"):
            rule_type = "Count entries NOT containing keyword"
            keywords  = count.split(":", 1)[1].strip()
            year      = ""
        elif cl.startswith("contains:"):
            rule_type = "Count entries containing keyword"
            keywords  = count.split(":", 1)[1].strip()
            year      = ""
        else:
            # Fallback: treat the whole count value as a keyword
            rule_type = "Count entries containing keyword"
            keywords  = count
            year      = ""

        rules.append({
            "name":      name,
            "section":   section,
            "rule_type": rule_type,
            "keywords":  keywords,
            "year":      year,
        })
    return rules


# ══════════════════════════════════════════════════════════════════════════════
# AI-interpreted extraction (Option B)
# ══════════════════════════════════════════════════════════════════════════════

_AI_SYSTEM_PROMPT = """\
You are a precise data-extraction assistant for faculty annual reports.
You will be given extraction rules and raw text from a faculty member's documents.
Follow the rules exactly and return ONLY a valid JSON object — no prose, no markdown fences.

Required JSON keys (use null if genuinely not found):
  last_name    : string
  first_name   : string
  title        : string  (Professor | Associate Professor | Assistant Professor)
  ug           : integer (undergrad course count)
  grad         : integer (grad course count)
  ms           : integer (MS/MSEN graduated, chair only)
  phd          : integer (PhD graduated, chair only)
  grants       : integer (funded grants where faculty is PI or CoPI)
  ch_co        : integer (current doctoral advisees: Chair + Co-Chair)
  cp           : integer (conference proceeding papers)
  journal      : integer (refereed journal papers)

If the rules define CUSTOM fields, include those keys too with integer values.
Never hallucinate counts. When uncertain, use 0 for counts and null for text fields.
"""

def extract_with_ai(
    rules_text: str,
    last_name: str,
    far_text: str,
    cv_text: str,
    xlsx_summary: str,
    custom_rules: list,
    api_key: str,
    base_url: str | None = None,
    model: str = "gpt-4o",
) -> dict | None:
    """
    Send rules + document text to an OpenAI-compatible LLM and parse the JSON result.
    Uses requests directly so it works with any OpenAI-compatible endpoint (including TAMU).
    Returns a result dict matching run_rules.extract_faculty() output, or None on failure.
    """
    import requests as _requests

    # Build the user message
    custom_rule_block = ""
    if custom_rules:
        lines = ["\n\nCUSTOM FIELDS (also include these as integer keys in your JSON):"]
        for cr in custom_rules:
            sec = f"in section '{cr['section']}'" if cr["section"] else "across entire document"
            lines.append(f"  - key: \"{cr['name']}\" — {cr['rule_type']} {sec}"
                         + (f", keywords: {cr['keywords']}" if cr.get("keywords") else "")
                         + (f", year: {cr['year']}" if cr.get("year") else ""))
        custom_rule_block = "\n".join(lines)

    user_msg = f"""\
=== EXTRACTION RULES ===
{rules_text}
{custom_rule_block}

=== FACULTY: {last_name} ===

--- FAR PDF TEXT ---
{far_text[:12000]}

--- CV PDF TEXT ---
{cv_text[:6000]}

--- SUPPLEMENTAL WORKBOOK DATA ---
{xlsx_summary[:3000] if xlsx_summary else "(not provided)"}

Now extract the data and return ONLY the JSON object."""

    # Determine endpoint — append /chat/completions to whatever base URL is given
    if base_url and base_url.strip():
        endpoint = base_url.rstrip("/") + "/chat/completions"
    else:
        endpoint = "https://api.openai.com/v1/chat/completions"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": _AI_SYSTEM_PROMPT},
            {"role": "user",   "content": user_msg},
        ],
        "temperature": 0,
        "max_tokens": 1200,
        "stream": False,          # explicitly disable streaming
    }

    def _parse_sse_content(text: str) -> str:
        """
        Reassemble content from a Server-Sent Events stream.
        TAMU sends metadata chunks + [DONE] early, then the real content chunks follow.
        So we do NOT stop at [DONE] — we collect delta.content from every chunk.
        """
        content_parts = []
        for line in text.splitlines():
            line = line.strip()
            if not line.startswith("data:"):
                continue
            payload_str = line[5:].strip()
            if payload_str == "[DONE]":
                continue   # skip but don't stop — real content may follow
            try:
                chunk = json.loads(payload_str)
                delta = (chunk.get("choices") or [{}])[0].get("delta") or {}
                part = delta.get("content") or ""
                if part:
                    content_parts.append(part)
            except Exception:
                continue
        return "".join(content_parts)

    try:
        resp = _requests.post(endpoint, json=payload, headers=headers, timeout=90)

        if not resp.ok:
            st.warning(f"⚠️ {last_name}: AI API returned {resp.status_code} — {resp.text[:200]}. Falling back.")
            return None

        # Try standard JSON parse first; fall back to SSE reassembly
        content = ""
        try:
            resp_json = resp.json()
            choices = resp_json.get("choices") or []
            if choices:
                message = choices[0].get("message") or {}
                content = message.get("content") or ""
        except Exception:
            pass   # not JSON — likely SSE stream

        # If content still empty, attempt SSE reassembly
        if not content and "data:" in resp.text:
            content = _parse_sse_content(resp.text)

        if not content:
            st.warning(f"⚠️ {last_name}: AI returned empty content. Falling back.")
            return None

        raw = content.strip()

        # Strip markdown fences if the model added them despite instructions
        raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.I)
        raw = re.sub(r'\s*```$', '', raw)
        data = json.loads(raw)

        # Normalise to match run_rules result dict format
        result = {
            "last_name":  str(data.get("last_name",  last_name)),
            "first_name": str(data.get("first_name", "")),
            "title":      str(data.get("title",      "")),
            "ug":         int(data.get("ug",      0) or 0),
            "grad":       int(data.get("grad",    0) or 0),
            "ms":         int(data.get("ms",      0) or 0),
            "phd":        int(data.get("phd",     0) or 0),
            "grants":     int(data.get("grants",  0) or 0),
            "ch_co":      int(data.get("ch_co",   0) or 0),
            "cp":         int(data.get("cp",      0) or 0),
            "journal":    int(data.get("journal", 0) or 0),
        }
        # Custom fields
        for cr in custom_rules:
            result[cr["name"]] = int(data.get(cr["name"], 0) or 0)

        return result

    except json.JSONDecodeError as e:
        st.warning(f"⚠️ {last_name}: AI returned invalid JSON — {e}. Falling back to rule-based extraction.")
        return None
    except Exception as e:
        st.warning(f"⚠️ {last_name}: AI extraction failed — {e}. Falling back to rule-based extraction.")
        return None


def _xlsx_to_summary(xlsx_path: str) -> str:
    """Convert an XLSX file to a plain-text summary for the AI prompt."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        lines = []
        for sheet_name in wb.sheetnames[:3]:  # max 3 sheets
            ws = wb[sheet_name]
            lines.append(f"[Sheet: {sheet_name}]")
            for row in ws.iter_rows(max_row=60, values_only=True):
                row_str = "\t".join(str(c) if c is not None else "" for c in row)
                if row_str.strip():
                    lines.append(row_str)
        return "\n".join(lines)
    except Exception:
        return ""


# ── Default rule editor text ───────────────────────────────────────────────────
DEFAULT_RULES_TEXT = """\
Faculty Annual Report Extraction Rules
=======================================
Use these rules exactly and conservatively.
Standard rules always run. Scroll to the bottom to add your own custom rules.


Processing Workflow
--------------------
1. Upload all FAR PDFs, CV PDFs, and supplemental XLSX files together.
2. Select which faculty to process (all detected, or specific names).
3. Click Run Extraction. Each faculty member is processed in order.
4. Download the output Excel workbook when processing is complete.
5. The Excel workbook contains 3 sheets: Summary, Teaching & Advising,
   and Research & Publications.


Global Instructions
--------------------
1. Supplemental data is authoritative when present for fields where these
   rules allow supplemental values.
2. When multiple sources disagree, follow the metric-specific priority rules
   defined below instead of strictly preferring the PDF.
3. If a row is incomplete, do not count it unless the relevant rule explicitly
   allows it.
4. When a table column appears wrapped or shifted onto a nearby line, reconstruct
   the row conservatively before deciding that a field is missing.
5. In Graduate Advising, inspect the far-right Role column carefully — Chair or
   Co-Chair may appear detached from the rest of the row.
6. If evidence is ambiguous, exclude it from counts.
7. Keep all count fields as integers. Use 0 for missing numeric fields only when
   absence is confirmed. Use empty strings for missing text fields.
8. Extraction failure is not evidence of absence. A blank field caused by a
   known PDF extraction issue is different from a field that is genuinely
   missing. Attempt repair before treating a blank as absent.


Companion CV Usage
-------------------
1. The companion CV is supplemental only.
2. Use the CV only for metrics whose field rules explicitly allow CV evidence.
3. For Courses, Faculty Names, and Titles, do not add or override values
   from the CV.
4. For Publications, Grants, and Graduate Advising counts, the CV may be used
   as additional evidence when the field rule says so. De-duplicate matching
   entries by normalised title, name, and year before counting.
5. Use only dedicated CV sections:
   - Publications: Peer-Reviewed Journal Articles, Refereed Conference
     Proceedings, or equivalent complete publication lists.
   - Grants: Research Funding, Extramural Funding, Grants, Sponsored Projects,
     or equivalent complete funding lists.
   - Students: Current Students, Graduate Student Committee Chair, Theses,
     Dissertations, or equivalent advising lists.


Supplemental Workbook Usage
-----------------------------
1. Titles: always use supplemental workbook values where allowed.
2. CH/CO: always use supplemental workbook value. Do NOT reconstruct from
   Graduate Advising.
3. Publications (CP Totals, Refereed Journal Papers):
   Use max(annual-report + allowed CV count, supplemental workbook count).
4. Grants: always use supplemental workbook values for External and Internal
   grants when present.
5. MS/MSEN and PhD graduates: use annual-report row-level counting. Use a
   supplemental value only when the annual-report table is demonstrably
   incomplete (missing roles, truncated, or missing rows).
6. Courses: use annual report only.


Field Rules
------------
1. Faculty Last Name, First Name, Title
   Extract from the top section of the FAR PDF.
   Normalise academic ranks: output "Professor", "Associate Professor", or
   "Assistant Professor" only — remove endowed chair names, center roles,
   and administrative titles.

2. UG Courses
   Source: Teaching section of the FAR PDF.
   Count distinct course numbers below 500 taught during the report year.
   Exclude course numbers 485, 491, 681, 684, 685, and 691.
   Exclude independent study, research, thesis, dissertation, seminar,
   internship, and special-problem courses.
   Treat a lecture and its companion lab/design-lab/honors row as one course.

3. Grad Courses
   Source: Teaching section of the FAR PDF.
   Count distinct course numbers 500 or above taught during the report year.
   Apply the same exclusions as UG Courses.

4. PhD Graduated
   Source: Graduate Advising table, annual report only.
   Count rows where Degree = PhD, Role = Chair, and Graduation Date is in
   the report year.
   Do not count Member, Co-Chair, or blank-role rows.

5. MS/MSEN Graduated
   Source: Graduate Advising table; supplemental workbook only if the table
   is demonstrably incomplete.
   Count rows where Degree = MS or MSEN, Role = Chair, and Graduation Date
   is in the report year.
   Confirm against companion CV advising section when available.

6. Total Grants
   Source: supplemental workbook when present; otherwise annual report and
   dedicated CV funding sections.
   Count grants where faculty is PI or Co-PI, status is Funded/Active/Awarded,
   and start date is in the report year.
   Gifts count as grants when they have an identified sponsor and a 2024 date.
   Exclude proposals that are only In Preparation, Submitted, or Pending.

7. CH/CO
   Source: always use supplemental workbook value.
   If supplemental is missing, count current doctoral advisees from Graduate
   Advising (Role = Chair or Co-Chair, End Semester = Ongoing) plus CV
   current-student sections.

8. CP Totals
   Source: max(annual-report conference entries, CV conference list,
   supplemental workbook count).
   Count conference proceeding entries whose venue year is the report year
   and faculty appears in the author list.
   Include accepted/to-appear entries at the top of a reverse-chronological
   CV list.

9. Refereed Journal Papers
   Source: max(annual-report journal entries, CV journal list,
   supplemental workbook count).
   Count peer-reviewed journal entries whose publication year is the report
   year and faculty appears in the author list.
   Do not count arXiv preprints.
   Do not count conference papers misclassified under journal sections.


CUSTOM RULES  (add your own below this line)
=============================================
To add a new column to the output Excel, write a rule using this format:

   Rule name:  [column label in the output]
   Look in:    [section heading in the CV — leave blank to search entire document]
   Count:      [choose one option below]

Count options:
   all entries              — every numbered item in the section
   contains: word           — items that include this word or phrase
   year: 2024               — items that mention a specific year
   any of: word1, word2     — items with at least one of these words (OR)
   all of: word1, word2     — items with every one of these words (AND)
   excludes: word           — items that do NOT contain this word

Examples (remove the leading # to activate):
   # Rule name:  Invited Talks
   # Look in:    Invited Talks
   # Count:      all entries

   # Rule name:  Book Chapters
   # Look in:    Book Chapters
   # Count:      all entries

   # Rule name:  Patents
   # Look in:    Patents
   # Count:      all entries

   # Rule name:  Awards
   # Look in:    Honors and Awards
   # Count:      contains: award

   # Rule name:  2024 Conference Talks
   # Look in:    Invited Talks
   # Count:      year: 2024

ADD YOUR RULES HERE
====================
(write your custom rules below — follow the format above)

"""

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    output_filename = st.text_input("Output filename", value="far_extraction_output.xlsx")

    st.markdown("---")
    st.markdown("#### 🤖 AI Extraction (Option B)")
    st.caption(
        "Provide an API key to let an AI model read the Rules Editor text and "
        "extract values directly. Leave blank to use the built-in rule-based engine."
    )
    ai_api_key = st.text_input(
        "AI API Key",
        type="password",
        placeholder="sk-… or your TAMU key",
        help="OpenAI-compatible API key. When provided, the AI reads your rules and performs extraction.",
    )
    ai_base_url = st.text_input(
        "API Base URL (optional)",
        placeholder="https://api.openai.com/v1",
        help="Leave blank for OpenAI. For TAMU or other providers, paste their endpoint URL here.",
    )
    ai_model = st.text_input(
        "Model name",
        value="protected.gpt-5",
        help="Model to use for AI extraction. For TAMU: protected.gpt-5. For OpenAI: gpt-4o.",
    )
    if ai_api_key:
        st.success("🟢 AI extraction enabled")
    else:
        st.info("🔵 Rule-based extraction (no API key)")

    st.markdown("---")
    st.markdown("**Expected file naming:**")
    st.code(
        "F180Vita_F.Lastname.pdf   ← FAR\n"
        "Lastname CV.pdf           ← CV\n"
        "Lastname.xlsx             ← Supplemental\n"
        "support_staff.xlsx        ← Staff",
        language=None,
    )

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_upload, tab_config, tab_rules = st.tabs(["📂 Upload & Run", "⚙️ Configuration", "📝 Rules Editor"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Upload & Run
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.subheader("1. Upload Files")
    uploaded_files = st.file_uploader(
        "Upload all files (FAR PDFs, CV PDFs, XLSX files) — select all at once",
        accept_multiple_files=True,
        type=["pdf", "xlsx", "xls"],
    )
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
        with st.expander("Show uploaded files"):
            for f in uploaded_files:
                st.text(f"{f.name}  ({f.size/1024:.1f} KB)")

    st.subheader("2. Select Faculty")
    run_mode = st.radio(
        "Process",
        ["All detected faculty", "Specific faculty only"],
        horizontal=True,
    )
    specific_names = ""
    if run_mode == "Specific faculty only":
        specific_names = st.text_input(
            "Last names (comma-separated)",
            placeholder="e.g.  Narayanan, Palermo, Qian",
        )

    st.subheader("3. Run Extraction")
    run_btn = st.button("▶ Run Extraction", type="primary", disabled=not uploaded_files)

    if run_btn and uploaded_files:
        # ── Apply UI configuration overrides ──────────────────────────────
        cfg = st.session_state.get("cfg", {})

        report_year = cfg.get("report_year", rr.REPORT_YEAR)
        rr.REPORT_YEAR            = report_year
        rr.Q4_START               = date(report_year, cfg.get("q4_month", 10), 1)
        rr.UG_COURSE_CEILING      = cfg.get("ug_ceiling", rr.UG_COURSE_CEILING)
        rr.GRANT_COUNTED_ROLES    = set(cfg.get("grant_roles", list(rr.GRANT_COUNTED_ROLES)))
        rr.GRANT_STATUS_KEYWORD   = cfg.get("grant_status_kw", rr.GRANT_STATUS_KEYWORD)
        rr.GRANT_PROGRESS_KEYWORD = cfg.get("grant_progress_kw", rr.GRANT_PROGRESS_KEYWORD)
        rr.GRANT_MIN_END_DATE     = date(report_year, cfg.get("grant_min_month", 10), 1)
        if cfg.get("shell_tokens"):
            rr.RESEARCH_SHELL_TOKENS = set(t.strip().upper() for t in cfg["shell_tokens"].split(",") if t.strip())
        if cfg.get("journal_hdrs"):
            rr.JOURNAL_HDR_KW = set(t.strip().lower() for t in cfg["journal_hdrs"].split(",") if t.strip())
        if cfg.get("conf_hdrs"):
            rr.CONF_HDR_KW = set(t.strip().lower() for t in cfg["conf_hdrs"].split(",") if t.strip())
        if cfg.get("cv_pattern"):
            rr.FILE_PATTERN_CV = cfg["cv_pattern"]
        if cfg.get("xlsx_pattern"):
            rr.FILE_PATTERN_XLSX = cfg["xlsx_pattern"]
        if cfg.get("staff_sheet"):
            rr.FILE_STAFF_SHEET = cfg["staff_sheet"]

        # ── Load custom rules from the Rules Editor ────────────────────────
        rules_text   = st.session_state.get("rules_text", DEFAULT_RULES_TEXT)
        custom_rules = _parse_custom_rules(rules_text)

        with tempfile.TemporaryDirectory() as tmpdir:
            for uf in uploaded_files:
                with open(os.path.join(tmpdir, uf.name), "wb") as fh:
                    fh.write(uf.getbuffer())

            _known = getattr(rr, 'KNOWN_FACULTY', ['Narayanan','Qian','Palermo','Hu','Duffield'])
            if run_mode == "All detected faculty":
                faculty_list = list(_known)
                for fp in glob.glob(os.path.join(tmpdir, "F180Vita_*.pdf")):
                    m = re.match(r'F180Vita_\w+\.(\w+)\.pdf', os.path.basename(fp))
                    if m and m.group(1) not in faculty_list:
                        faculty_list.append(m.group(1))
                faculty_list = [
                    ln for ln in faculty_list
                    if glob.glob(os.path.join(tmpdir, f"F180Vita_*.{ln}.pdf"))
                ]
            else:
                faculty_list = [n.strip() for n in specific_names.split(",") if n.strip()]

            if not faculty_list:
                st.warning("⚠️ No faculty detected. Check FAR PDFs follow the naming pattern `F180Vita_F.Lastname.pdf`.")
                st.stop()

            st.info(f"Found {len(faculty_list)} faculty: {', '.join(faculty_list)}")

            all_far_data = {}
            with st.spinner("Pre-parsing FAR PDFs…"):
                for far_path in glob.glob(os.path.join(tmpdir, "F180Vita_*.pdf")):
                    m = re.match(r'F180Vita_\w+\.(\w+)\.pdf', os.path.basename(far_path))
                    if not m: continue
                    ln = m.group(1)
                    try:
                        all_far_data[ln] = (rr.parse_far(far_path), rr.pdf_full_text(far_path))
                    except Exception as e:
                        st.warning(f"Could not pre-parse {os.path.basename(far_path)}: {e}")

            st.markdown("---")
            st.subheader("Results")
            if ai_api_key:
                st.caption("🤖 AI extraction is active — rules text is sent to the AI model for interpretation.")
            else:
                st.caption("🔵 Rule-based extraction — using built-in pipeline.")
            progress  = st.progress(0, text="Starting…")
            log_area  = st.empty()
            results   = []
            log_lines = []

            for i, last_name in enumerate(faculty_list):
                progress.progress(i / len(faculty_list), text=f"Processing {last_name}…")
                log_lines.append(f"⏳ **{last_name}**…")
                log_area.markdown("\n\n".join(log_lines))

                r = None

                # ── Try AI extraction if API key provided ──────────────────
                if ai_api_key:
                    far_text_ai = all_far_data.get(last_name, (None, ""))[1]
                    cv_path_ai  = os.path.join(tmpdir, f"{last_name} CV.pdf")
                    cv_text_ai  = ""
                    if os.path.exists(cv_path_ai):
                        try:
                            cv_text_ai = rr.pdf_full_text(cv_path_ai)
                        except Exception:
                            pass
                    xlsx_path_ai = os.path.join(tmpdir, f"{last_name}.xlsx")
                    xlsx_summary_ai = _xlsx_to_summary(xlsx_path_ai) if os.path.exists(xlsx_path_ai) else ""

                    r = extract_with_ai(
                        rules_text      = rules_text,
                        last_name       = last_name,
                        far_text        = far_text_ai,
                        cv_text         = cv_text_ai,
                        xlsx_summary    = xlsx_summary_ai,
                        custom_rules    = custom_rules,
                        api_key         = ai_api_key,
                        base_url        = ai_base_url if ai_base_url else None,
                        model           = ai_model or "gpt-4o",
                    )

                # ── Fallback: rule-based extraction ───────────────────────
                if r is None:
                    old_stdout = sys.stdout
                    sys.stdout = io.StringIO()
                    try:
                        r = rr.extract_faculty(
                            last_name,
                            input_dir=tmpdir,
                            api_key=None,
                            all_far_data=all_far_data,
                        )
                    except Exception as e:
                        r = None
                        log_lines[-1] = f"❌ **{last_name}** — {e}"
                    finally:
                        sys.stdout = old_stdout

                    # Run custom rules on top of rule-based result
                    if r and custom_rules:
                        cv_text = all_far_data.get(last_name, (None, ""))[1]
                        for cr in custom_rules:
                            r[cr["name"]] = _run_custom_rule(cv_text, cr)

                if r:
                    results.append(r)
                    base = (f"✅ **{last_name}** — "
                            f"UG={r['ug']} Grad={r['grad']} MS={r['ms']} PhD={r['phd']} | "
                            f"Grants={r['grants']} CH/CO={r['ch_co']} CP={r['cp']} Journal={r['journal']}")
                    extras = "  ".join(f"{cr['name']}={r.get(cr['name'],0)}" for cr in custom_rules)
                    log_lines[-1] = base + (f"  |  {extras}" if extras else "")
                log_area.markdown("\n\n".join(log_lines))

            progress.progress(1.0, text="Done!")

            if not results:
                st.error("No results produced. Check file naming and try again.")
                st.stop()

            out_path = os.path.join(tmpdir, output_filename)
            rr.generate_excel(results, out_path)
            with open(out_path, "rb") as fh:
                excel_bytes = fh.read()

            st.markdown("---")
            st.subheader("📊 Summary")
            import pandas as pd
            rows = []
            for r in results:
                row = {
                    "Last Name": r["last_name"], "First Name": r["first_name"],
                    "Title": r["title"], "UG": r["ug"], "Grad": r["grad"],
                    "MS": r["ms"], "PhD": r["phd"], "Grants": r["grants"],
                    "CH/CO": r["ch_co"], "CP": r["cp"], "Journal": r["journal"],
                }
                for cr in custom_rules:
                    row[cr["name"]] = r.get(cr["name"], 0)
                rows.append(row)
            st.dataframe(pd.DataFrame(rows).set_index("Last Name"), use_container_width=True)

            st.download_button(
                label=f"⬇️  Download {output_filename}",
                data=excel_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Configuration
# ══════════════════════════════════════════════════════════════════════════════
with tab_config:
    st.subheader("⚙️ Extraction Configuration")
    st.caption("Changes here apply to the next run. No coding needed.")

    cfg = st.session_state.get("cfg", {})

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📅 Report Period")
        report_year = st.number_input("Report Year", min_value=2010, max_value=2100,
                                      value=cfg.get("report_year", rr.REPORT_YEAR), step=1)
        q4_month = st.selectbox("Q4 starts in (month)",
                                options=list(range(1, 13)),
                                index=cfg.get("q4_month", 10) - 1,
                                format_func=lambda m: date(2000, m, 1).strftime("%B"))
        grant_min_month = st.selectbox("Grant must be active through (month)",
                                       options=list(range(1, 13)),
                                       index=cfg.get("grant_min_month", 10) - 1,
                                       format_func=lambda m: date(2000, m, 1).strftime("%B"),
                                       help="Set to January to count any grant active at any point during the year.")
    with col2:
        st.markdown("#### 📚 Course Rules")
        ug_ceiling = st.number_input(
            "Undergrad course number ceiling",
            min_value=100, max_value=900,
            value=cfg.get("ug_ceiling", rr.UG_COURSE_CEILING), step=100,
            help="Courses BELOW this number count as undergrad. Common values: 400 or 500.")
        shell_tokens = st.text_area(
            "Exclude these course title words (comma-separated)",
            value=cfg.get("shell_tokens", ", ".join(sorted(rr.RESEARCH_SHELL_TOKENS))),
            help="Courses whose titles contain any of these words are excluded from UG and Grad counts.",
            height=80,
        )

    st.markdown("#### 💰 Grant Rules")
    col3, col4 = st.columns(2)
    with col3:
        grant_roles = st.multiselect(
            "Count grants where faculty role is",
            options=["PI", "CoPI", "Other"],
            default=cfg.get("grant_roles", sorted(rr.GRANT_COUNTED_ROLES)),
        )
        grant_status_kw = st.text_input("Grant status must contain",
                                        value=cfg.get("grant_status_kw", rr.GRANT_STATUS_KEYWORD))
    with col4:
        grant_progress_kw = st.text_input("Grant status must also contain",
                                          value=cfg.get("grant_progress_kw", rr.GRANT_PROGRESS_KEYWORD))

    st.markdown("#### 📄 Publication Section Headers")
    col5, col6 = st.columns(2)
    with col5:
        journal_hdrs = st.text_area(
            "Journal section headings (comma-separated)",
            value=cfg.get("journal_hdrs", ", ".join(sorted(rr.JOURNAL_HDR_KW))),
            height=90)
    with col6:
        conf_hdrs = st.text_area(
            "Conference section headings (comma-separated)",
            value=cfg.get("conf_hdrs", ", ".join(sorted(rr.CONF_HDR_KW))),
            height=90)

    st.markdown("#### 🗂️ File Naming")
    col7, col8, col9 = st.columns(3)
    with col7:
        cv_pattern   = st.text_input("CV filename pattern",
                                     value=cfg.get("cv_pattern", rr.FILE_PATTERN_CV),
                                     help="Use {last} as placeholder for last name.")
    with col8:
        xlsx_pattern = st.text_input("Supplemental XLSX pattern",
                                     value=cfg.get("xlsx_pattern", rr.FILE_PATTERN_XLSX),
                                     help="Use {last} as placeholder for last name.")
    with col9:
        staff_sheet  = st.text_input("Staff sheet filename",
                                     value=cfg.get("staff_sheet", rr.FILE_STAFF_SHEET))

    if st.button("💾 Save Configuration", type="primary"):
        st.session_state["cfg"] = {
            "report_year": report_year, "q4_month": q4_month,
            "grant_min_month": grant_min_month, "ug_ceiling": ug_ceiling,
            "shell_tokens": shell_tokens, "grant_roles": grant_roles,
            "grant_status_kw": grant_status_kw, "grant_progress_kw": grant_progress_kw,
            "journal_hdrs": journal_hdrs, "conf_hdrs": conf_hdrs,
            "cv_pattern": cv_pattern, "xlsx_pattern": xlsx_pattern,
            "staff_sheet": staff_sheet,
        }
        st.success("✅ Configuration saved — will apply on next run.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Rules Editor
# ══════════════════════════════════════════════════════════════════════════════
with tab_rules:
    st.subheader("📝 Rules")
    st.caption(
        "Read the standard rules below and add your own at the bottom of the text box. "
        "Click **Save Rules** when done — your custom rules will appear as extra columns in the output."
    )

    rules_text = st.text_area(
        label="Rules text",
        value=st.session_state.get("rules_text", DEFAULT_RULES_TEXT),
        height=600,
    )

    col_save, col_reset, col_info = st.columns([1, 1, 4])
    with col_save:
        if st.button("💾 Save Rules", type="primary"):
            st.session_state["rules_text"]  = rules_text
            st.session_state["custom_rules"] = _parse_custom_rules(rules_text)
            n = len(st.session_state["custom_rules"])
            st.success(f"✅ Saved — {n} custom rule{'s' if n != 1 else ''} active.")
    with col_reset:
        if st.button("↩️ Reset to defaults"):
            st.session_state["rules_text"]   = DEFAULT_RULES_TEXT
            st.session_state["custom_rules"] = []
            st.rerun()

    # Live preview of parsed custom rules
    parsed = _parse_custom_rules(rules_text)
    if parsed:
        st.markdown("---")
        st.markdown(f"**{len(parsed)} custom rule(s) detected — will add columns to output:**")
        for cr in parsed:
            sec = f"in section *{cr['section']}*" if cr["section"] else "across entire CV"
            rt  = cr["rule_type"]
            if rt == "Count all entries in section":
                detail = f"count all entries {sec}"
            elif rt == "Count entries from a specific year":
                detail = f"count entries from year **{cr['year']}** {sec}"
            elif rt in ("Count entries matching ALL keywords", "Count entries matching ANY keyword"):
                logic = "ALL of" if "ALL" in rt else "ANY of"
                detail = f"count entries with {logic} `{cr['keywords']}` {sec}"
            elif rt == "Count entries NOT containing keyword":
                detail = f"count entries that do NOT contain `{cr['keywords']}` {sec}"
            else:
                detail = f"count entries containing `{cr['keywords']}` {sec}"
            st.markdown(f"- **{cr['name']}** — {detail}")


