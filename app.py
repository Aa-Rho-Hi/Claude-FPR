import streamlit as st
import tempfile, os, glob, re, sys, io
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


# ── Default rule editor text ───────────────────────────────────────────────────
DEFAULT_RULES_TEXT = """\
Faculty Annual Report Extraction Rules
=======================================
Use these rules to control what the pipeline counts for each faculty member.
Standard rules always run. Add your own custom rules at the bottom.


STANDARD RULES  (always applied)
----------------------------------
These run automatically and cannot be removed.

1. UG Courses
   Count the number of distinct undergraduate courses (course number below 500)
   taught during the report year.

2. Grad Courses
   Count the number of distinct graduate courses (course number 500 or above)
   taught during the report year.

3. MS Graduated
   Count MS/MEN students who graduated during the report year with this faculty
   member as committee chair.

4. PhD Graduated
   Count PhD students who graduated during the report year with this faculty
   member as committee chair.

5. Grants
   Count funded, in-progress grants where faculty is PI or Co-PI, that started
   during the report year and are still active through October.

6. CH/CO
   Count active graduate students for whom this faculty member is the chair or
   co-chair of the thesis or dissertation committee.

7. CP
   Count conference papers published during the report year (takes the highest
   count across the FAR, CV, and supplemental spreadsheet).

8. Journal
   Count refereed journal papers published during the report year (takes the
   highest count across the CV and supplemental spreadsheet).


CUSTOM RULES  (add your own below)
------------------------------------
To add a new column to the output, write a rule using this format:

   Rule name:  [what to call this column in the output Excel]
   Look in:    [section heading in the CV — leave blank to search the entire document]
   Count:      [what to count — choose one of the options below]

Count options:
   all entries              — count every numbered item in the section
   contains: word           — count items that include this word or phrase
   year: 2024               — count items that mention a specific year
   any of: word1, word2     — count items containing at least one of these words
   all of: word1, word2     — count items that contain every one of these words
   excludes: word           — count items that do NOT contain this word


EXAMPLES  (copy, edit, and add below the line to activate)
------------------------------------------------------------
   Rule name:  Invited Talks
   Look in:    Invited Talks
   Count:      all entries

   Rule name:  Book Chapters
   Look in:    Book Chapters
   Count:      all entries

   Rule name:  Patents
   Look in:    Patents
   Count:      all entries

   Rule name:  Awards
   Look in:    Honors and Awards
   Count:      contains: award

   Rule name:  2024 Talks
   Look in:    Invited Talks
   Count:      year: 2024


ADD YOUR RULES HERE
====================
(write your rules below this line — follow the format shown above)

"""

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    output_filename = st.text_input("Output filename", value="far_extraction_output.xlsx")
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
            progress  = st.progress(0, text="Starting…")
            log_area  = st.empty()
            results   = []
            log_lines = []

            for i, last_name in enumerate(faculty_list):
                progress.progress(i / len(faculty_list), text=f"Processing {last_name}…")
                log_lines.append(f"⏳ **{last_name}**…")
                log_area.markdown("\n\n".join(log_lines))

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

                # Run custom rules
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


