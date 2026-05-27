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

def _parse_custom_rules(text):
    """Parse the Rules Editor text into a list of {name, section, keyword} dicts."""
    rules = []
    STANDARD = {"ug courses","grad courses","ms graduated","phd graduated",
                 "grants","ch/co","cp","journal"}
    for line in text.split("\n"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = [p.strip() for p in line.split("|")]
        if len(parts) != 3:
            continue
        name, section, keyword = parts
        if not name or not keyword:
            continue
        if name.lower() in STANDARD:
            continue  # skip standard-rule description lines
        rules.append({"name": name, "section": section, "keyword": keyword.lower()})
    return rules


def _run_custom_rule(cv_text, cr):
    """Count lines in cv_text that match the custom rule's section + keyword."""
    count      = 0
    section    = cr["section"].lower()
    keyword    = cr["keyword"].lower()
    in_section = not section  # no section filter → scan everything
    for line in cv_text.split("\n"):
        stripped = line.strip()
        ll = stripped.lower()
        if section and section in ll:
            in_section = True
        elif in_section and section and re.match(r'^[A-Z][A-Z\s]{5,}$', stripped):
            in_section = False
        if in_section and keyword in ll:
            count += 1
    return count


# ── Default rule editor text ───────────────────────────────────────────────────
DEFAULT_RULES_TEXT = """\
# ─────────────────────────────────────────────────────────────────────────────
# STANDARD RULES  (always applied — edit descriptions for reference only)
# ─────────────────────────────────────────────────────────────────────────────

UG Courses       | Counts distinct undergrad course numbers (below the UG ceiling) taught this year.
Grad Courses     | Counts distinct graduate course numbers (at or above the UG ceiling) taught this year.
MS Graduated     | Counts MS/MEN students who graduated this year with faculty as chair.
PhD Graduated    | Counts PhD students who graduated this year with faculty as chair.
Grants           | Counts funded-in-progress PI/CoPI grants that started this year and are active through Q4.
CH/CO            | Counts active graduate chair/co-chair advisees (ongoing + current CV members).
CP               | Counts conference papers published this year (best of FAR, CV, and supplemental).
Journal          | Counts refereed journal papers published this year (best of CV and supplemental).

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM RULES  (add your own below — one rule per line)
# Format:  Column Name  |  Section in CV to search  |  Keyword to count
#
# • Column Name   — label that appears in the output Excel
# • Section       — heading in the CV where the pipeline should look
#                   (leave blank to search the entire CV)
# • Keyword       — any line containing this word is counted
#
# Examples (remove the leading # to activate):
# ─────────────────────────────────────────────────────────────────────────────
# Invited Talks   | Invited Talks              | invited
# Book Chapters   | Book Chapters              | chapter
# Patents         | Patents                    | patent
# Awards          | Honors and Awards          | award
# Editorials      | Editorial                  | editor
# Media Coverage  | Media                      | coverage
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

        # ── Parse custom rules from the Rules Editor ───────────────────────
        rules_text = st.session_state.get("rules_text", DEFAULT_RULES_TEXT)
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
    st.subheader("📝 Rules Editor")
    st.caption(
        "The standard rules are shown below and always run. "
        "To add a custom metric, uncomment one of the examples or add your own line "
        "in the format:  **Column Name  |  Section in CV  |  Keyword to count**"
    )

    rules_text = st.text_area(
        label="Rules",
        value=st.session_state.get("rules_text", DEFAULT_RULES_TEXT),
        height=500,
        label_visibility="collapsed",
        help="Lines starting with # are comments. Custom rule lines must have exactly two | separators.",
    )

    col_a, col_b = st.columns([1, 4])
    with col_a:
        if st.button("💾 Save Rules", type="primary"):
            st.session_state["rules_text"] = rules_text
            custom = _parse_custom_rules(rules_text)
            st.success(f"✅ Saved — {len(custom)} custom rule(s) active.")
    with col_b:
        if st.button("↩️ Reset to defaults"):
            st.session_state["rules_text"] = DEFAULT_RULES_TEXT
            st.rerun()

    # Live preview of active custom rules
    custom_preview = _parse_custom_rules(rules_text)
    if custom_preview:
        st.markdown("---")
        st.markdown("**Active custom rules (will add columns to output):**")
        for cr in custom_preview:
            section_label = f"in section *{cr['section']}*" if cr["section"] else "across entire CV"
            st.markdown(f"- **{cr['name']}** — count lines with `{cr['keyword']}` {section_label}")
    else:
        st.info("No active custom rules. Uncomment or add lines below the dashed separator to add one.")


