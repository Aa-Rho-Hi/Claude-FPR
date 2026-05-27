import streamlit as st
import tempfile, os, glob, re, sys, io, importlib
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

# ── Title ──────────────────────────────────────────────────────────────────────
st.title("📄 Faculty Annual Report Extraction Pipeline")
st.caption("Upload FAR PDFs, CV PDFs, and supplemental XLSX files to extract structured data into Excel.")

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input(
        "LlamaParse API Key (optional)",
        type="password",
        help="Only needed if the pipeline uses LlamaParse for OCR fallback.",
    )
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
tab_upload, tab_config, tab_rules = st.tabs(["📂 Upload & Run", "⚙️ Configuration", "➕ Custom Rules"])

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
        # ── Apply UI configuration to run_rules module ─────────────────────
        cfg = st.session_state.get("cfg", {})

        report_year = cfg.get("report_year", rr.REPORT_YEAR)
        rr.REPORT_YEAR           = report_year
        rr.Q4_START              = date(report_year, cfg.get("q4_month", 10), 1)
        rr.UG_COURSE_CEILING     = cfg.get("ug_ceiling", rr.UG_COURSE_CEILING)
        rr.GRANT_COUNTED_ROLES   = set(cfg.get("grant_roles", list(rr.GRANT_COUNTED_ROLES)))
        rr.GRANT_STATUS_KEYWORD  = cfg.get("grant_status_kw", rr.GRANT_STATUS_KEYWORD)
        rr.GRANT_PROGRESS_KEYWORD= cfg.get("grant_progress_kw", rr.GRANT_PROGRESS_KEYWORD)
        rr.GRANT_MIN_END_DATE    = date(report_year, cfg.get("grant_min_month", 10), 1)
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

        # Custom rules from tab 3
        custom_rules = st.session_state.get("custom_rules", [])

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
                st.warning("⚠️ No faculty detected. Check that FAR PDFs follow the naming pattern `F180Vita_F.Lastname.pdf`.")
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
            progress = st.progress(0, text="Starting…")
            log_area = st.empty()
            results  = []
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
                        api_key=api_key or None,
                        all_far_data=all_far_data,
                    )
                except Exception as e:
                    r = None
                    log_lines[-1] = f"❌ **{last_name}** — {e}"
                finally:
                    sys.stdout = old_stdout

                # Run custom rules
                if r and custom_rules:
                    cv_text  = all_far_data.get(last_name, (None, ""))[1]
                    far_data = all_far_data.get(last_name, (None, None))[0] or {}
                    for cr in custom_rules:
                        col  = cr["name"]
                        kw   = cr["keyword"].lower()
                        sec  = cr["section"].lower()
                        count = 0
                        in_section = not sec  # if no section filter, scan everything
                        for line in cv_text.split("\n"):
                            ll = line.lower()
                            if sec and sec in ll:
                                in_section = True
                            elif in_section and re.match(r'^[A-Z][A-Z\s]{5,}$', line.strip()):
                                in_section = False
                            if in_section and kw and kw in ll:
                                count += 1
                        r[col] = count

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
            standard_cols = ["Last Name","First Name","Title","UG","Grad","MS","PhD","Grants","CH/CO","CP","Journal"]
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
                                       help="Set to January to count any grant active at any point in the year.")

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
            help="Select which collaborator roles qualify a grant for counting.",
        )
        grant_status_kw = st.text_input(
            "Grant status must contain",
            value=cfg.get("grant_status_kw", rr.GRANT_STATUS_KEYWORD),
            help="e.g. 'funded' — the status field must contain this word.",
        )
    with col4:
        grant_progress_kw = st.text_input(
            "Grant status must also contain",
            value=cfg.get("grant_progress_kw", rr.GRANT_PROGRESS_KEYWORD),
            help="e.g. 'progress' — both keywords must appear in the status.",
        )

    st.markdown("#### 📄 Publication Section Headers")
    col5, col6 = st.columns(2)
    with col5:
        journal_hdrs = st.text_area(
            "Journal section headings (comma-separated)",
            value=cfg.get("journal_hdrs", ", ".join(sorted(rr.JOURNAL_HDR_KW))),
            height=90,
            help="Phrases that mark a journal publication section in FAR/CV PDFs.",
        )
    with col6:
        conf_hdrs = st.text_area(
            "Conference section headings (comma-separated)",
            value=cfg.get("conf_hdrs", ", ".join(sorted(rr.CONF_HDR_KW))),
            height=90,
            help="Phrases that mark a conference publication section in FAR/CV PDFs.",
        )

    st.markdown("#### 🗂️ File Naming")
    col7, col8, col9 = st.columns(3)
    with col7:
        cv_pattern = st.text_input("CV filename pattern",
                                   value=cfg.get("cv_pattern", rr.FILE_PATTERN_CV),
                                   help="Use {last} as placeholder for last name.")
    with col8:
        xlsx_pattern = st.text_input("Supplemental XLSX pattern",
                                     value=cfg.get("xlsx_pattern", rr.FILE_PATTERN_XLSX),
                                     help="Use {last} as placeholder for last name.")
    with col9:
        staff_sheet = st.text_input("Staff sheet filename",
                                    value=cfg.get("staff_sheet", rr.FILE_STAFF_SHEET))

    if st.button("💾 Save Configuration", type="primary"):
        st.session_state["cfg"] = {
            "report_year": report_year,
            "q4_month": q4_month,
            "grant_min_month": grant_min_month,
            "ug_ceiling": ug_ceiling,
            "shell_tokens": shell_tokens,
            "grant_roles": grant_roles,
            "grant_status_kw": grant_status_kw,
            "grant_progress_kw": grant_progress_kw,
            "journal_hdrs": journal_hdrs,
            "conf_hdrs": conf_hdrs,
            "cv_pattern": cv_pattern,
            "xlsx_pattern": xlsx_pattern,
            "staff_sheet": staff_sheet,
        }
        st.success("✅ Configuration saved — will apply on next run.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Custom Rules
# ══════════════════════════════════════════════════════════════════════════════
with tab_rules:
    st.subheader("➕ Custom Metrics")
    st.caption(
        "Add extra columns to the output by telling the pipeline what to count. "
        "Each rule scans the CV PDF for a keyword inside a named section."
    )

    if "custom_rules" not in st.session_state:
        st.session_state["custom_rules"] = []

    # Display existing rules
    rules = st.session_state["custom_rules"]
    if rules:
        st.markdown("**Active custom rules:**")
        to_delete = []
        for idx, cr in enumerate(rules):
            col_a, col_b, col_c, col_d = st.columns([2, 3, 3, 1])
            col_a.markdown(f"**{cr['name']}**")
            col_b.markdown(f"Section: `{cr['section'] or 'entire CV'}`")
            col_c.markdown(f"Keyword: `{cr['keyword']}`")
            if col_d.button("🗑️", key=f"del_{idx}"):
                to_delete.append(idx)
        for idx in reversed(to_delete):
            rules.pop(idx)
        st.session_state["custom_rules"] = rules
        st.markdown("---")

    # Add new rule form
    st.markdown("**Add a new rule:**")
    col1, col2, col3 = st.columns([2, 3, 3])
    with col1:
        new_name = st.text_input("Column name", placeholder="e.g. Invited Talks",
                                 help="This becomes a new column in the output Excel.")
    with col2:
        new_section = st.text_input("Look inside section (optional)",
                                    placeholder="e.g. Invited Talks",
                                    help="Leave blank to search the entire CV. "
                                         "Enter a section heading to restrict the search.")
    with col3:
        new_keyword = st.text_input("Count lines containing",
                                    placeholder="e.g. invited talk",
                                    help="Each line in the section that contains this word is counted.")

    if st.button("➕ Add Rule"):
        if not new_name:
            st.error("Please enter a column name.")
        elif not new_keyword:
            st.error("Please enter a keyword to count.")
        elif any(cr["name"] == new_name for cr in rules):
            st.error(f"A rule named '{new_name}' already exists.")
        else:
            st.session_state["custom_rules"].append({
                "name": new_name,
                "section": new_section.strip(),
                "keyword": new_keyword.strip(),
            })
            st.success(f"✅ Rule '{new_name}' added.")
            st.rerun()

    st.markdown("---")
    st.markdown("**Examples of rules you can add:**")
    examples = [
        ("Invited Talks",   "Invited Talks",        "invited"),
        ("Book Chapters",   "Book Chapters",         "chapter"),
        ("Patents",         "Patents",               "patent"),
        ("Awards",          "Honors and Awards",     "award"),
        ("Editorials",      "Editorial",             "editor"),
    ]
    ex_cols = st.columns(len(examples))
    for col, (name, section, kw) in zip(ex_cols, examples):
        with col:
            st.markdown(f"**{name}**")
            st.caption(f"Section: _{section}_\nKeyword: _{kw}_")
            if st.button(f"Add", key=f"ex_{name}"):
                if not any(cr["name"] == name for cr in st.session_state["custom_rules"]):
                    st.session_state["custom_rules"].append(
                        {"name": name, "section": section, "keyword": kw}
                    )
                    st.rerun()
