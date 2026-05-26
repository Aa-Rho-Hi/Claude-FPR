import streamlit as st
import tempfile, os, glob, re, sys, io, pickle, importlib

st.set_page_config(
    page_title="FAR Extraction Pipeline",
    page_icon="📄",
    layout="wide",
)

# ── Import run_rules at startup so errors surface immediately ──────────────────
try:
    import run_rules as rr
    _import_ok = True
except Exception as _e:
    _import_ok = False
    _import_err = str(_e)

if not _import_ok:
    st.error(f"❌ Failed to load run_rules.py: {_import_err}")
    st.stop()

# ── Page ───────────────────────────────────────────────────────────────────────
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

# ── Upload ─────────────────────────────────────────────────────────────────────
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

# ── Faculty selection ──────────────────────────────────────────────────────────
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

# ── Run ────────────────────────────────────────────────────────────────────────
st.subheader("3. Run Extraction")
run_btn = st.button("▶ Run Extraction", type="primary", disabled=not uploaded_files)

if run_btn and uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:

        # Save uploaded files to temp dir
        for uf in uploaded_files:
            with open(os.path.join(tmpdir, uf.name), "wb") as fh:
                fh.write(uf.getbuffer())

        # Determine faculty list
        if run_mode == "All detected faculty":
            faculty_list = list(rr.KNOWN_FACULTY)
            for fp in glob.glob(os.path.join(tmpdir, "F180Vita_*.pdf")):
                m = re.match(r'F180Vita_\w+\.(\w+)\.pdf', os.path.basename(fp))
                if m and m.group(1) not in faculty_list:
                    faculty_list.append(m.group(1))
            # Keep only faculty whose FAR was actually uploaded
            faculty_list = [
                ln for ln in faculty_list
                if glob.glob(os.path.join(tmpdir, f"F180Vita_*.{ln}.pdf"))
            ]
        else:
            faculty_list = [n.strip() for n in specific_names.split(",") if n.strip()]

        if not faculty_list:
            st.warning("⚠️ No faculty detected. Check that FAR PDFs follow the naming pattern `F180Vita_F.Lastname.pdf`.")
            st.stop()

        st.info(f"Found {len(faculty_list)} faculty to process: {', '.join(faculty_list)}")

        # Pre-pass: build all_far_data (no disk cache on Streamlit Cloud)
        all_far_data = {}
        with st.spinner("Pre-parsing FAR PDFs for cross-referencing…"):
            for far_path in glob.glob(os.path.join(tmpdir, "F180Vita_*.pdf")):
                m = re.match(r'F180Vita_\w+\.(\w+)\.pdf', os.path.basename(far_path))
                if not m:
                    continue
                ln = m.group(1)
                try:
                    all_far_data[ln] = (rr.parse_far(far_path), rr.pdf_full_text(far_path))
                except Exception as e:
                    st.warning(f"Could not pre-parse {os.path.basename(far_path)}: {e}")

        # Process each faculty
        st.markdown("---")
        st.subheader("Results")
        progress = st.progress(0, text="Starting…")
        log_area = st.empty()
        results = []
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

            if r:
                results.append(r)
                log_lines[-1] = (
                    f"✅ **{last_name}** — "
                    f"UG={r['ug']} Grad={r['grad']} MS={r['ms']} PhD={r['phd']} | "
                    f"Grants={r['grants']} CH/CO={r['ch_co']} CP={r['cp']} Journal={r['journal']}"
                )
            log_area.markdown("\n\n".join(log_lines))

        progress.progress(1.0, text="Done!")

        if not results:
            st.error("No results produced. Check file naming and try again.")
            st.stop()

        # Generate Excel
        out_path = os.path.join(tmpdir, output_filename)
        rr.generate_excel(results, out_path)
        with open(out_path, "rb") as fh:
            excel_bytes = fh.read()

        # Summary table
        st.markdown("---")
        st.subheader("📊 Summary")
        import pandas as pd
        rows = [{
            "Last Name":  r["last_name"],
            "First Name": r["first_name"],
            "Title":      r["title"],
            "UG":         r["ug"],
            "Grad":       r["grad"],
            "MS":         r["ms"],
            "PhD":        r["phd"],
            "Grants":     r["grants"],
            "CH/CO":      r["ch_co"],
            "CP":         r["cp"],
            "Journal":    r["journal"],
        } for r in results]
        st.dataframe(pd.DataFrame(rows).set_index("Last Name"), use_container_width=True)

        st.download_button(
            label=f"⬇️  Download {output_filename}",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
