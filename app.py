import streamlit as st
import tempfile, os, glob, re, sys, io, pickle
from pathlib import Path

st.set_page_config(
    page_title="FAR Extraction Pipeline",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Faculty Annual Report Extraction Pipeline")
st.caption("Upload FAR PDFs, CV PDFs, and supplemental XLSX files to extract structured data into Excel.")

# ── Sidebar: settings ──────────────────────────────────────────────────────────
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
    st.markdown("- FAR PDF: `F180Vita_F.Lastname.pdf`")
    st.markdown("- CV PDF: `Lastname CV.pdf`")
    st.markdown("- Supplemental XLSX: `Lastname.xlsx`")
    st.markdown("- Support staff XLSX: `support_staff.xlsx`")

# ── File upload ────────────────────────────────────────────────────────────────
st.subheader("1. Upload Files")
uploaded_files = st.file_uploader(
    "Upload all files (FAR PDFs, CV PDFs, XLSX files)",
    accept_multiple_files=True,
    type=["pdf", "xlsx", "xls"],
    help="You can select all files at once. Mix of FAR PDFs, CV PDFs, and supplemental workbooks.",
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded")
    with st.expander("Uploaded files", expanded=False):
        for f in uploaded_files:
            st.text(f"  {f.name}  ({f.size/1024:.1f} KB)")

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
        placeholder="e.g. Narayanan, Palermo, Qian",
    )

# ── Run ────────────────────────────────────────────────────────────────────────
st.subheader("3. Run Extraction")
run_btn = st.button("▶ Run Extraction", type="primary", disabled=not uploaded_files)

if run_btn and uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        # Save uploaded files to temp dir
        for uf in uploaded_files:
            dest = os.path.join(tmpdir, uf.name)
            with open(dest, "wb") as fh:
                fh.write(uf.getbuffer())

        # Add run_rules to path (must be in same folder as app.py)
        app_dir = os.path.dirname(os.path.abspath(__file__))
        if app_dir not in sys.path:
            sys.path.insert(0, app_dir)

        try:
            import run_rules as rr
        except ImportError:
            st.error("❌ `run_rules.py` not found. Make sure it is in the same folder as `app.py`.")
            st.stop()

        # Determine faculty list
        if run_mode == "All detected faculty":
            faculty_list = list(rr.KNOWN_FACULTY)
            for fp in glob.glob(os.path.join(tmpdir, "F180Vita_*.pdf")):
                m = re.match(r'F180Vita_\w+\.(\w+)\.pdf', os.path.basename(fp))
                if m and m.group(1) not in faculty_list:
                    faculty_list.append(m.group(1))
        else:
            faculty_list = [n.strip() for n in specific_names.split(",") if n.strip()]

        if not faculty_list:
            st.warning("No faculty to process. Check uploaded file names or enter last names above.")
            st.stop()

        # Pre-pass: build all_far_data with pickle cache
        cache_path = os.path.join(app_dir, ".far_cache.pkl")
        try:
            with open(cache_path, "rb") as cf:
                _cache = pickle.load(cf)
        except Exception:
            _cache = {}

        all_far_data = {}
        _cache_dirty = False
        for far_path in glob.glob(os.path.join(tmpdir, "F180Vita_*.pdf")):
            m = re.match(r'F180Vita_\w+\.(\w+)\.pdf', os.path.basename(far_path))
            if not m:
                continue
            ln = m.group(1)
            try:
                mtime = os.path.getmtime(far_path)
                cache_key = (far_path, mtime)
                if cache_key in _cache:
                    all_far_data[ln] = _cache[cache_key]
                else:
                    parsed = (rr.parse_far(far_path), rr.pdf_full_text(far_path))
                    all_far_data[ln] = parsed
                    _cache[cache_key] = parsed
                    _cache_dirty = True
            except Exception:
                pass

        if _cache_dirty:
            try:
                with open(cache_path, "wb") as cf:
                    pickle.dump(_cache, cf)
            except Exception:
                pass

        # Process each faculty with live progress
        st.markdown("---")
        st.subheader("Results")
        progress = st.progress(0, text="Starting…")
        log_area = st.empty()
        results = []
        log_lines = []

        for i, last_name in enumerate(faculty_list):
            progress.progress((i) / len(faculty_list), text=f"Processing {last_name}…")
            log_lines.append(f"⏳ Processing **{last_name}**…")
            log_area.markdown("\n\n".join(log_lines))

            # Capture stdout from extract_faculty
            old_stdout = sys.stdout
            sys.stdout = buf = io.StringIO()
            try:
                r = rr.extract_faculty(
                    last_name,
                    input_dir=tmpdir,
                    api_key=api_key or None,
                    all_far_data=all_far_data,
                )
            except Exception as e:
                r = None
                log_lines[-1] = f"❌ **{last_name}** — error: {e}"
            finally:
                sys.stdout = old_stdout

            if r:
                results.append(r)
                log_lines[-1] = (
                    f"✅ **{last_name}** — "
                    f"UG={r['ug']} Grad={r['grad']} MS={r['ms']} PhD={r['phd']} "
                    f"Grants={r['grants']} CH/CO={r['ch_co']} CP={r['cp']} Journal={r['journal']}"
                )
            log_area.markdown("\n\n".join(log_lines))

        progress.progress(1.0, text="Done!")

        if not results:
            st.error("No results produced. Check that FAR PDFs match the expected naming pattern.")
            st.stop()

        # Generate Excel in memory
        out_path = os.path.join(tmpdir, output_filename)
        rr.generate_excel(results, out_path)

        with open(out_path, "rb") as fh:
            excel_bytes = fh.read()

        # Summary table
        st.markdown("---")
        st.subheader("Summary")
        import pandas as pd
        rows = []
        for r in results:
            rows.append({
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
            })
        st.dataframe(pd.DataFrame(rows).set_index("Last Name"), use_container_width=True)

        # Download button
        st.download_button(
            label=f"⬇️ Download {output_filename}",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
