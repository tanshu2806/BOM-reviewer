import json
import os
import shutil
import time
import uuid
from pathlib import Path

import streamlit as st

from run_check import run

ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "ui_outputs"
OUTPUT_RETENTION_SECONDS = 24 * 60 * 60

SEVERITY_ORDER = {"red": 0, "orange": 1, "yellow": 2, "blue": 3, "green": 4, "grey": 5}


st.set_page_config(page_title="BOM Completeness Checker", page_icon="📐", layout="wide")

st.markdown(
    """
    <style>
    .main { background: #f5f7fb; }
    .stApp { background: linear-gradient(180deg, #f4f7fb 0%, #eef4ff 100%); color: #0b1f33; }
    .block-container { padding-top: 1.5rem; }
    div[data-testid="stSidebar"] { background: #e8eef7; }
    h1, h2, h3, h4, h5, p, label { color: #0b1f33; }
    .metric-card { background: white; border-radius: 12px; padding: 1rem; box-shadow: 0 1px 4px rgba(0,0,0,0.08); }
    .status-pill { padding: 0.3rem 0.7rem; border-radius: 999px; font-weight: 600; }
    div[data-testid="stFileUploader"] section { background: white; border-color: #d1d5db; }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def get_sample_files():
    sample_dir = ROOT / "samples"
    drawing = sample_dir / "sample_01_drawing.pdf"
    bom = sample_dir / "sample_01_bom.xlsx"
    return drawing, bom


def download_button_for(path: Path, label: str, mime: str):
    if path.exists():
        with open(path, "rb") as f:
            data = f.read()
        st.download_button(label=label, data=data, file_name=path.name, mime=mime)


def load_report(report_path: Path):
    if not report_path.exists():
        return None
    with open(report_path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def get_output_run_dir(path):
    candidate = Path(path).resolve()
    if (
        candidate.parent == OUTPUT_ROOT.resolve()
        and candidate.name.startswith(("run_", "sample_run_"))
    ):
        return candidate
    return None


def remove_output_run(path):
    run_dir = get_output_run_dir(path)
    if run_dir is not None and run_dir.is_dir():
        shutil.rmtree(run_dir)


def cleanup_expired_outputs():
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    latest_downloads = st.session_state.get("latest_downloads", {})
    if latest_downloads:
        active_dir = get_output_run_dir(
            latest_downloads.get("output_dir", latest_downloads["annotated_pdf"])
        )
        if active_dir is not None and active_dir.is_dir():
            os.utime(active_dir, None)

    now = time.time()
    for path in OUTPUT_ROOT.iterdir():
        run_dir = get_output_run_dir(path)
        if run_dir is not None and run_dir.is_dir():
            if now - run_dir.stat().st_mtime > OUTPUT_RETENTION_SECONDS:
                shutil.rmtree(run_dir)


def count_by_type(issues):
    counts = {}
    for issue in issues:
        key = issue.get("type", "unknown")
        counts[key] = counts.get(key, 0) + 1
    return dict(sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])))


cleanup_expired_outputs()


def render_summary(report):
    if not report:
        return

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown("<div class='metric-card'><h4>Completeness</h4><h2>{:.1f}%</h2></div>".format(report.get("completeness", 0)), unsafe_allow_html=True)
    with col2:
        st.markdown("<div class='metric-card'><h4>Drawing items</h4><h2>{}</h2></div>".format(report.get("n_drawing_items", 0)), unsafe_allow_html=True)
    with col3:
        st.markdown("<div class='metric-card'><h4>BOM rows</h4><h2>{}</h2></div>".format(report.get("n_bom_rows", 0)), unsafe_allow_html=True)
    with col4:
        reliability = report.get("reliability", "UNKNOWN")
        color = {"HIGH": "#2e7d32", "MEDIUM": "#ef6c00", "LOW": "#c62828"}.get(reliability, "#455a64")
        st.markdown(
            f"<div class='metric-card'><h4>Reliability</h4><div class='status-pill' style='background:{color}; color:white; display:inline-block;'>{reliability}</div></div>",
            unsafe_allow_html=True,
        )

    st.subheader("Issue overview")
    type_counts = count_by_type(report.get("issues", []))
    if type_counts:
        cols = st.columns(min(len(type_counts), 4))
        for idx, (issue_type, count) in enumerate(type_counts.items()):
            with cols[idx % 4]:
                st.markdown(
                    f"<div class='metric-card'><h5>{issue_type}</h5><h3>{count}</h3></div>",
                    unsafe_allow_html=True,
                )
    else:
        st.info("No issues were flagged in this run.")


st.title("📐 BOM Completeness Checker")
st.caption("Prototype dashboard for comparing an engineering drawing against a BOM and highlighting mismatches.")

with st.sidebar:
    st.header("Project status")
    st.write("This UI wraps the prototype checker in this repo and is designed for demo presentations.")
    st.markdown("- Vector or OCR-based scanned drawing support")
    st.markdown("- BOM validation and mismatch reporting")
    st.markdown("- Annotated PDF + Excel + JSON output")

    sample_drawing, sample_bom = get_sample_files()
    if sample_drawing.exists() and sample_bom.exists():
        if st.button("Use sample data"):
            st.session_state["sample_mode"] = True
            st.session_state["selected_drawing"] = str(sample_drawing)
            st.session_state["selected_bom"] = str(sample_bom)
            st.rerun()

st.subheader("Upload your files")

upload_widget_version = st.session_state.setdefault("upload_widget_version", 0)
if st.session_state.get("original_uploads") or st.session_state.get("latest_downloads"):
    if st.button("Clear session"):
        latest_downloads = st.session_state.get("latest_downloads", {})
        output_dir = latest_downloads.get("output_dir", latest_downloads.get("annotated_pdf"))
        if output_dir:
            remove_output_run(output_dir)
        for key in [
            "original_uploads",
            "latest_downloads",
            "sample_mode",
            "selected_drawing",
            "selected_bom",
        ]:
            st.session_state.pop(key, None)
        st.session_state["upload_widget_version"] = upload_widget_version + 1
        st.rerun()

col1, col2 = st.columns(2)
with col1:
    drawing_file = st.file_uploader(
        "Drawing PDF", type=["pdf"], key=f"drawing_upload_{upload_widget_version}"
    )
with col2:
    bom_file = st.file_uploader(
        "BOM Excel", type=["xlsx", "xls", "xlsm"], key=f"bom_upload_{upload_widget_version}"
    )

config_file = st.file_uploader(
    "Optional config override (.json)", type=["json"], key=f"config_upload_{upload_widget_version}"
)

original_uploads = dict(st.session_state.get("original_uploads", {}))
if drawing_file is not None:
    original_uploads["drawing"] = drawing_file.name
if bom_file is not None:
    original_uploads["bom"] = bom_file.name
if config_file is not None:
    original_uploads["config"] = config_file.name
if original_uploads:
    st.session_state["original_uploads"] = original_uploads
    st.caption(
        "Selected input files — "
        f"Drawing: {original_uploads.get('drawing', 'not selected')} · "
        f"BOM: {original_uploads.get('bom', 'not selected')}"
        + (f" · Config: {original_uploads['config']}" if "config" in original_uploads else "")
    )

run_check = st.button("Run BOM Check")

if run_check:
    if drawing_file is None or bom_file is None:
        st.error("Please upload both a drawing PDF and a BOM Excel file before running the check.")
    else:
        OUTPUT_ROOT.mkdir(exist_ok=True)
        run_id = uuid.uuid4().hex[:8]
        out_dir = OUTPUT_ROOT / f"run_{run_id}"
        out_dir.mkdir(exist_ok=True)

        drawing_path = out_dir / Path(drawing_file.name).name
        bom_path = out_dir / Path(bom_file.name).name
        config_path = out_dir / "config.json" if config_file is not None else None

        drawing_path.write_bytes(drawing_file.getvalue())
        bom_path.write_bytes(bom_file.getvalue())
        if config_file is not None:
            config_path.write_bytes(config_file.getvalue())

        with st.spinner("Running the drawing-to-BOM validation pipeline..."):
            try:
                run(str(drawing_path), str(bom_path), str(out_dir), str(config_path) if config_path else None, quiet=True)
            except Exception as exc:
                st.exception(exc)
                st.stop()

        stem = drawing_path.stem
        annotated_pdf = out_dir / f"{stem}_annotated.pdf"
        checked_xlsx = out_dir / f"{stem}_BOM_checked.xlsx"
        report_json = out_dir / f"{stem}_report.json"

        report = load_report(report_json)
        if report is None:
            st.error("The report was not generated correctly.")
            st.stop()

        st.session_state["latest_downloads"] = {
            "annotated_pdf": str(annotated_pdf),
            "checked_xlsx": str(checked_xlsx),
            "drawing_name": drawing_file.name,
            "bom_name": bom_file.name,
            "output_dir": str(out_dir),
        }

        st.success("Check completed successfully.")
        render_summary(report)

        st.subheader("Findings")
        issues = report.get("issues", [])
        if issues:
            for issue in sorted(issues, key=lambda item: (SEVERITY_ORDER.get(item.get("severity", "green"), 99), item.get("message", ""))):
                severity = issue.get("severity", "green").upper()
                if severity == "RED":
                    color = "#d32f2f"
                elif severity == "ORANGE":
                    color = "#f57c00"
                elif severity == "YELLOW":
                    color = "#f9a825"
                elif severity == "BLUE":
                    color = "#1976d2"
                else:
                    color = "#4caf50"

                st.markdown(
                    f"<div style='padding:0.8rem 1rem; border-left:5px solid {color}; background:var(--card-background); color:var(--app-text); border-radius:10px; margin-bottom:0.75rem;'>"
                    f"<strong>[{severity}]</strong> {issue.get('message', 'No message provided')}"
                    f"</div>",
                    unsafe_allow_html=True,
                )
        else:
            st.info("No mismatches or exceptions were identified in this run.")

        if report.get("reliability") != "HIGH":
            st.warning("This run was flagged with lower reliability. Review the generated issues before making engineering decisions.")

else:
    st.info("Upload a drawing and BOM to start the validation workflow.")

    sample_drawing, sample_bom = get_sample_files()
    if sample_drawing.exists() and sample_bom.exists():
        with st.container():
            st.caption("Sample files available in the project: sample_01_drawing.pdf and sample_01_bom.xlsx")
            if st.session_state.pop("sample_mode", False) or st.button("Quick preview with sample files"):
                OUTPUT_ROOT.mkdir(exist_ok=True)
                run_id = uuid.uuid4().hex[:8]
                out_dir = OUTPUT_ROOT / f"sample_run_{run_id}"
                out_dir.mkdir(exist_ok=True)
                drawing_path = sample_drawing
                bom_path = sample_bom
                with st.spinner("Running sample check..."):
                    run(str(drawing_path), str(bom_path), str(out_dir), quiet=True)
                stem = drawing_path.stem
                annotated_pdf = out_dir / f"{stem}_annotated.pdf"
                checked_xlsx = out_dir / f"{stem}_BOM_checked.xlsx"
                report_json = out_dir / f"{stem}_report.json"
                report = load_report(report_json)
                if report:
                    st.session_state["latest_downloads"] = {
                        "annotated_pdf": str(annotated_pdf),
                        "checked_xlsx": str(checked_xlsx),
                        "drawing_name": sample_drawing.name,
                        "bom_name": sample_bom.name,
                        "output_dir": str(out_dir),
                    }
                    st.session_state["original_uploads"] = {
                        "drawing": sample_drawing.name,
                        "bom": sample_bom.name,
                    }
                    render_summary(report)
                else:
                    st.warning("Sample report could not be loaded.")

latest_downloads = st.session_state.get("latest_downloads")
if latest_downloads:
    st.subheader("Downloads")
    c1, c2 = st.columns(2)
    annotated_pdf = Path(latest_downloads["annotated_pdf"])
    checked_xlsx = Path(latest_downloads["checked_xlsx"])
    with c1:
        if annotated_pdf.is_file():
            st.caption(f"Annotated from: {latest_downloads['drawing_name']}")
            download_button_for(annotated_pdf, "Download annotated PDF", "application/pdf")
        else:
            st.error(f"Annotated PDF is missing: {annotated_pdf.name}")
    with c2:
        if checked_xlsx.is_file():
            st.caption(f"Checked from: {latest_downloads['bom_name']}")
            download_button_for(
                checked_xlsx,
                "Download checked BOM Excel",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error(f"Checked BOM Excel is missing: {checked_xlsx.name}")
