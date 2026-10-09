import json
import os
import shutil
import uuid
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
from run_check import run
from bomcheck.extract import load_config

app = Flask(__name__, static_folder="static", template_folder="templates")

ROOT_DIR = Path(__file__).resolve().parent
RUNS_DIR = ROOT_DIR / "ui_outputs"
SAMPLES_DIR = ROOT_DIR / "samples"
CONFIG_PATH = ROOT_DIR / "bomcheck" / "config.json"

RUNS_DIR.mkdir(parents=True, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config", methods=["GET", "POST"])
def manage_config():
    if request.method == "GET":
        try:
            cfg = load_config(str(CONFIG_PATH))
            return jsonify({"status": "success", "config": cfg})
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 500
    else:
        try:
            new_cfg = request.get_json()
            if not new_cfg:
                return jsonify({"status": "error", "message": "Invalid JSON body"}), 400
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(new_cfg, f, indent=2)
            return jsonify({"status": "success", "message": "Configuration updated successfully."})
        except Exception as e:
            return jsonify({"status": "error", "message": str(e)}), 500


@app.route("/api/check", methods=["POST"])
def run_bom_check():
    try:
        use_sample = request.form.get("use_sample") == "true" or (request.is_json and request.json.get("use_sample"))
        
        run_id = uuid.uuid4().hex[:8]
        out_dir = RUNS_DIR / f"run_{run_id}"
        out_dir.mkdir(parents=True, exist_ok=True)
        
        if use_sample:
            drawing_src = SAMPLES_DIR / "sample_01_drawing.pdf"
            bom_src = SAMPLES_DIR / "sample_01_bom.xlsx"
            if not drawing_src.exists() or not bom_src.exists():
                return jsonify({"status": "error", "message": "Sample files not found in repository."}), 404
            
            drawing_path = out_dir / drawing_src.name
            bom_path = out_dir / bom_src.name
            shutil.copy(drawing_src, drawing_path)
            shutil.copy(bom_src, bom_path)
            custom_config_path = None
        else:
            drawing_file = request.files.get("drawing")
            bom_file = request.files.get("bom")
            config_file = request.files.get("config")
            
            if not drawing_file or not bom_file:
                return jsonify({"status": "error", "message": "Both Drawing PDF and BOM Excel files are required."}), 400
            
            drawing_path = out_dir / Path(drawing_file.filename).name
            bom_path = out_dir / Path(bom_file.filename).name
            
            drawing_file.save(drawing_path)
            bom_file.save(bom_path)
            
            custom_config_path = None
            if config_file:
                custom_config_path = out_dir / "config_override.json"
                config_file.save(custom_config_path)

        # Run pipeline catching SystemExit & Exception
        try:
            drawing_info, bom_info, result = run(
                drawing_path=str(drawing_path),
                bom_path=str(bom_path),
                out_dir=str(out_dir),
                config_path=str(custom_config_path) if custom_config_path else None,
                quiet=True
            )
        except SystemExit as se:
            return jsonify({"status": "error", "message": str(se)}), 200
        except Exception as ex:
            import traceback
            traceback.print_exc()
            return jsonify({"status": "error", "message": f"Pipeline execution failed: {str(ex)}"}), 500

        stem = drawing_path.stem
        annotated_pdf = out_dir / f"{stem}_annotated.pdf"
        checked_xlsx = out_dir / f"{stem}_BOM_checked.xlsx"
        report_json = out_dir / f"{stem}_report.json"
        
        report_data = {}
        if report_json.exists():
            with open(report_json, "r", encoding="utf-8") as f:
                report_data = json.load(f)

        return jsonify({
            "status": "success",
            "run_id": run_id,
            "drawing_name": drawing_path.name,
            "bom_name": bom_path.name,
            "annotated_pdf_filename": annotated_pdf.name if annotated_pdf.exists() else None,
            "checked_xlsx_filename": checked_xlsx.name if checked_xlsx.exists() else None,
            "report": report_data
        })

    except Exception as e:
        return jsonify({"status": "error", "message": f"Unexpected server error: {str(e)}"}), 500


@app.route("/api/download/<run_id>/<filename>")
def download_file(run_id, filename):
    run_dir = RUNS_DIR / f"run_{run_id}"
    target_file = run_dir / filename
    if not target_file.exists() or not target_file.is_file():
        return jsonify({"status": "error", "message": "File not found."}), 404
    return send_file(target_file, as_attachment=True)


@app.route("/api/view/<run_id>/<filename>")
def view_file(run_id, filename):
    run_dir = RUNS_DIR / f"run_{run_id}"
    target_file = run_dir / filename
    if not target_file.exists() or not target_file.is_file():
        return jsonify({"status": "error", "message": "File not found."}), 404
    mime_type = "application/pdf" if filename.endswith(".pdf") else "application/octet-stream"
    return send_file(target_file, mimetype=mime_type)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
