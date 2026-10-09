#!/usr/bin/env python3
"""
Drawing-vs-BOM completeness checker (prototype).

Usage:
    python run_check.py DRAWING.pdf BOM.xlsx [--out outputs] [--config path/to/config.json]

Produces (in --out):
    <drawing>_annotated.pdf, <drawing>_BOM_checked.xlsx, <drawing>_report.json
Runs fully offline. Original files are never modified.
"""
import argparse, os, sys, warnings
warnings.filterwarnings("ignore")
from bomcheck.extract import load_config, read_drawing, read_bom
from bomcheck.match import compare
from bomcheck.report import annotate_pdf, write_bom_xlsx, write_json


def run(drawing_path, bom_path, out_dir="outputs", config_path=None, quiet=False):
    cfg = load_config(config_path)
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(drawing_path))[0]
    drawing = read_drawing(drawing_path, cfg)
    bom = read_bom(bom_path, cfg)
    result = compare(drawing, bom, cfg)
    meta = dict(drawing=drawing_path, bom=bom_path, bom_revision=bom["revision"])
    annotate_pdf(drawing_path, os.path.join(out_dir, f"{stem}_annotated.pdf"), drawing, result, meta)
    write_bom_xlsx(bom_path, os.path.join(out_dir, f"{stem}_BOM_checked.xlsx"), bom, drawing, result, meta)
    write_json(os.path.join(out_dir, f"{stem}_report.json"), drawing, bom, result, meta)
    if not quiet:
        print(f"Drawing: {os.path.basename(drawing_path)}  ({'scanned/OCR' if drawing['scanned'] else 'vector'}, Rev {drawing['revision']})")
        print(f"BOM    : {os.path.basename(bom_path)}  (sheet '{bom['sheet']}', Rev {bom['revision']})")
        if result["reliability"] != "HIGH":
            print(f"*** RUN RELIABILITY: {result['reliability']} ***  " + "; ".join(result["reliability_reasons"]))
        completeness = f"{result['completeness']:.1f}%" if result["completeness"] is not None else "n/a"
        print(f"Items on drawing: {result['n_drawing_items']}   BOM lines: {result['n_bom_rows']}   completeness: {completeness}")
        for i in sorted(result["issues"], key=lambda i: ["red", "orange", "yellow", "blue"].index(i["severity"])):
            print(f"  [{i['severity']:6}] {i['message']}")
        print(f"Outputs in {out_dir}/")
    return drawing, bom, result


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("drawing"); ap.add_argument("bom")
    ap.add_argument("--out", default="outputs"); ap.add_argument("--config")
    a = ap.parse_args()
    run(a.drawing, a.bom, a.out, a.config)
