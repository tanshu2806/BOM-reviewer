"""
H (Highlight) stage.

Outputs
  <name>_annotated.pdf     the ORIGINAL drawing with coloured boxes + short labels on every flagged item,
                           a legend on each page and a summary page at the end
  <name>_BOM_checked.xlsx  sheet 'BOM (checked)': original rows + 'Check status' / 'Check note' columns,
                           with MISSING parts INSERTED as red rows (a missing part is not in the BOM, so it has to be added)
                           sheet 'Discrepancies': one line per issue;  sheet 'Summary': counts + legend
  <name>_report.json       machine-readable result

The ORIGINAL files are never modified.
"""
import json, os
import warnings
warnings.filterwarnings("ignore")
import pymupdf as fitz
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

RGB = {"red": (0.85, 0.0, 0.0), "orange": (1.0, 0.5, 0.0), "yellow": (0.9, 0.72, 0.0), "blue": (0.1, 0.35, 0.95)}
FILL = {"red": "FFC7CE", "orange": "FFD9A0", "yellow": "FFF2A8", "blue": "CFE0FF", "green": "E2F0D9", "grey": "EDEDED"}
LABEL = {"omission": "MISSING IN BOM", "quantity": "QTY MISMATCH", "spec": "SPEC MISMATCH", "multiplier": "QTY x{N}?",
         "extra": "EXTRA IN BOM", "duplicate": "DUPLICATE IN BOM", "confirm": "CONFIRM MATCH", "implied": "BOM LACKS: ",
         "revision": "REVISION"}
SEV_ORDER = {"red": 0, "orange": 1, "yellow": 2, "blue": 3}


def _label(iss):
    t = LABEL.get(iss["type"], iss["type"].upper())
    if iss["type"] == "implied":
        return t + iss.get("label", "")
    if iss["type"] == "multiplier":
        return "QTY: BOM %g, expected %g" % (iss.get("bom_qty", 0), iss.get("expected", 0))
    if iss["type"] == "quantity" and "drawing_qty" in iss:
        return "QTY: dwg %g / BOM %g" % (iss["drawing_qty"], iss["bom_qty"])
    return t


# ------------------------------------------------------------------------------ annotated PDF
def annotate_pdf(src, dst, drawing, result, meta):
    doc = fitz.open(src)
    by_page = {}
    for iss in result["issues"]:
        if iss.get("bbox") is not None and iss.get("page") is not None:
            by_page.setdefault(iss["page"], []).append(iss)
    balloons = {}
    for p in drawing["pages"]:
        for b in p["balloons"]:
            balloons[(p["page"], b["id"])] = b["bbox"]

    for pno, page in enumerate(doc):
        sh = page.new_shape()
        labels = []
        for iss in sorted(by_page.get(pno, []), key=lambda i: -SEV_ORDER[i["severity"]]):
            col = RGB[iss["severity"]]
            r = fitz.Rect(iss["bbox"])
            sh.draw_rect(r)
            sh.finish(color=col, fill=col, width=1.4, fill_opacity=0.22, stroke_opacity=1)
            labels.append((r, col, _label(iss), iss["type"] in ("implied",)))
            if iss["type"] == "omission" and (pno, iss["drawing_id"]) in balloons:      # also circle the balloon on the sheet
                br = fitz.Rect(balloons[(pno, iss["drawing_id"])]) + (-3, -3, 3, 3)
                sh.draw_oval(br)
                sh.finish(color=col, fill=col, width=2.0, fill_opacity=0.25, stroke_opacity=1)
        sh.commit()
        for r, col, text, right_side in labels:
            w = max(60, 4.1 * len(text) + 6)
            if right_side:                                   # notes: label goes after the END of that note line (free space)
                same = [x for x in drawing["pages"][pno]["words"] if abs(x["yc"] - (r.y0 + r.y1) / 2) < 3.5]
                xend = max([x["x1"] for x in same] + [r.x1]) + 8
                page.insert_text((xend, r.y1 - 1.5), "<- " + text, fontsize=6.8, fontname="hebo", color=col)
                continue
            else:
                box = fitz.Rect(r.x0 - w - 4, r.y0 - 1, r.x0 - 4, r.y1 + 1)
            if (box.x0 < 5 or box.x1 > page.rect.width - 5) and not right_side:         # keep on the page
                box = fitz.Rect(r.x0, r.y1 + 1, r.x0 + w, r.y1 + 11)
            page.insert_textbox(box, text, fontsize=6.5, fontname="helv", color=col, align=0 if right_side else 2)
        # legend
        x, y = 28, 26
        banner = result["reliability"] != "HIGH"
        page.draw_rect(fitz.Rect(x - 4, y - 4, x + 250, y + (80 if banner else 62)), color=(0.3, 0.3, 0.3), fill=(1, 1, 1), width=0.6)
        if banner:
            bc = (0.85, 0, 0) if result["reliability"] == "LOW" else (0.9, 0.5, 0)
            page.insert_text((x, y + 74), f"RUN RELIABILITY: {result['reliability']} - " + ("DO NOT RELY ON THIS RUN" if result["reliability"] == "LOW" else "review flags with care"),
                             fontsize=7, fontname="hebo", color=bc)
        page.insert_text((x, y + 6), "BOM CHECK  (automated - verify before acting)", fontsize=6.5, fontname="hebo", color=(0, 0, 0))
        for k, (name, c) in enumerate([("red = on drawing, missing in BOM", "red"), ("orange = qty / spec / revision mismatch", "orange"),
                                       ("yellow = confirm match / extra / duplicate", "yellow"), ("blue = suggestion (implied item, rule-based)", "blue")]):
            yy = y + 16 + 11 * k
            page.draw_rect(fitz.Rect(x, yy - 6, x + 8, yy + 2), color=RGB[c], fill=RGB[c])
            page.insert_text((x + 13, yy + 1), name, fontsize=6.3, fontname="helv", color=(0, 0, 0))

    # summary page
    sp = doc.new_page(width=595, height=842)
    lines = [f"BOM CHECK SUMMARY", "",
             f"Drawing : {os.path.basename(meta['drawing'])}   (Rev {drawing['revision'] or '?'})",
             f"BOM     : {os.path.basename(meta['bom'])}   (Rev {meta['bom_revision'] or '?'})",
             f"Input   : {'scanned (OCR)' if drawing['scanned'] else 'vector PDF'}",
             f"RUN RELIABILITY : {result['reliability']}" + ("".join("\n   - " + x for x in result["reliability_reasons"])), "",
             f"Drawing items checked : {result['n_drawing_items']}",
             f"BOM lines read        : {result['n_bom_rows']}",
             f"BOM completeness      : {result['completeness']:.1f}%  (drawing items that have a BOM line)", ""]
    counts = {}
    for i in result["issues"]:
        counts[i["type"]] = counts.get(i["type"], 0) + 1
    lines.append("Findings: " + (", ".join(f"{k} {v}" for k, v in sorted(counts.items())) if counts else "none"))
    lines.append("")
    for i in sorted(result["issues"], key=lambda i: SEV_ORDER[i["severity"]]):
        lines.append(f"[{i['severity'].upper():6}] {i['message']}")
    lines += ["", "Automated check - not a substitute for engineering review. Blue items are rule-based suggestions."]
    sp.insert_textbox(fitz.Rect(40, 40, 555, 800), "\n".join(lines), fontsize=7.5, fontname="cour")
    doc.save(dst)
    doc.close()


# ------------------------------------------------------------------------------ Excel
def write_bom_xlsx(src, dst, bom, drawing, result, meta):
    wb_src = openpyxl.load_workbook(src)                       # keep formulas/styles of the original
    ws_src = wb_src[bom["sheet"]]
    hr, ncol = bom["header_row"], ws_src.max_column
    st_col, nt_col = ncol + 1, ncol + 2

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BOM (checked)"
    arial = Font(name="Arial", size=10)
    # copy header block + header row
    for r in range(1, hr + 1):
        for c in range(1, ncol + 1):
            cell = ws_src.cell(r, c)
            n = ws.cell(r, c, cell.value)
            n.font = Font(name="Arial", size=cell.font.sz or 10, bold=cell.font.b)
    for c, name in ((st_col, "Check status"), (nt_col, "Check note")):
        cell = ws.cell(hr, c, name)
        cell.font = Font(name="Arial", bold=True)
    for c in range(1, nt_col + 1):
        ws.cell(hr, c).fill = PatternFill("solid", fgColor="D9D9D9")
    for k, w in ws_src.column_dimensions.items():
        ws.column_dimensions[k].width = w.width
    ws.column_dimensions[get_column_letter(st_col)].width = 26
    ws.column_dimensions[get_column_letter(nt_col)].width = 80

    # status per BOM row
    by_row = {}
    for iss in result["issues"]:
        if iss.get("bom_row"):
            by_row.setdefault(iss["bom_row"], []).append(iss)
    drow_of = {m["bom_row"]: m for m in result["matches"] if m["bom_row"]}

    # where to insert each MISSING item: right after the BOM row matched to the nearest preceding drawing item
    order = [m for m in result["matches"]]
    missing = [(k, m) for k, m in enumerate(order) if m["status"] == "MISSING"]
    after_row = {}
    for k, m in missing:
        prev = next((order[q]["bom_row"] for q in range(k - 1, -1, -1) if order[q]["bom_row"]), None)
        after_row.setdefault(prev, []).append(m)
    ditem = {d["id"]: d for d in result["ditems"]}
    iss_missing = {i["drawing_id"]: i for i in result["issues"] if i["type"] == "omission" and i.get("drawing_id")}

    colmap = {canon: c for c, canon in bom["colidx"].items()}
    out_r = hr + 1

    def put_missing(m):
        nonlocal out_r
        d = ditem.get(m["drawing_id"])
        if d is None:
            return
        for canon, val in (("item", f"{d['id']} (dwg)"), ("description", d["description"]), ("material", d.get("material", "")),
                           ("qty", d.get("qty", "")), ("size", d.get("size", "")), ("partno", d.get("partno", ""))):
            if canon in colmap and val not in (None, ""):
                ws.cell(out_r, colmap[canon], val)
        ws.cell(out_r, st_col, "MISSING - ADD TO BOM")
        ws.cell(out_r, nt_col, f"Found on drawing (page {d['page'] + 1}, item {d['id']}). Not in the original BOM - proposed line, please review.")
        for c in range(1, nt_col + 1):
            ws.cell(out_r, c).fill = PatternFill("solid", fgColor=FILL["red"])
            ws.cell(out_r, c).font = Font(name="Arial", size=10, bold=(c == st_col))
        out_r += 1

    for m in after_row.get(None, []):
        put_missing(m)
    for b in bom["rows"]:
        r = b["row"]
        for c in range(1, ncol + 1):
            cell = ws_src.cell(r, c)
            n = ws.cell(out_r, c, cell.value)
            n.font = arial
        iss = by_row.get(r, [])
        if r in set(result["nondrawn_rows"]):
            status, note, color = "NON-DRAWN ITEM", "Consumable/non-drawn line (not expected on the drawing); covered by rule-based check", "grey"
        elif iss:
            iss_sorted = sorted(iss, key=lambda i: SEV_ORDER[i["severity"]])
            color = iss_sorted[0]["severity"]
            status = " + ".join(dict.fromkeys(_label(i).replace("MISMATCH", "MISMATCH") for i in iss_sorted))
            note = " | ".join(i["message"] for i in iss_sorted)
        else:
            m = drow_of.get(r)
            status, note, color = "OK", (f"Matches drawing item {m['drawing_id']} (score {m['score']:.0f})" if m else ""), "green"
        ws.cell(out_r, st_col, status).font = Font(name="Arial", size=10, bold=(status != "OK"))
        ws.cell(out_r, nt_col, note).font = arial
        ws.cell(out_r, nt_col).alignment = Alignment(wrap_text=True, vertical="top")
        for c in range(1, nt_col + 1):
            ws.cell(out_r, c).fill = PatternFill("solid", fgColor=FILL[color])
        out_r += 1
        for m in after_row.get(r, []):
            put_missing(m)
    ws.freeze_panes = ws.cell(hr + 1, 1)
    if result["reliability"] != "HIGH":
        c1 = ws.cell(1, st_col, f"RUN RELIABILITY: {result['reliability']} - " + ("DO NOT RELY ON THIS RUN. " if result["reliability"] == "LOW" else "") + "; ".join(result["reliability_reasons"]))
        c1.font = Font(name="Arial", bold=True, color="C00000")

    # revision cell highlight
    rev_issue = next((i for i in result["issues"] if i["type"] == "revision"), None)
    if rev_issue:
        for row in ws.iter_rows(min_row=1, max_row=hr - 1):
            for cell in row:
                if isinstance(cell.value, str) and "rev" in cell.value.lower():
                    cell.fill = PatternFill("solid", fgColor=FILL["orange"])
                    cell.value = cell.value + "   <-- " + rev_issue["message"]

    # Discrepancies sheet
    wd = wb.create_sheet("Discrepancies")
    heads = ["#", "Severity", "Type", "Drawing item", "Drawing description", "BOM row", "Message"]
    for c, h in enumerate(heads, start=1):
        cell = wd.cell(1, c, h); cell.font = Font(name="Arial", bold=True); cell.fill = PatternFill("solid", fgColor="D9D9D9")
    for n, i in enumerate(sorted(result["issues"], key=lambda i: SEV_ORDER[i["severity"]]), start=1):
        vals = [n, i["severity"].upper(), i["type"], i.get("drawing_id"), i.get("drawing_desc"), i.get("bom_row"), i["message"]]
        for c, v in enumerate(vals, start=1):
            cell = wd.cell(n + 1, c, v); cell.font = arial
            cell.fill = PatternFill("solid", fgColor=FILL[i["severity"]])
    for c, w in enumerate([5, 10, 12, 12, 42, 9, 110], start=1):
        wd.column_dimensions[get_column_letter(c)].width = w

    # Summary sheet
    wsu = wb.create_sheet("Summary")
    rows = [("BOM CHECK SUMMARY", ""), ("", ""), ("Drawing", os.path.basename(meta["drawing"])), ("Drawing revision", drawing["revision"] or "not found"),
            ("BOM", os.path.basename(meta["bom"])), ("BOM revision", meta["bom_revision"] or "not found"),
            ("Drawing input type", "scanned (OCR)" if drawing["scanned"] else "vector PDF"),
            ("RUN RELIABILITY", result["reliability"] + ("  -  " + "; ".join(result["reliability_reasons"]) if result["reliability_reasons"] else "")),
            ("Drawing items checked", result["n_drawing_items"]), ("BOM lines read", result["n_bom_rows"]),
            ("BOM completeness (drawing items with a BOM line)", round(result["completeness"], 1) if result["completeness"] is not None else "n/a")]
    counts = {}
    for i in result["issues"]:
        counts[i["type"]] = counts.get(i["type"], 0) + 1
    rows += [("", ""), ("Findings by type", "")] + sorted(counts.items())
    rows += [("", ""), ("Legend", ""), ("Red", "On drawing, missing in BOM (proposed line inserted)"), ("Orange", "Quantity / specification / revision mismatch"),
             ("Yellow", "Probable match to confirm, extra line, or duplicate line"), ("Blue", "Rule-based suggestion (implied item) - not a hard error"),
             ("Green", "Matched, no issue"), ("Grey", "Non-drawn item (e.g. consumables)"),
             ("", ""), ("Note", "Automated first-pass check. Thresholds/rules come from config.json (assumptions) and must be tuned on real client data.")]
    for r, (a, b) in enumerate(rows, start=1):
        wsu.cell(r, 1, a).font = Font(name="Arial", bold=(r == 1 or b == ""))
        wsu.cell(r, 2, b).font = arial
    wsu.column_dimensions["A"].width = 50; wsu.column_dimensions["B"].width = 90
    wb.save(dst)


def write_json(dst, drawing, bom, result, meta):
    clean = [{k: v for k, v in i.items()} for i in result["issues"]]
    with open(dst, "w") as f:
        json.dump(dict(meta=meta, reliability=result["reliability"], reliability_reasons=result["reliability_reasons"],
                       drawing_revision=drawing["revision"], bom_revision=bom["revision"], scanned=drawing["scanned"],
                       completeness=result["completeness"], n_drawing_items=result["n_drawing_items"], n_bom_rows=result["n_bom_rows"],
                       matches=result["matches"], issues=clean), f, indent=2, default=str)
