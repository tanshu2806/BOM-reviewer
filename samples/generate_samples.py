"""
Synthetic proxy-data generator (Phase 0 substitute, since real client data is not available pre-MoU).

Creates, per sample:
  sample_XX_drawing.pdf   vector PDF: schematic + balloons + parts table + notes + title block
  sample_XX_bom.xlsx      BOM with INJECTED errors
  sample_XX_truth.json    ground truth of every injected error
Also writes scanned (image-only, noisy) copies of some drawings: sample_XX_drawing_scan.pdf

ASSUMPTIONS baked in (see assumption register): English, metric, parts table with balloons,
Excel BOM, boiler economiser sub-system. Style varies between samples so the checker is not
tuned to ONE layout.
"""
import json, math, os, random, sys
from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------------------------------
# Item catalogue (boiler economiser sub-system).  draw / bom = wording variants per side.
# drawn=False -> exists only in the BOM (consumables); balloon=False -> in table, no balloon.
# ----------------------------------------------------------------------------------------
CAT = [
 dict(k="inlet_header",  draw=["INLET HEADER 219.1 OD X 8.0 THK"],  bom=["Header, inlet, 219.1 x 8.0 mm", "Inlet header 219.1 OD x 8 thk"], mat=("SA 106 GR.B","SA106 Gr.B"), qty=(1,1), balloon=True),
 dict(k="outlet_header", draw=["OUTLET HEADER 219.1 OD X 8.0 THK"], bom=["Header, outlet, 219.1 x 8.0 mm", "Outlet header 219.1 OD x 8 thk"], mat=("SA 106 GR.B","SA106 Gr.B"), qty=(1,1), balloon=True),
 dict(k="tube",          draw=["ECONOMISER TUBE 38.1 OD X 3.6 THK"], bom=["Tube, economiser, 38.1 x 3.6 mm", "Economiser tube 38.1 OD x 3.6 thk"], mat=("SA 210 GR.A1","SA210 Gr.A1"), qty=(96,144), balloon=True),
 dict(k="return_bend",   draw=["RETURN BEND 38.1 OD"],               bom=["Return bend 38.1 OD", "U-bend, 38.1 OD"], mat=("SA 210 GR.A1","SA210 Gr.A1"), qty=(24,48), balloon=True),
 dict(k="support_plate", draw=["TUBE SUPPORT PLATE 12 THK"],         bom=["Plate, tube support, 12 mm", "Tube support plate 12 thk"], mat=("IS 2062 GR.B","IS2062 Gr.B"), qty=(10,14), balloon=True),
 dict(k="inlet_flange",  draw=["INLET FLANGE 200 NB CL.150"],        bom=["Flange, inlet, WNRF, 200 NB, 150#", "Inlet flange 200 NB Class 150"], mat=("SA 105","SA105"), qty=(1,1), balloon=True),
 dict(k="outlet_flange", draw=["OUTLET FLANGE 200 NB CL.150"],       bom=["Flange, outlet, WNRF, 200 NB, 150#", "Outlet flange 200 NB Class 150"], mat=("SA 105","SA105"), qty=(1,1), balloon=True),
 dict(k="safety_valve",  draw=["SAFETY VALVE 50 NB"],                bom=["Safety relief valve, 50 NB", "Safety valve 50 NB"], mat=("SA 216 WCB","SA216 WCB"), qty=(1,1), balloon=True),
 dict(k="drain_valve",   draw=["DRAIN VALVE 25 NB"],                 bom=["Gate valve, drain, 25 NB", "Drain valve 25 NB"], mat=("SA 105","SA105"), qty=(2,2), balloon=True),
 dict(k="vent_valve",    draw=["VENT VALVE 15 NB"],                  bom=["Globe valve, vent, 15 NB", "Vent valve 15 NB"], mat=("SA 105","SA105"), qty=(1,1), balloon=True),
 dict(k="thermowell",    draw=["THERMOWELL 1/2 NPT"],                bom=["Thermowell, temperature, 1/2 NPT", "Thermowell 1/2 NPT"], mat=("SS 316","SS316"), qty=(2,2), balloon=True),
 dict(k="pr_tapping",    draw=["PRESSURE TAPPING 1/2 NB"],           bom=["Pressure gauge tapping, 1/2 NB", "Pressure tapping 1/2 NB"], mat=("SA 105","SA105"), qty=(2,2), balloon=True),
 dict(k="gasket",        draw=["GASKET 200 NB CL.150"],              bom=["Gasket, spiral wound, 200 NB, 150#", "Spiral wound gasket 200 NB Class 150"], mat=("SS 304 + GRAPHITE","SS304+Graphite"), qty=(2,2), balloon=False),
 dict(k="stud",          draw=["STUD BOLT M20 X 90 WITH 2 NUTS"],    bom=["Stud bolt with nuts, M20 x 90", "Stud bolt M20 x 90 with 2 nuts"], mat=("SA 193 B7","SA193 B7"), qty=(16,16), balloon=False),
 dict(k="lifting_lug",   draw=["LIFTING LUG 16 THK"],                bom=["Lug, lifting, 16 mm", "Lifting lug 16 thk"], mat=("IS 2062 GR.B","IS2062 Gr.B"), qty=(4,4), balloon=True),
 dict(k="casing_plate",  draw=["CASING PLATE 3 THK"],                bom=["Casing plate, 3 mm", "Casing plate 3 thk"], mat=("IS 2062","IS2062"), qty=(16,24), balloon=True),
 dict(k="exp_joint",     draw=["EXPANSION JOINT 600 X 600"],         bom=["Expansion joint, fabric type, 600 x 600", "Expansion joint 600 x 600"], mat=("FABRIC","Fabric"), qty=(1,1), balloon=True),
 dict(k="name_plate",    draw=["NAME PLATE"],                        bom=["Name plate, stainless steel", "Name plate"], mat=("SS 304","SS304"), qty=(1,1), balloon=True),
 dict(k="insp_door",     draw=["INSPECTION DOOR 300 X 300"],         bom=["Inspection door, 300 x 300", "Inspection door 300 x 300"], mat=("IS 2062","IS2062"), qty=(2,2), balloon=True),
]
# BOM-only (non-drawn) items. Implied by NOTES on the drawing (rule-based check, "L3").
NON_DRAWN = [
 dict(k="electrode",  rule="Welding consumables", bom=["Welding electrode E7018 3.15 mm", "Electrode E7018, 3.15 mm"], mat=("E7018","E7018"), uom="kg"),
 dict(k="insulation", rule="Insulation",          bom=["Mineral wool insulation 100 mm thk", "Insulation, mineral wool, 100 mm"], mat=("MINERAL WOOL","Mineral wool"), uom="m2"),
 dict(k="cladding",   rule="Cladding",            bom=["Aluminium cladding 0.6 mm", "Cladding sheet, aluminium, 0.6 mm"], mat=("ALUMINIUM","Aluminium"), uom="m2"),
 dict(k="primer",     rule="Paint / primer",      bom=["Red oxide primer", "Primer, red oxide, 1 coat"], mat=("-","-"), uom="ltr"),
]
EXTRA_POOL = [
 ("Soot blower, retractable", "Carbon steel", 1),
 ("Pressure reducing valve, 25 NB", "SA 216 WCB", 1),
 ("Hand rail, 42 NB", "IS 1239", 12),
 ("Spare tube 38.1 x 3.6, 6 m", "SA 210 GR.A1", 4),
]

DRAW_STYLES = [
 dict(cols=["Item", "Description", "Material", "Qty", "Part No", "Remarks"],   widths=[34,230,100,34,74,70], font="Helvetica"),
 dict(cols=["Sr No", "Part Description", "Material", "Size", "Qty"],           widths=[38,230,100,80,34],     font="Helvetica"),
 dict(cols=["Pos", "Description", "Material Spec", "Qty"],                      widths=[30,240,110,34],        font="Courier"),
]
BOM_STYLES = [
 dict(cols=["Item No", "Description", "Material", "Size", "Qty", "UoM", "Remarks"], keys=["id","desc","mat","size","qty","uom","rem"]),
 dict(cols=["Sr. No.", "Item Description", "Material Spec", "Qty", "Unit", "Drg Ref"], keys=["id","desc","mat","qty","uom","dref"]),
 dict(cols=["Pos", "Description", "Matl", "Nos"],                                    keys=["id","desc","mat","qty"]),
]

def qty_for(rng, item, per_bank_mult):
    lo, hi = item["qty"]
    q = rng.choice(range(lo, hi + 1, 2 if hi - lo >= 2 and item["k"] in ("tube","return_bend") else 1)) if hi > lo else lo
    return q

# ----------------------------------------------------------------------------------------
def build_sample(idx, seed, out_dir):
    rng = random.Random(seed)
    dstyle = rng.choice(DRAW_STYLES)
    bstyle = rng.choice(BOM_STYLES)
    renumber = rng.random() < 0.5          # BOM has its own numbering (no ID link)  -> tests fuzzy fallback
    drawing_rev = rng.choice(["B", "C", "D"])
    mult_note = rng.random() < 0.5         # "TUBE QTY IS PER BANK. TYPICAL FOR 4 BANKS."
    n_banks = 4

    # choose items for this drawing (14-19)
    items = [dict(it) for it in CAT]
    must = {"inlet_header","outlet_header","tube","inlet_flange","outlet_flange"}
    optional = [i for i in items if i["k"] not in must]
    drop = rng.sample(optional, k=rng.randint(0, 3))
    items = [i for i in items if i not in drop]
    for n, it in enumerate(items, start=1):
        it["id"] = str(n)
        it["draw_desc"] = rng.choice(it["draw"])
        it["draw_mat"] = it["mat"][0]
        it["bom_desc"] = rng.choice(it["bom"])
        it["bom_mat"] = it["mat"][1]
        it["qty_draw"] = qty_for(rng, it, mult_note)
        it["qty_bom"] = it["qty_draw"] * n_banks if (mult_note and it["k"] == "tube") else it["qty_draw"]

    # notes (drive the implied-item rules)
    notes = ["ALL DIMENSIONS ARE IN MM.", "ALL WELDING AS PER IBR / ASME SEC I."]
    nd = [d for d in NON_DRAWN if d["k"] == "electrode"]
    if rng.random() < 0.8:
        notes.append("INSULATE WITH 100 THK MINERAL WOOL" + (" AND CLAD WITH 0.6 THK ALUMINIUM SHEET." if rng.random() < 0.6 else "."))
    if rng.random() < 0.4:
        notes.append("PAINT: ONE COAT RED OXIDE PRIMER.")
    if mult_note:
        notes.append(f"TUBE QTY IS PER BANK. TYPICAL FOR {n_banks} BANKS.")
    text_notes = " ".join(notes).upper()
    nd_items = [d for d in NON_DRAWN if (d["k"] == "electrode") or
                (d["k"] == "insulation" and "INSULAT" in text_notes) or
                (d["k"] == "cladding" and "CLAD" in text_notes) or
                (d["k"] == "primer" and "PRIMER" in text_notes)]
    for d in nd_items:
        d["bom_desc_chosen"] = rng.choice(d["bom"])
        d["qty_bom"] = {"kg": rng.choice([60, 96, 120]), "m2": rng.choice([180, 220, 260]), "ltr": rng.choice([20, 30])}[d["uom"]]

    # ------------------------------------------------------------------ draw the drawing PDF
    pdf_path = os.path.join(out_dir, f"sample_{idx:02d}_drawing.pdf")
    W, H = landscape(A3)
    c = canvas.Canvas(pdf_path, pagesize=(W, H))
    c.setTitle(f"Economiser assembly {idx:02d}")
    c.setLineWidth(1.2); c.rect(20, 20, W - 40, H - 40)

    # schematic (stylised economiser) in the left ~60% of the sheet
    c.setLineWidth(1.0)
    x0, x1 = 90, 520
    c.rect(x0, 520, x1 - x0, 24)            # top header
    c.rect(x0, 300, x1 - x0, 24)            # bottom header
    for k in range(16):                     # tubes
        x = x0 + 20 + k * 31
        c.line(x, 324, x, 520)
    c.setDash(4, 3); c.rect(x0 - 25, 280, x1 - x0 + 50, 290); c.setDash()
    c.rect(x0 - 40, 522, 14, 20); c.rect(x1 + 26, 302, 14, 20)       # flanges
    c.circle(x0 + 120, 590, 0.1); 
    # balloons: placed around the schematic with leader lines
    balloon_slots = [(x,610) for x in range(100,540,70)] + [(x,250) for x in range(100,540,70)] \
                    + [(50,330),(50,420),(50,500)] + [(572,y) for y in (560,500,440,380,320)]
    rng.shuffle(balloon_slots)
    bi = 0
    c.setFont("Helvetica", 8)
    for it in items:
        if not it["balloon"] or bi >= len(balloon_slots): it["has_balloon"] = False; continue
        bx, by = balloon_slots[bi]; bi += 1
        tx = min(max(bx, x0 + 5), x1 - 5) + rng.randint(-12, 12)
        ty = min(max(by, 305), 540)
        c.line(bx, by, tx, ty)
        c.setFillColorRGB(1, 1, 1); c.circle(bx, by, 9, stroke=1, fill=1); c.setFillColorRGB(0, 0, 0)
        c.drawCentredString(bx, by - 3, it["id"])
        it["has_balloon"] = True

    # parts table (top-right or lower-right)
    cols, widths = dstyle["cols"], dstyle["widths"]
    tw = sum(widths); row_h = 15
    tx0 = W - 40 - tw
    ty0 = H - 60 if rng.random() < 0.6 else 60 + row_h * (len(items) + 1) + 180
    font = dstyle["font"]
    c.setFont(font + ("-Bold" if font == "Helvetica" else "-Bold"), 8)
    y = ty0
    c.setFillColorRGB(0.88, 0.88, 0.88); c.rect(tx0, y - row_h, tw, row_h, stroke=0, fill=1); c.setFillColorRGB(0, 0, 0)
    xx = tx0
    for name, w in zip(cols, widths):
        c.drawString(xx + 3, y - 11, name); xx += w
    c.setFont(font, 8)
    y -= row_h
    colkeys = {"Item":"id","Sr No":"id","Pos":"id","Description":"desc","Part Description":"desc","Material":"mat","Material Spec":"mat",
               "Size":"size","Qty":"qty","Part No":"pno","Remarks":"rem"}
    for it in items:
        it["pno"] = f"EC-{idx:02d}-{it['id'].zfill(2)}"
        it["size"] = ""
        vals = {"id": it["id"], "desc": it["draw_desc"], "mat": it["draw_mat"], "size": it["size"], "qty": str(it["qty_draw"]), "pno": it["pno"],
                "rem": "TYP." if (it["k"] == "tube" and mult_note) else ""}
        xx = tx0
        for name, w in zip(cols, widths):
            c.drawString(xx + 3, y - 11, vals[colkeys[name]]); xx += w
        y -= row_h
    # grid lines
    nrows = len(items) + 1
    for r in range(nrows + 1):
        c.line(tx0, ty0 - r * row_h, tx0 + tw, ty0 - r * row_h)
    xx = tx0
    for w in widths + [0]:
        c.line(xx, ty0, xx, ty0 - nrows * row_h); xx += w

    # notes (left-bottom)
    c.setFont("Helvetica-Bold", 9); c.drawString(40, 200, "NOTES:")
    c.setFont("Helvetica", 8)
    for n, t in enumerate(notes, start=1):
        c.drawString(40, 200 - 14 * n, f"{n}. {t}")

    # title block
    c.rect(W - 360, 20, 320, 120)
    c.setFont("Helvetica-Bold", 11); c.drawString(W - 350, 118, "ECONOMISER ASSEMBLY")
    c.setFont("Helvetica", 9)
    c.drawString(W - 350, 96, f"DRG NO : EC-{idx:02d}-GA-001")
    c.drawString(W - 350, 80, f"DRG REV : {drawing_rev}")
    c.drawString(W - 350, 64, "SCALE : NTS    SIZE : A3")
    c.drawString(W - 350, 48, "SAMPLE (SYNTHETIC) DRAWING")
    c.save()

    # ------------------------------------------------------------------ build BOM with injected errors
    bom_rows = []
    for it in items:
        bom_rows.append(dict(src=it["k"], draw_id=it["id"], desc=it["bom_desc"], mat=it["bom_mat"], qty=it["qty_bom"], uom="Nos",
                             size="", rem="", dref=f"EC-{idx:02d}-GA-001"))
    for d in nd_items:
        bom_rows.append(dict(src=d["k"], draw_id=None, desc=d["bom_desc_chosen"], mat=d["mat"][1], qty=d["qty_bom"], uom=d["uom"],
                             size="", rem="consumable", dref=f"EC-{idx:02d}-GA-001"))

    truth = []
    hotspots = ["thermowell","pr_tapping","gasket","stud","name_plate","insp_door","lifting_lug","vent_valve","drain_valve","casing_plate"]
    n_err = rng.randint(4, 6)
    kinds = ["omission", "omission", "quantity", "spec", "extra", "duplicate", "implied", "revision"]
    rng.shuffle(kinds)
    chosen = kinds[:n_err]
    if mult_note and rng.random() < 0.7: chosen.append("multiplier")
    drawn_keys = {i["k"] for i in items}
    touched = set()
    bom_rev = drawing_rev

    for kind in chosen:
        if kind == "omission":
            cand = [r for r in bom_rows if r["src"] in hotspots and r["draw_id"] and r["src"] not in touched]
            if not cand: continue
            r = rng.choice(cand); bom_rows.remove(r); touched.add(r["src"])
            truth.append(dict(type="omission", key=r["draw_id"], desc=r["desc"]))
        elif kind == "quantity":
            cand = [r for r in bom_rows if r["draw_id"] and r["src"] not in touched and r["src"] != "tube"]
            if not cand: continue
            r = rng.choice(cand); old = r["qty"]; r["qty"] = max(1, old - rng.choice([1, 2, 4])) if old > 2 else old + 1; touched.add(r["src"])
            truth.append(dict(type="quantity", key=r["draw_id"], desc=r["desc"], drawing_qty=old, bom_qty=r["qty"]))
        elif kind == "spec":
            cand = [r for r in bom_rows if r["draw_id"] and r["src"] in ("tube","support_plate","casing_plate","inlet_flange","outlet_flange","lifting_lug","drain_valve","vent_valve","inlet_header","outlet_header") and r["src"] not in touched]
            if not cand: continue
            r = rng.choice(cand); touched.add(r["src"])
            if r["src"] == "tube":
                r["desc"] = r["desc"].replace("3.6", "4.0")
            elif r["src"] in ("support_plate","casing_plate","lifting_lug"):
                r["desc"] = r["desc"].replace("12", "10").replace("16", "12").replace("3 ", "5 ").replace("3 mm", "5 mm").replace("3 thk", "5 thk")
            elif r["src"] in ("inlet_header","outlet_header"):
                r["desc"] = r["desc"].replace("8.0", "10.0").replace(" 8 ", " 10 ")
            else:
                r["desc"] = r["desc"].replace("25 NB", "40 NB").replace("15 NB", "25 NB").replace("200 NB", "250 NB")
            if r["desc"] == next(i["bom_desc"] for i in items if i["id"] == r["draw_id"]):
                r["mat"] = "SA 335 P11"           # fall back to a material change
            truth.append(dict(type="spec", key=r["draw_id"], desc=r["desc"]))
        elif kind == "extra":
            e = rng.choice(EXTRA_POOL)
            bom_rows.append(dict(src="extra", draw_id=None, desc=e[0], mat=e[1], qty=e[2], uom="Nos", size="", rem="", dref=""))
            truth.append(dict(type="extra", key=e[0], desc=e[0]))
        elif kind == "duplicate":
            cand = [r for r in bom_rows if r["draw_id"] and r["src"] not in touched]
            r = rng.choice(cand); touched.add(r["src"])
            alt = [i for i in items if i["k"] == r["src"]][0]
            other = [d for d in alt["bom"] if d != r["desc"]]
            dup = dict(r); dup["desc"] = other[0] if other else r["desc"] + " (alt)"; dup["draw_id"] = None; dup["src"] = "dup"
            bom_rows.append(dup)
            truth.append(dict(type="duplicate", key=dup["desc"], desc=dup["desc"], item=r["draw_id"]))
        elif kind == "implied":
            cand = [r for r in bom_rows if r["src"] in ("electrode","insulation","cladding","primer")]
            if not cand: continue
            r = rng.choice(cand); bom_rows.remove(r)
            rule = next(d["rule"] for d in NON_DRAWN if d["k"] == r["src"])
            truth.append(dict(type="implied", key=rule, desc=r["desc"]))
        elif kind == "revision":
            bom_rev = chr(ord(drawing_rev) - 1)
            truth.append(dict(type="revision", key="revision", desc=f"drawing {drawing_rev} vs BOM {bom_rev}"))
        elif kind == "multiplier":
            r = next((r for r in bom_rows if r["src"] == "tube"), None)
            if r and "tube" not in touched:
                tdraw = next(i for i in items if i["k"] == "tube")["qty_draw"]
                r["qty"] = tdraw; touched.add("tube")
                truth.append(dict(type="multiplier", key=tdraw and next(i for i in items if i["k"] == "tube")["id"], desc=r["desc"], drawing_qty=tdraw, expected=tdraw * n_banks))

    # presentation order / numbering of the BOM
    if renumber:
        rng.shuffle(bom_rows) if rng.random() < 0.4 else None
        for n, r in enumerate(bom_rows, start=1): r["bom_id"] = str(n)
    else:
        nxt = max(int(i["id"]) for i in items) + 1
        for r in bom_rows:
            if r["draw_id"]: r["bom_id"] = r["draw_id"]
            else: r["bom_id"] = str(nxt); nxt += 1
        bom_rows.sort(key=lambda r: int(r["bom_id"]))

    wb = Workbook(); ws = wb.active; ws.title = "BOM"
    ws["A1"] = "BILL OF MATERIALS - ECONOMISER ASSEMBLY (SYNTHETIC SAMPLE)"; ws["A1"].font = Font(name="Arial", bold=True, size=12)
    ws["A2"] = f"BOM Rev: {bom_rev}"; ws["A3"] = f"Drawing ref: EC-{idx:02d}-GA-001"
    hdr_row = 5
    for j, name in enumerate(bstyle["cols"], start=1):
        cell = ws.cell(hdr_row, j, name); cell.font = Font(name="Arial", bold=True)
    for i, r in enumerate(bom_rows, start=hdr_row + 1):
        vals = {"id": r["bom_id"], "desc": r["desc"], "mat": r["mat"], "size": r["size"], "qty": r["qty"], "uom": r["uom"], "rem": r["rem"], "dref": r["dref"]}
        for j, key in enumerate(bstyle["keys"], start=1):
            ws.cell(i, j, vals[key]).font = Font(name="Arial")
    for j, w in enumerate([10, 48, 20, 10, 10, 12, 18], start=1): ws.column_dimensions[chr(64 + j)].width = w
    bom_path = os.path.join(out_dir, f"sample_{idx:02d}_bom.xlsx"); wb.save(bom_path)

    # ground truth keys refer to DRAWING item ids for omission/quantity/spec/multiplier
    with open(os.path.join(out_dir, f"sample_{idx:02d}_truth.json"), "w") as f:
        json.dump(dict(sample=idx, drawing_rev=drawing_rev, bom_rev=bom_rev, renumbered=renumber, style_draw=dstyle["cols"], style_bom=bstyle["cols"],
                       n_items=len(items), n_bom_rows=len(bom_rows), errors=truth,
                       drawing_items=[dict(id=i["id"], desc=i["draw_desc"], qty=i["qty_draw"]) for i in items]), f, indent=2)
    return pdf_path


def make_scan(pdf_path, out_path, seed=0, dpi=200):
    """Rasterise a vector PDF and add scan-like degradation: slight skew, blur, noise, JPEG artefacts."""
    import fitz, numpy as np, cv2
    rng = np.random.default_rng(seed)
    doc = fitz.open(pdf_path); page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), colorspace=fitz.csGRAY)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w).copy()
    ang = rng.uniform(-0.3, 0.3)
    M = cv2.getRotationMatrix2D((pix.w / 2, pix.h / 2), ang, 1.0)
    img = cv2.warpAffine(img, M, (pix.w, pix.h), borderValue=255)
    img = cv2.GaussianBlur(img, (3, 3), 0.8)
    noise = rng.normal(0, 6, img.shape)
    img = np.clip(img + noise, 0, 255).astype(np.uint8)
    ok, enc = cv2.imencode(".jpg", img, [int(cv2.IMWRITE_JPEG_QUALITY), 70])
    out = fitz.open(); p = out.new_page(width=page.rect.width, height=page.rect.height)
    p.insert_image(p.rect, stream=enc.tobytes())
    out.save(out_path)


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 15
    out_dir = sys.argv[2] if len(sys.argv) > 2 else HERE
    seed0 = int(sys.argv[3]) if len(sys.argv) > 3 else 1000
    os.makedirs(out_dir, exist_ok=True)
    for i in range(1, n + 1):
        p = build_sample(i, seed=seed0 + i, out_dir=out_dir)
        if i % 3 == 0:
            make_scan(p, p.replace("_drawing.pdf", "_drawing_scan.pdf"), seed=i)
    print(f"generated {n} samples in {out_dir}")
