"""
N (Normalise) + M (Match) + C (Compare) stages.

Issue types (O-Q-S-E-R-D + L3):
  omission   drawing item has no BOM line                         RED
  quantity   matched, quantities differ                           ORANGE
  spec       matched, dimension / material conflict               ORANGE
  multiplier note says "per bank x N" but BOM qty is not N x      ORANGE
  revision   drawing revision != BOM revision                     ORANGE
  extra      BOM line with no drawing evidence                    YELLOW
  duplicate  BOM line repeats an already-matched drawing item     YELLOW
  confirm    probable match (score between 'possible' and 'match') YELLOW
  implied    L3 rule: drawing note implies an item the BOM lacks  BLUE  (suggestion, not a hard error)
"""
import re
from rapidfuzz import fuzz
from bomcheck.extract import valid_quantity, validate_table

NUM_RE = re.compile(r"\d+(?:\.\d+)?(?:/\d+)?")
TOK_RE = re.compile(r"\d+(?:\.\d+)?(?:/\d+)?|[a-z]+")


# ------------------------------------------------------------------------------ normalisation
def _canon_num(s):
    if "/" in s:
        return s
    try:
        return "%g" % float(s)
    except ValueError:
        return s


def tokens(text, cfg):
    """-> (alpha_tokens list, number_tokens set)."""
    t = str(text or "").lower()
    for k, v in cfg.get("phrase_synonyms", {}).items():
        t = t.replace(k, v)
    t = t.replace("ø", " ").replace("⌀", " ").replace("×", " ").replace("#", " ")
    alpha, nums = [], set()
    for tok in TOK_RE.findall(t):
        if tok[0].isdigit():
            nums.add(_canon_num(tok))
        else:
            tok = cfg["abbreviations"].get(tok, tok)
            for p in tok.split():
                if p and p not in ("x",):
                    alpha.append(p[:-1] if len(p) > 4 and p.endswith("s") else p)   # crude plural strip
    return alpha, nums


def compact_material(s):
    s = re.sub(r"[^a-z0-9]", "", str(s or "").lower())
    return "" if s in ("", "na") else s


def parse_qty(v):
    if v is None or not valid_quantity(v):
        return None
    m = re.match(r"\s*([\d,]+(?:\.\d+)?)", str(v))
    return float(m.group()) if m else None


TOKEN_FUZZY = 88          # overwritten from config.json (thresholds.token_fuzzy) in compare()


def _tok_eq(a, b):
    return a == b or (min(len(a), len(b)) >= 5 and fuzz.ratio(a, b) >= TOKEN_FUZZY)     # tolerate OCR / typing typos


def dice(A, B):
    if not A and not B:
        return 100.0
    if not A or not B:
        return 0.0
    used, hit = set(), 0
    for a in A:
        for j, b in enumerate(B):
            if j not in used and _tok_eq(a, b):
                used.add(j); hit += 1; break
    dice_ = 100.0 * 2 * hit / (len(A) + len(B))
    overlap = 100.0 * hit / min(len(A), len(B))      # BOM wording is often longer ("gasket, spiral wound") than the drawing's
    return 0.5 * dice_ + 0.5 * overlap


def num_sim(a, b):
    if not a and not b:
        return 100.0
    if not a or not b:
        return 40.0
    if a == b:
        return 100.0
    return max(40.0, 100.0 * len(a & b) / len(a | b))


# ------------------------------------------------------------------------------ prepared items
def prep_d(d, cfg):
    al, nu = tokens(d["description"] + " " + d.get("size", ""), cfg)
    d["_al"], d["_nu"] = al, nu
    d["_mat"] = compact_material(d.get("material"))
    d["_qty"] = parse_qty(d.get("qty"))
    return d


def prep_b(b, cfg):
    al, nu = tokens(b["description"] + " " + b.get("size", ""), cfg)
    b["_al"], b["_nu"] = al, nu
    b["_mat"] = compact_material(b.get("material"))
    b["_qty"] = parse_qty(b.get("qty"))
    return b


def score(d, b, cfg):
    w = cfg["weights"]
    s = w["text"] * dice(d["_al"], b["_al"]) + w["numbers"] * num_sim(d["_nu"], b["_nu"])
    if d["id"] and d["id"] == b["id"]:
        s += cfg["thresholds"]["id_bonus"]
    dp, bp = re.sub(r"\W", "", d.get("partno", "").lower()), re.sub(r"\W", "", b.get("partno", "").lower())
    if dp and dp == bp:
        s = max(s, 98.0)
    return min(100.0, s)


# ------------------------------------------------------------------------------ main compare
def compare(drawing, bom, cfg):
    global TOKEN_FUZZY
    th = cfg["thresholds"]
    TOKEN_FUZZY = th.get("token_fuzzy", 88)
    non_drawn = re.compile(cfg["non_drawn_patterns"], re.I)
    issues, matches = [], []

    # ---- drawing items: table rows (primary), plus balloon-only ids (L1)
    ditems = []
    for p in drawing["pages"]:
        if p["table"]:
            for r in p["table"]["rows"]:
                ditems.append(prep_d(dict(r, page=p["page"], source="table", has_balloon=False), cfg))
    table_ids = {d["id"] for d in ditems}
    balloon_by_id = {}
    for p in drawing["pages"]:
        for bl in p["balloons"]:
            balloon_by_id.setdefault(bl["id"], dict(bl, page=p["page"]))
    for d in ditems:
        d["has_balloon"] = d["id"] in balloon_by_id
    balloon_only = [bid for bid in balloon_by_id if bid not in table_ids]

    brows = [prep_b(dict(r), cfg) for r in bom["rows"]]
    cand = [b for b in brows if not non_drawn.search(b["description"] + " " + b["remarks"])]
    nd_rows = [b for b in brows if b not in cand]

    extraction_errors = list(drawing.get("extraction_errors", []))
    extraction_warnings = list(drawing.get("extraction_warnings", []))
    for p in drawing["pages"]:
        if p.get("table"):
            table_errors, table_warnings = validate_table(p["table"])
            extraction_errors.extend(f"Page {p['page'] + 1}: {message}" for message in table_errors)
            extraction_warnings.extend(f"Page {p['page'] + 1}: {message}" for message in table_warnings)
    bom_columns = set(bom.get("colidx", {}).values())
    if not {"description", "qty"} <= bom_columns:
        extraction_errors.append("The BOM is missing a recognized Description or Qty column.")
    for row in brows:
        if not re.search(r"[A-Za-z]{2,}", row["description"]):
            extraction_errors.append(f"BOM row {row['row']} has no meaningful description.")
        if "qty" in bom_columns and not valid_quantity(row.get("qty")):
            extraction_errors.append(f"BOM row {row['row']} has an invalid or unreadable quantity.")
    if not ditems and not balloon_by_id:
        extraction_errors.append("No usable drawing parts table or numbered balloons were extracted.")

    if extraction_errors:
        reasons = list(dict.fromkeys(extraction_errors + extraction_warnings))
        reasons.append("Comparison skipped because extraction validation failed; no omission or extra findings were generated.")
        return dict(
            reliability="LOW",
            reliability_reasons=reasons,
            mean_ocr_conf=round(
                sum(d.get("conf", 100.0) for d in ditems) / len(ditems), 1
            ) if ditems else 0.0,
            linked_fraction=0.0,
            issues=[],
            matches=[],
            ditems=ditems,
            bom_rows=brows,
            nondrawn_rows=[b["row"] for b in nd_rows],
            completeness=None,
            n_drawing_items=len(ditems),
            n_bom_rows=len(brows),
            n_balloons=len(balloon_by_id),
            balloon_only=balloon_only,
            extraction_errors=list(dict.fromkeys(extraction_errors)),
            extraction_warnings=list(dict.fromkeys(extraction_warnings)),
            comparison_skipped=True,
        )

    # ---- one-to-one assignment, best score first
    pairs = sorted(((score(d, b, cfg), i, j) for i, d in enumerate(ditems) for j, b in enumerate(cand)), key=lambda t: -t[0])
    d_to_b, b_used = {}, set()
    for sc, i, j in pairs:
        if sc < th["possible"] or i in d_to_b or j in b_used:
            continue
        d_to_b[i] = (j, sc); b_used.add(j)

    mult = cfg.get("multiplier") or {}
    all_text = drawing["text"].upper()
    n_mult = None
    if mult and re.search(mult["note_regex"], all_text):
        m = re.search(mult["count_regex"], all_text)
        n_mult = int(m.group(1)) if m else None

    def add(typ, sev, msg, d=None, b=None, extra=None):
        iss = dict(type=typ, severity=sev, message=msg, drawing_id=d["id"] if d else None, drawing_desc=d["description"] if d else None,
                   page=d["page"] if d else None, bbox=d["bbox"] if d and "bbox" in d else None,
                   bom_row=b["row"] if b else None, bom_desc=b["description"] if b else None)
        if d is not None and d.get("conf", 100) < 60:
            iss["message"] += "  [low OCR confidence - verify on sheet]"
            iss["low_conf"] = True
        if extra:
            iss.update(extra)
        issues.append(iss)

    for i, d in enumerate(ditems):
        item_label = f"Item {d['id']}" if d.get("id") else "Drawing component"
        if i not in d_to_b:
            add("omission", "red", f"{item_label} '{d['description']}' (qty {d.get('qty','?')}) is on the drawing but has no BOM line", d=d)
            matches.append(dict(drawing_id=d["id"], bom_row=None, score=0, status="MISSING"))
            continue
        j, sc = d_to_b[i]
        b = cand[j]
        status, flagged = "OK", False
        # quantity (with multiplier note)
        applies = n_mult and re.search(mult["applies_to"], d["description"].lower())
        if d["_qty"] is not None and b["_qty"] is not None:
            if applies:
                exp = d["_qty"] * n_mult
                if abs(b["_qty"] - exp) > 1e-9:
                    if abs(b["_qty"] - d["_qty"]) < 1e-9:
                        add("multiplier", "orange", f"{item_label}: drawing note says qty is per bank, x{n_mult}; BOM has {b['_qty']:g} but total should be {exp:g}", d=d, b=b,
                            extra=dict(drawing_qty=d["_qty"], bom_qty=b["_qty"], expected=exp))
                    else:
                        add("quantity", "orange", f"{item_label}: quantity differs (drawing {d['_qty']:g} per bank x{n_mult} = {exp:g}, BOM {b['_qty']:g})", d=d, b=b)
                    flagged = True; status = "QTY"
            elif abs(d["_qty"] - b["_qty"]) > 1e-9:
                add("quantity", "orange", f"{item_label}: quantity differs (drawing {d['_qty']:g}, BOM {b['_qty']:g})", d=d, b=b,
                    extra=dict(drawing_qty=d["_qty"], bom_qty=b["_qty"]))
                flagged = True; status = "QTY"
        # spec: real numeric conflict (both sides carry a value the other lacks) or material mismatch
        d_only, b_only = d["_nu"] - b["_nu"], b["_nu"] - d["_nu"]
        spec_msgs = []
        if d_only and b_only:
            spec_msgs.append(f"dimension: drawing {sorted(d_only)} vs BOM {sorted(b_only)}")
        if d["_mat"] and b["_mat"] and fuzz.ratio(d["_mat"], b["_mat"]) < th["material_ratio"]:
            spec_msgs.append(f"material: drawing '{d.get('material')}' vs BOM '{b.get('material')}'")
        if spec_msgs:
            add("spec", "orange", f"{item_label}: specification mismatch - " + "; ".join(spec_msgs), d=d, b=b)
            flagged = True; status = "SPEC" if status == "OK" else status + "+SPEC"
        # probable-match band
        if sc < th["match"] and not flagged:
            add("confirm", "yellow", f"{item_label}: probable match to BOM row {b['row']} (score {sc:.0f}) - please confirm", d=d, b=b)
            status = "CONFIRM"
        matches.append(dict(drawing_id=d["id"], bom_row=b["row"], score=round(sc, 1), status=status))

    # ---- BOM rows with no drawing item: duplicate or extra
    for j, b in enumerate(cand):
        if j in b_used:
            continue
        best = max(((score(d, b, cfg), i) for i, d in enumerate(ditems)), default=(0, None))
        if best[1] is not None and best[0] >= th.get("duplicate", th["match"]) and best[1] in d_to_b:
            d = ditems[best[1]]
            add("duplicate", "yellow", f"BOM row {b['row']} '{b['description']}' repeats item {d['id']} (already covered by BOM row {cand[d_to_b[best[1]][0]]['row']})", d=d, b=b)
        else:
            add("extra", "yellow", f"BOM row {b['row']} '{b['description']}' has no matching item on the drawing (extra, or from another revision)", b=b)

    # ---- balloons that are neither in the table nor matchable (L1)
    for bid in balloon_only:
        bl = balloon_by_id[bid]
        brow = next((b for b in cand if b["id"] == bid and cand.index(b) not in b_used), None)
        if brow is None:
            issues.append(dict(type="omission", severity="red", message=f"Balloon {bid} on the drawing is in neither the parts table nor the BOM", drawing_id=bid,
                               drawing_desc="(balloon only)", page=bl["page"], bbox=bl["bbox"], bom_row=None, bom_desc=None))

    # ---- revision
    if drawing["revision"] and bom["revision"] and drawing["revision"].upper() != bom["revision"].upper():
        issues.append(dict(type="revision", severity="orange", drawing_id=None, drawing_desc=None, page=None, bbox=None, bom_row=None, bom_desc=None,
                           message=f"Revision mismatch: drawing is Rev {drawing['revision']}, BOM is Rev {bom['revision']}"))

    # ---- L3 implied-item rules (suggestions)
    bom_text = " ".join(b["description"] + " " + b["remarks"] for b in brows)
    for rule in cfg.get("implied_rules", []):
        if re.search(rule["when"], all_text) and not re.search(rule["expect"], bom_text, re.I):
            loc = None
            for p in drawing["pages"]:
                hits = [w for w in p["words"] if re.search(rule["when"], w["text"].upper())]
                for w in hits:
                    if not p["table"] or not (p["table"]["rows"][0]["bbox"][0] <= w["xc"]):    # prefer a note, not the table
                        loc = (p["page"], (w["x0"] - 2, w["y0"] - 2, w["x1"] + 2, w["y1"] + 2)); break
                if loc:
                    break
            issues.append(dict(type="implied", severity="blue", drawing_id=None, drawing_desc=None, page=loc[0] if loc else None, bbox=loc[1] if loc else None,
                               bom_row=None, bom_desc=None, label=rule["label"],
                               message=f"Suggestion ({rule['label']}): the drawing notes imply it, but no matching line was found in the BOM"))

    n_d = len(ditems)
    n_missing = sum(1 for m in matches if m["status"] == "MISSING")
    # ---- reliability gate: say "do not trust this run" instead of flooding the user with false alarms
    rel = cfg.get("reliability", {})
    confs = [d.get("conf", 100.0) for d in ditems]
    mean_conf = sum(confs) / len(confs) if confs else 0.0
    frac = (len(b_used) / len(cand)) if cand else 0.0
    reasons = list(dict.fromkeys(extraction_warnings))
    if n_d < rel.get("min_rows", 3):
        reasons.append(f"only {n_d} parts-table rows were read")
    if frac < rel.get("min_match_fraction", 0.5):
        reasons.append(f"only {100*frac:.0f}% of BOM lines could be linked to a drawing item (wording/extraction problem?)")
    if drawing["scanned"] and mean_conf < rel.get("min_ocr_conf", 70):
        reasons.append(f"low OCR confidence ({mean_conf:.0f}/100)")
    if reasons:
        reliability = "LOW"
    elif frac < rel.get("warn_match_fraction", 0.75) or (drawing["scanned"] and mean_conf < 80):
        reliability = "MEDIUM"; reasons.append("some lines are only weakly linked, or OCR confidence is moderate - review flags with care")
    else:
        reliability = "HIGH"
    return dict(reliability=reliability, reliability_reasons=reasons, mean_ocr_conf=round(mean_conf, 1), linked_fraction=round(frac, 3),issues=issues, matches=matches, ditems=ditems, bom_rows=brows, nondrawn_rows=[b["row"] for b in nd_rows],
                completeness=(100.0 * (n_d - n_missing) / n_d) if n_d else None, n_drawing_items=n_d, n_bom_rows=len(brows),
                n_balloons=len(balloon_by_id), balloon_only=balloon_only,
                extraction_errors=[], extraction_warnings=list(dict.fromkeys(extraction_warnings)),
                comparison_skipped=False)
