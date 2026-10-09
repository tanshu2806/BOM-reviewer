"""
E (Extract) stage of the pipeline.

Drawing  -> per page: words, parts-table rows (with bounding boxes), balloons, full text, revision
BOM      -> rows (with Excel row numbers), revision

Design notes
  * ONE table parser works on word boxes, so it serves both vector PDFs (words from the PDF text layer)
    and scanned PDFs (words from OCR).  The matcher never knows which one it was.
  * Column layout is read from the header row (header synonyms live in config.json), not hard-coded.
  * Assumption (config/ASSUMPTION D9): data cells are left-aligned under their header text.
"""
import json, os, re, statistics
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
from rapidfuzz import fuzz
import warnings
warnings.filterwarnings("ignore")
import pymupdf as fitz

HERE = os.path.dirname(os.path.abspath(__file__))
ITEM_RE = re.compile(r"^[A-Za-z]{0,2}[-._]?[0-9]{1,4}(?:[-._][0-9A-Za-z]{1,4})?[A-Za-z]?$")
REV_RES = [re.compile(r"\bDRG\.?\s*REV(?:ISION)?\.?\s*[:\-]?\s*([A-Z0-9]{1,2})\b"),
           re.compile(r"\bREV(?:ISION)?\.?\s*[:\-]?\s*([A-Z]|\d{1,2})\b")]


def load_config(path=None):
    with open(path or os.path.join(HERE, "config.json")) as f:
        return json.load(f)


def _norm_hdr(s):
    return re.sub(r"\s+", " ", re.sub(r"[.:,;()]", " ", str(s).lower().replace("'", "'"))).strip()


def _syn_list(cfg):
    """[(token_tuple, canon)] longest first."""
    out = []
    for canon, syns in cfg["header_synonyms"].items():
        for s in syns:
            out.append((tuple(_norm_hdr(s).split()), canon))
    out.sort(key=lambda t: -len(t[0]))
    return out


def find_revision(text):
    t = text.upper()
    for rx in REV_RES:
        m = rx.search(t)
        if m:
            return m.group(1)
    return None


# ------------------------------------------------------------------------------ words
def _mk_word(x0, y0, x1, y1, text, conf=100.0):
    return dict(x0=x0, y0=y0, x1=x1, y1=y1, text=text, conf=conf, xc=(x0 + x1) / 2, yc=(y0 + y1) / 2)


def vector_words(page):
    return [_mk_word(w[0], w[1], w[2], w[3], w[4]) for w in page.get_text("words") if w[4].strip()]


class Frame:
    """Maps pixel coordinates of the PRE-PROCESSED (deskewed) image back to the ORIGINAL page, in PDF points,
    so highlights drawn on the original scan land on the right spot."""
    def __init__(self, z, Minv=None):
        self.z, self.Minv = z, Minv

    def pt(self, x, y):
        if self.Minv is not None:
            x, y = self.Minv[0, 0] * x + self.Minv[0, 1] * y + self.Minv[0, 2], self.Minv[1, 0] * x + self.Minv[1, 1] * y + self.Minv[1, 2]
        return x / self.z, y / self.z

    def box(self, x0, y0, x1, y1):
        pts = [self.pt(x, y) for x, y in ((x0, y0), (x1, y0), (x0, y1), (x1, y1))]
        return min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts)


def preprocess_scan(img, z):
    """Flatten uneven illumination, estimate skew from long near-horizontal rules, rotate to straighten."""
    import cv2, numpy as np
    bg = cv2.GaussianBlur(cv2.dilate(img, np.ones((15, 15), np.uint8)), (0, 0), 25)
    flat = cv2.divide(img, bg, scale=255)
    bw = cv2.adaptiveThreshold(flat, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 15)
    W = img.shape[1]
    lines = cv2.HoughLinesP(bw, 1, np.pi / 1440, threshold=150, minLineLength=int(0.08 * W), maxLineGap=6)
    angle = 0.0
    if lines is not None:
        angs, wts = [], []
        for x1, y1, x2, y2 in lines.reshape(-1, 4):
            a = np.degrees(np.arctan2(y2 - y1, x2 - x1))
            if abs(a) <= 5:
                angs.append(a); wts.append(np.hypot(x2 - x1, y2 - y1))
        if angs:
            o = np.argsort(angs); cs = np.cumsum(np.array(wts)[o])
            angle = float(np.array(angs)[o][np.searchsorted(cs, cs[-1] / 2)])            # length-weighted median
    if abs(angle) < 0.05:
        return flat, Frame(z, None), 0.0
    h, w = flat.shape
    M = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    out = cv2.warpAffine(flat, M, (w, h), borderValue=255)
    return out, Frame(z, cv2.invertAffineTransform(M)), angle


def render_gray(page, dpi):
    import numpy as np
    z = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(z, z), colorspace=fitz.csGRAY)
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w).copy(), z


_TESSERACT_CMD = None
_EASYOCR_READER = None


def _get_tesseract():
    global _TESSERACT_CMD
    if _TESSERACT_CMD is not None:
        return _TESSERACT_CMD
    import shutil
    cmd = shutil.which("tesseract")
    if cmd:
        _TESSERACT_CMD = cmd
        return cmd
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        r"D:\Program Files\Tesseract-OCR\tesseract.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            try:
                import pytesseract
                pytesseract.pytesseract.tesseract_cmd = c
                _TESSERACT_CMD = c
                return c
            except Exception:
                pass
    return None


def _get_easyocr():
    global _EASYOCR_READER
    if _EASYOCR_READER is not None:
        return _EASYOCR_READER if _EASYOCR_READER is not False else None
    try:
        import easyocr
        _EASYOCR_READER = easyocr.Reader(['en'], gpu=False)
        return _EASYOCR_READER
    except Exception:
        _EASYOCR_READER = False
        return None


def ocr_words(img, frame, cfg):
    tess = _get_tesseract()
    if tess:
        try:
            import pytesseract
            from pytesseract import Output
            pytesseract.pytesseract.tesseract_cmd = tess
            d = pytesseract.image_to_data(img, config=f"--psm {cfg['ocr']['psm']}", output_type=Output.DICT)
            words = []
            for i, t in enumerate(d["text"]):
                t = t.strip()
                if not t:
                    continue
                conf = float(d["conf"][i])
                if conf < cfg["thresholds"]["ocr_min_conf"]:
                    continue
                x, y, w, h = d["left"][i], d["top"][i], d["width"][i], d["height"][i]
                words.append(_mk_word(*frame.box(x, y, x + w, y + h), t, conf))
            if words:
                return words
        except Exception:
            pass

    reader = _get_easyocr()
    if reader:
        try:
            results = reader.readtext(img)
            words = []
            for bbox, text, conf in results:
                t = text.strip()
                c = float(conf * 100 if conf <= 1.0 else conf)
                if not t or c < cfg["thresholds"]["ocr_min_conf"]:
                    continue
                xs = [p[0] for p in bbox]
                ys = [p[1] for p in bbox]
                x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
                words.append(_mk_word(*frame.box(x0, y0, x1, y1), t, c))
            return words
        except Exception:
            pass

    return []


def find_table_region(img, z):
    """Locate the parts-table grid on a scanned sheet: the region with the most horizontal rules,
    ignoring the page border (too large) and tiny boxes. Returns (x0,y0,x1,y1,linemask) in pixels or None."""
    import cv2, numpy as np
    bw = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 15)
    hk = cv2.getStructuringElement(cv2.MORPH_RECT, (int(50 * z), 1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT, (1, int(25 * z)))
    hmask = cv2.morphologyEx(bw, cv2.MORPH_OPEN, hk)
    vmask = cv2.morphologyEx(bw, cv2.MORPH_OPEN, vk)
    grid = cv2.dilate(cv2.bitwise_or(hmask, vmask), np.ones((5, 5), np.uint8))
    cnts, _ = cv2.findContours(grid, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    H, W = img.shape
    best = None
    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)
        if w * h > 0.5 * W * H or w < 120 * z or h < 40 * z:
            continue
        sub = hmask[y:y + h, x:x + w]
        n, _ = cv2.connectedComponents(cv2.dilate(sub, np.ones((3, 1), np.uint8)))
        if not best or n - 1 > best[0]:
            best = (n - 1, x, y, x + w, y + h)
    if not best or best[0] < 5:
        return None
    _, x0, y0, x1, y1 = best
    n, _, stats, _ = cv2.connectedComponentsWithStats(hmask)
    horizontal_rules = []
    for x, y, w, h, _ in stats[1:n]:
        if w >= 120 * z and h <= 4 * z:
            horizontal_rules.append((x, y, x + w, y + h))
    x_groups = []
    x_tolerance = int(5 * z)
    for rule in sorted(horizontal_rules, key=lambda r: r[0]):
        group = next((g for g in x_groups
                      if abs(rule[0] - g[0][0]) <= x_tolerance
                      and abs(rule[2] - g[0][2]) <= x_tolerance), None)
        if group is None:
            x_groups.append([rule])
        else:
            group.append(rule)
    dense_rule_run = None
    max_row_gap = int(18 * z)
    for group in x_groups:
        runs = []
        for rule in sorted(group, key=lambda r: r[1]):
            center_y = (rule[1] + rule[3]) // 2
            if runs and center_y - runs[-1][-1] <= max_row_gap:
                runs[-1].append(center_y)
            else:
                runs.append([center_y])
        for run in runs:
            if len(run) >= 6 and (dense_rule_run is None or len(run) > len(dense_rule_run)):
                dense_rule_run = (group, run)
    if dense_rule_run:
        group, run = dense_rule_run
        rules = [r for r in group if run[0] - max_row_gap <= (r[1] + r[3]) // 2 <= run[-1] + max_row_gap]
        x0 = min(r[0] for r in rules)
        y0 = min(r[1] for r in rules)
        x1 = max(r[2] for r in rules)
        y1 = max(r[3] for r in rules)
    return x0, y0, x1, y1, cv2.dilate(cv2.bitwise_or(hmask, vmask), np.ones((3, 3), np.uint8)), vmask


def ocr_table_words(img, frame, cfg):
    tess = _get_tesseract()
    if not tess:
        return [], None, []
    try:
        import cv2, pytesseract
        from pytesseract import Output
        pytesseract.pytesseract.tesseract_cmd = tess
        z = frame.z
        reg = find_table_region(img, z)
        if not reg:
            return [], None, []
        x0, y0, x1, y1, lines, vmask = reg
        m = int(6 * z)
        X0, Y0, X1, Y1 = max(0, x0 - m), max(0, y0 - m), min(img.shape[1], x1 + m), min(img.shape[0], y1 + m)
        crop = img[Y0:Y1, X0:X1].copy()
        crop[lines[Y0:Y1, X0:X1] > 0] = 255                     # erase grid lines so they do not confuse OCR
        f = 1.6
        crop = cv2.resize(crop, None, fx=f, fy=f, interpolation=cv2.INTER_CUBIC)
        d = pytesseract.image_to_data(crop, config="--psm 6", output_type=Output.DICT)
        words = []
        for i, t in enumerate(d["text"]):
            t = t.strip()
            if not t or float(d["conf"][i]) < cfg["thresholds"]["ocr_min_conf"]:
                continue
            x, y, w, h = d["left"][i] / f + X0, d["top"][i] / f + Y0, d["width"][i] / f, d["height"][i] / f
            words.append(_mk_word(*frame.box(x, y, x + w, y + h), t, float(d["conf"][i])))
        import numpy as np
        proj = (vmask[y0:y1, x0:x1] > 0).sum(axis=0)
        xs = np.where(proj > 0.6 * (y1 - y0))[0]                       # columns where a vertical rule spans most of the table height
        groups = []
        for x in xs:
            if groups and x - groups[-1][-1] <= 6:
                groups[-1].append(x)
            else:
                groups.append([x])
        grid_x = [frame.pt(x0 + float(np.mean(g)), (y0 + y1) / 2)[0] for g in groups]
        return words, frame.box(X0, Y0, X1, Y1), grid_x
    except Exception:
        return [], None, []


# ------------------------------------------------------------------------------ table from words
def group_lines(words):
    if not words:
        return []
    hs = [w["y1"] - w["y0"] for w in words]
    tol = 0.6 * statistics.median(hs)
    lines = []
    for w in sorted(words, key=lambda w: (w["yc"], w["x0"])):
        if lines and abs(w["yc"] - lines[-1]["yc"]) <= tol:
            L = lines[-1]
            L["words"].append(w)
            L["yc"] = sum(x["yc"] for x in L["words"]) / len(L["words"])
        else:
            lines.append(dict(yc=w["yc"], words=[w]))
    for L in lines:
        L["words"].sort(key=lambda w: w["x0"])
    return lines


def _hdr_tok_eq(a, b, fuzzy):
    if a == b:
        return True
    return fuzzy and len(a) >= 2 and len(b) >= 3 and fuzz.ratio(a, b) >= 75       # OCR-garbled header words ("ty" ~ "qty")


def _match_header(line, syns, fuzzy=False):
    """Return [(canon, x0, x1)] for header phrases found in a line of words."""
    toks = [_norm_hdr(w["text"]) for w in line["words"]]
    found, i = [], 0
    while i < len(toks):
        hit = None
        for syn, canon in syns:
            n = len(syn)
            if len(toks[i:i + n]) == n and all(_hdr_tok_eq(a, b, fuzzy) for a, b in zip(toks[i:i + n], syn)):
                ws = line["words"][i:i + n]
                hgt = sum(w["y1"] - w["y0"] for w in ws) / n
                # words of ONE header phrase are separated by a space; separate header cells by a column gap
                if all(ws[k + 1]["x0"] - ws[k]["x1"] <= 0.7 * hgt for k in range(n - 1)):
                    hit = (canon, n)
                    break
        if hit:
            canon, n = hit
            ws = line["words"][i:i + n]
            found.append((canon, ws[0]["x0"], ws[-1]["x1"]))
            i += n
        else:
            i += 1
    seen, uniq = set(), []
    for f in found:                 # first occurrence of each canonical column only
        if f[0] not in seen:
            seen.add(f[0]); uniq.append(f)
    return uniq


def parse_table(words, cfg, region_left=None, fuzzy=False, grid_x=None):
    syns = _syn_list(cfg)
    lines = group_lines(words)
    hdr = None
    for L in lines:
        h = _match_header(L, syns, fuzzy)
        canons = {c for c, _, _ in h}
        need = 2 if grid_x else 3                                   # with a detected grid, the header only has to LABEL columns
        if len(h) >= need and "description" in canons and "qty" in canons:
            hdr = (L, sorted(h, key=lambda t: t[1])); break
    if not hdr:
        return None
    hline, cols = hdr
    # scanned sheets: the 'Item' header is often unreadable; the grid's left edge tells us where that column starts
    if region_left is not None and "item" not in {c for c, _, _ in cols} and cols[0][1] - region_left > 8:
        cols = [("item", region_left + 2, region_left + 4)] + cols
    pad = 2.0
    bounds = []
    if grid_x and len(grid_x) >= 3:                                 # column edges from the ruled grid (no alignment assumption)
        edges = sorted(grid_x)
        lab = []
        for a, b in zip(edges[:-1], edges[1:]):
            lab.append([next((c for c, hx0, hx1 in cols if a <= (hx0 + hx1) / 2 < b), None), a + 1.5, b])
        names = [l[0] for l in lab]
        if "item" not in names and names[0] is None:
            lab[0][0] = "item"
        if "description" in names and "qty" in names:
            di, qi = names.index("description"), names.index("qty")
            gap = [k for k in range(di + 1, qi) if lab[k][0] is None]
            if len(gap) == 1 and "material" not in names:
                lab[gap[0]][0] = "material"
        bounds = [(c, l, r) for c, l, r in lab if c]
    if not bounds:
        for i, (canon, x0, x1) in enumerate(cols):
            left = x0 - pad
            right = (cols[i + 1][1] - pad) if i + 1 < len(cols) else 1e9
            bounds.append((canon, left, right))
    table_left = bounds[0][1]
    above = [L for L in lines if L["yc"] < hline["yc"] - 1]
    below = [L for L in lines if L["yc"] > hline["yc"] + 1]
    # Keep only words that sit inside the table's x range.
    for side in (above, below):
        for L in side:
            L["words"] = [w for w in L["words"] if w["x0"] >= table_left]
        side[:] = [L for L in side if L["words"]]

    def cell_text(L, i):
        _, lft, rgt = bounds[i]
        ws = [w for w in L["words"] if lft <= w["x0"] < rgt]
        return " ".join(w["text"] for w in ws), ws

    item_i = next((i for i, b in enumerate(bounds) if b[0] == "item"), 0)
    qty_i = next((i for i, b in enumerate(bounds) if b[0] == "qty"), None)
    def row_evidence(side):
        return sum(
            bool(ITEM_RE.match(cell_text(L, item_i)[0].strip()))
            or (qty_i is not None and bool(cell_text(L, qty_i)[0].strip()))
            for L in side
        )

    if row_evidence(above) > row_evidence(below):
        below = sorted(above, key=lambda L: L["yc"], reverse=True)
        direction = -1
    else:
        below = sorted(below, key=lambda L: L["yc"])
        direction = 1

    if (
        not grid_x
        and {c for c, _, _ in bounds} == {"item", "description", "qty"}
        and {"item", "description", "qty"} <= {c for c, _, _ in cols}
    ):
        header_centers = {c: (x0 + x1) / 2 for c, x0, x1 in cols}
        item_center = header_centers["item"]
        description_center = header_centers["description"]
        qty_center = header_centers["qty"]
        item_ends, description_starts, description_ends, qty_starts = [], [], [], []
        for L in below:
            item_words = [w for w in L["words"] if ITEM_RE.fullmatch(w["text"].strip()) and w["xc"] < description_center]
            qty_words = [w for w in L["words"] if re.fullmatch(r"\d+(?:[.,]\d+)?", w["text"].strip()) and w["xc"] > description_center]
            if not item_words or not qty_words:
                continue
            item_word = min(item_words, key=lambda w: abs(w["xc"] - item_center))
            qty_word = min(qty_words, key=lambda w: abs(w["xc"] - qty_center))
            description_words = [w for w in L["words"] if item_word["x1"] <= w["x0"] < qty_word["x0"]]
            if description_words:
                item_ends.append(item_word["x1"])
                description_starts.append(min(w["x0"] for w in description_words))
                description_ends.append(max(w["x1"] for w in description_words))
                qty_starts.append(qty_word["x0"])
        if item_ends and qty_starts:
            item_description_edge = (max(item_ends) + min(description_starts)) / 2
            description_qty_edge = (max(description_ends) + min(qty_starts)) / 2
            if item_description_edge < description_qty_edge:
                bounds = [
                    ("item", bounds[0][1], item_description_edge),
                    ("description", item_description_edge, description_qty_edge),
                    ("qty", description_qty_edge, bounds[-1][2]),
                ]

    rows, prev_yc, pitches = [], hline["yc"], []
    for L in below:
        gap = direction * (L["yc"] - prev_yc)
        if pitches and gap > 2.2 * statistics.median(pitches):
            break                                   # table ended (gap to next text block)
        item_txt = cell_text(L, item_i)[0].strip()
        has_qty = qty_i is not None and bool(cell_text(L, qty_i)[0].strip())
        id_ok = bool(ITEM_RE.match(item_txt))
        if id_ok or has_qty:                        # a line with a quantity is a row even if OCR garbled its item number
            r = dict(fields={}, words=[], y0=1e9, y1=0, id_unreadable=not id_ok)
            rows.append(r)
            if prev_yc != hline["yc"]:
                pitches.append(gap)
        elif rows and not item_txt and (not pitches or gap <= 1.4 * statistics.median(pitches)):
            r = rows[-1]                            # continuation line (wrapped cell)
        else:
            break
        for i, (canon, _, _) in enumerate(bounds):
            txt, ws = cell_text(L, i)
            if txt:
                r["fields"][canon] = (r["fields"].get(canon, "") + " " + txt).strip()
        r["words"] += L["words"]
        r["y0"] = min(r["y0"], min(w["y0"] for w in L["words"]))
        r["y1"] = max(r["y1"], max(w["y1"] for w in L["words"]))
        prev_yc = L["yc"]
    if not rows:
        return None
    right = max(w["x1"] for r in rows for w in r["words"]) + 2
    out = []
    last_num = None
    for r in rows:
        f = r["fields"]
        rid = f.get("item", "").strip()
        inferred = False
        if r["id_unreadable"]:
            rid = str(last_num + 1) if last_num is not None else "?"
            inferred = True
        if rid.isdigit():
            last_num = int(rid)
        out.append(dict(id=rid, id_inferred=inferred, description=f.get("description", ""), material=f.get("material", ""),
                        size=f.get("size", ""), qty=f.get("qty", ""), partno=f.get("partno", ""), remarks=f.get("remarks", ""),
                        bbox=(table_left - 1, r["y0"] - 1.5, right, r["y1"] + 1.5),
                        conf=min(w["conf"] for w in r["words"])))
    return dict(columns=[b[0] for b in bounds], rows=out, header_bbox=(table_left - 1, hline["words"][0]["y0"] - 1.5, right, hline["words"][0]["y1"] + 1.5))


def infer_grid_header(words, grid_x):
    edges = sorted(grid_x or [])
    if len(edges) < 4:
        return []
    cells = list(zip(edges[:-1], edges[1:]))
    if cells[0][1] - cells[0][0] > 0.18 * (edges[-1] - edges[0]):
        return []

    lines = group_lines(words)

    def cell_words(line, index):
        left, right = cells[index]
        return [w for w in line["words"] if left <= w["xc"] < right]

    alpha_scores = [
        sum(len(re.findall(r"[A-Za-z]{2,}", w["text"])) for line in lines for w in cell_words(line, i))
        for i in range(1, len(cells))
    ]
    description_i = 1 + max(range(len(alpha_scores)), key=alpha_scores.__getitem__)
    if alpha_scores[description_i - 1] < 5:
        return []

    qty_candidates = []
    for i in range(description_i + 1, len(cells)):
        values = [w["text"] for line in lines for w in cell_words(line, i)]
        numeric_count = sum(bool(re.fullmatch(r"[\W_]*\d+(?:[.,]\d+)?[\W_]*", value)) for value in values)
        if numeric_count >= 4 and numeric_count / max(1, len(values)) >= 0.6:
            qty_candidates.append((numeric_count / len(values), numeric_count, -(cells[i][1] - cells[i][0]), i))
    if not qty_candidates:
        return []
    qty_i = max(qty_candidates)[-1]

    data_y = []
    for line in lines:
        item_text = " ".join(w["text"] for w in cell_words(line, 0))
        item_text = re.sub(r"^[^\w]+|[^\w]+$", "", item_text)
        description = " ".join(w["text"] for w in cell_words(line, description_i))
        qty_values = [w["text"] for w in cell_words(line, qty_i)]
        has_qty = any(re.fullmatch(r"[\W_]*\d+(?:[.,]\d+)?[\W_]*", value) for value in qty_values)
        if re.search(r"[A-Za-z]{2,}", description) and (ITEM_RE.fullmatch(item_text) or has_qty):
            data_y.append(line["yc"])
    data_y.sort()
    pitches = [b - a for a, b in zip(data_y, data_y[1:]) if b > a]
    if len(data_y) < 4 or not pitches:
        return []
    pitch = statistics.median(pitches)
    header_y = data_y[-1] + pitch
    height = statistics.median(w["y1"] - w["y0"] for w in words)

    def header_word(index, text):
        left, right = cells[index]
        xc = (left + right) / 2
        return _mk_word(xc - 0.3, header_y - height / 2, xc + 0.3, header_y + height / 2, text)

    return [header_word(0, "item"), header_word(description_i, "description"), header_word(qty_i, "qty")]


def parse_table_fallback(words, cfg):
    """Fallback table parser when standard header matching fails."""
    lines = group_lines(words)
    if len(lines) < 2:
        return None
    candidate_rows = []
    for L in lines:
        txts = [w["text"].strip() for w in L["words"]]
        if not txts:
            continue
        first = txts[0]
        if ITEM_RE.match(first) and len(L["words"]) >= 2:
            rid = first
            rest_words = L["words"][1:]
            desc = " ".join(w["text"] for w in rest_words if not w["text"].isdigit())
            qty_word = next((w["text"] for w in rest_words if w["text"].isdigit()), "1")
            x0 = min(w["x0"] for w in L["words"])
            x1 = max(w["x1"] for w in L["words"])
            y0 = min(w["y0"] for w in L["words"])
            y1 = max(w["y1"] for w in L["words"])
            candidate_rows.append(dict(id=rid, id_inferred=False, description=desc or "Part", material="",
                                       size="", qty=qty_word, partno="", remarks="",
                                       bbox=(x0 - 2, y0 - 1.5, x1 + 2, y1 + 1.5), conf=100.0))
    if len(candidate_rows) >= 2:
        tbl_x0 = min(r["bbox"][0] for r in candidate_rows)
        tbl_y0 = min(r["bbox"][1] for r in candidate_rows)
        tbl_x1 = max(r["bbox"][2] for r in candidate_rows)
        tbl_y1 = max(r["bbox"][3] for r in candidate_rows)
        return dict(columns=["item", "description", "qty"], rows=candidate_rows,
                    header_bbox=(tbl_x0, max(0, tbl_y0 - 15), tbl_x1, tbl_y0))
    return None


# ------------------------------------------------------------------------------ balloons
def vector_balloons(page, words):
    out, seen = [], set()
    for d in page.get_drawings():
        r = d["rect"]
        w, h = r.width, r.height
        if not (12 <= w <= 40 and 12 <= h <= 40 and 0.8 <= w / max(h, 1e-6) <= 1.25):
            continue
        ncurve = sum(1 for it in d["items"] if it[0] == "c")
        if ncurve < 4:
            continue
        key = (round(r.x0), round(r.y0))
        if key in seen:
            continue
        seen.add(key)
        inside = [x for x in words if r.x0 <= x["xc"] <= r.x1 and r.y0 <= x["yc"] <= r.y1]
        txt = "".join(x["text"] for x in sorted(inside, key=lambda x: x["x0"])).strip()
        if ITEM_RE.match(txt):
            out.append(dict(id=txt, bbox=(r.x0, r.y0, r.x1, r.y1), conf=100.0))
    return out


def scan_balloons(img, frame, cfg):
    """Hough circles + digit OCR inside each circle (scanned drawings)."""
    tess = _get_tesseract()
    if not tess or img is None:
        return []
    try:
        import cv2, numpy as np, pytesseract
        pytesseract.pytesseract.tesseract_cmd = tess
        z = frame.z
        blur = cv2.GaussianBlur(img, (5, 5), 1.2)
        rmin, rmax = int(7 * z), int(12 * z)
        circles = cv2.HoughCircles(blur, cv2.HOUGH_GRADIENT, dp=1.2, minDist=int(14 * z), param1=120, param2=22, minRadius=rmin, maxRadius=rmax)
        out = []
        if circles is None:
            return out
        inv = 255 - img
        for x, y, r in np.round(circles[0]).astype(int):
            ang = np.linspace(0, 2 * np.pi, 48, endpoint=False)
            px = np.clip((x + r * np.cos(ang)).astype(int), 0, img.shape[1] - 1)
            py = np.clip((y + r * np.sin(ang)).astype(int), 0, img.shape[0] - 1)
            ring = np.array([inv[max(0, yy - 2):yy + 3, max(0, xx - 2):xx + 3].max() for xx, yy in zip(px, py)])
            if (ring > 80).mean() < 0.8:                         # a drawn balloon has a (nearly) closed ring
                continue
            m = int(r * 0.72)
            crop = img[max(0, y - m):y + m, max(0, x - m):x + m]
            if crop.size == 0:
                continue
            crop = cv2.resize(crop, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
            _, bw = cv2.threshold(crop, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            bw = cv2.copyMakeBorder(bw, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=255)
            txt = pytesseract.image_to_string(bw, config="--psm 7 -c tessedit_char_whitelist=0123456789").strip()
            if ITEM_RE.match(txt):
                out.append(dict(id=txt, bbox=frame.box(x - r, y - r, x + r, y + r), conf=60.0))
        return out
    except Exception:
        return []


# ------------------------------------------------------------------------------ drawing
def read_drawing(path, cfg):
    doc = fitz.open(path)
    pages, scanned_any = [], False
    for pno, page in enumerate(doc):
        words = vector_words(page)
        scanned = len(words) < 20
        img = frame = region = None
        tw, grid_x = [], None
        if scanned:
            scanned_any = True
            try:
                img0, z = render_gray(page, cfg["ocr"]["dpi"])
                img, frame, skew = preprocess_scan(img0, z)
                words = ocr_words(img, frame, cfg) or words
                tw, region, grid_x = ocr_table_words(img, frame, cfg)
                if tw:                                           # trust the block-OCR inside the table region
                    rx0, ry0, rx1, ry1 = region
                    words = [w for w in words if not (rx0 <= w["xc"] <= rx1 and ry0 <= w["yc"] <= ry1)] + tw
            except Exception:
                pass
        table = parse_table(words, cfg, region_left=(region[0] if scanned and tw else None), fuzzy=scanned, grid_x=(grid_x if scanned and tw else None))
        if table is None and scanned and tw and grid_x:
            synthetic_header = infer_grid_header(tw, grid_x)
            if synthetic_header:
                table_words = [w for w in words if region and region[0] <= w["xc"] <= region[2] and region[1] <= w["yc"] <= region[3]]
                table = parse_table(table_words + synthetic_header, cfg, region_left=region[0], fuzzy=True, grid_x=grid_x)
        if table is None:
            table = parse_table_fallback(words, cfg)
        balloons = scan_balloons(img, frame, cfg) if scanned else vector_balloons(page, words)
        text = " ".join(w["text"] for w in sorted(words, key=lambda w: (round(w["yc"] / 4), w["x0"])))
        pages.append(dict(page=pno, scanned=scanned, table=table, balloons=balloons, text=text, words=words,
                          size=(page.rect.width, page.rect.height)))
    all_text = " ".join(p["text"] for p in pages)
    return dict(path=path, pages=pages, scanned=scanned_any, revision=find_revision(all_text), text=all_text)


# ------------------------------------------------------------------------------ BOM
def read_bom(path, cfg):
    import openpyxl
    syn = {}
    for canon, syns in cfg["header_synonyms"].items():
        for s in syns:
            syn.setdefault(_norm_hdr(s), canon)
    wb = openpyxl.load_workbook(path, data_only=True)
    best = None
    for ws in wb.worksheets:
        for r in range(1, min(ws.max_row, 40) + 1):
            m = {}
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if v is None:
                    continue
                canon = syn.get(_norm_hdr(v))
                if canon and canon not in m.values():
                    m[c] = canon
            canons = set(m.values())
            if len(m) >= 3 and "description" in canons and "qty" in canons:
                if not best or len(m) > len(best[2]):
                    best = (ws, r, m)
                break
    if not best:
        raise ValueError("BOM header row not found (need at least Description + Qty + one more known column). Edit header_synonyms in config.json.")
    ws, hr, colmap = best
    above = " ".join(str(ws.cell(r, c).value) for r in range(1, hr) for c in range(1, ws.max_column + 1) if ws.cell(r, c).value is not None)
    rows, blanks = [], 0
    for r in range(hr + 1, ws.max_row + 1):
        f = {canon: ws.cell(r, c).value for c, canon in colmap.items()}
        if not any(v not in (None, "") for v in f.values()):
            blanks += 1
            if blanks >= 3:
                break
            continue
        blanks = 0
        if not str(f.get("description") or "").strip():
            continue
        rows.append(dict(row=r, id=str(f.get("item") or "").strip(), description=str(f.get("description") or "").strip(),
                         material=str(f.get("material") or "").strip(), size=str(f.get("size") or "").strip(),
                         qty=f.get("qty"), uom=str(f.get("uom") or "").strip(), partno=str(f.get("partno") or "").strip(),
                         remarks=str(f.get("remarks") or "").strip()))
    return dict(path=path, sheet=ws.title, header_row=hr, colmap={ws.cell(hr, c).value: canon for c, canon in colmap.items()},
                colidx=colmap, rows=rows, revision=find_revision(above), max_col=ws.max_column)
