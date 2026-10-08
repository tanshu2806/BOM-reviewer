# Drawing ↔ BOM Completeness Checker — Prototype

Takes an **engineering drawing (PDF)** and a **BOM (Excel)**, compares them, and **highlights what is missing or inconsistent on both the drawing sheet and the BOM sheet**.
Runs **fully offline**. Original files are never modified.

> Built **before the MoU**, so every input detail is an *assumption* (see `bomcheck/config.json` and the assumption register).
> Tested only on **synthetic** boiler-economiser data. Treat accuracy figures as "pipeline works", not "will work on your drawings".

## Quick start

```bash
pip install -r requirements.txt          # + Tesseract OCR installed on the machine (only needed for scanned PDFs)
python run_check.py samples/sample_01_drawing.pdf samples/sample_01_bom.xlsx --out outputs
```

### Presentation UI demo

```bash
streamlit run app.py
```

This starts a simple web dashboard that lets you upload a drawing PDF and BOM Excel file, run the checker, and download the annotated PDF, Excel output, and JSON report.

Outputs (in `--out`):

| File | What it is |
|---|---|
| `<drawing>_annotated.pdf` | The **original drawing** with coloured boxes + labels on every flagged item, a legend, and a summary page |
| `<drawing>_BOM_checked.xlsx` | Colour-coded BOM with **Check status / Check note** columns, **inserted red "MISSING – ADD TO BOM" rows**, plus `Discrepancies` and `Summary` sheets |
| `<drawing>_report.json` | Machine-readable result |

Colours: 🔴 on drawing but missing in BOM · 🟠 quantity / spec / revision mismatch · 🟡 confirm match / extra line / duplicate · 🔵 rule-based suggestion (implied item) · 🟢 OK · ⚪ non-drawn item (consumable)

## Pipeline (I-E-N-M-C-H)

| Stage | Module | What it does |
|---|---|---|
| **I**ngest | `extract.py` | Detects vector vs scanned PDF; reads BOM (header row auto-detected via synonyms) |
| **E**xtract | `extract.py` | Parts table → rows with bounding boxes; balloons; notes; revision. Scans: shadow removal → deskew → table-grid detection → OCR (Tesseract) |
| **N**ormalise | `match.py` | Abbreviations, phrase synonyms, units/numbers, material codes |
| **M**atch | `match.py` | Fuzzy one-to-one assignment (text similarity + number similarity + ID/part-no bonus) |
| **C**ompare | `match.py` | Omission, quantity, spec, extra, duplicate, revision, "per-bank ×N" multiplier, implied-item rules; **run-reliability gate** |
| **H**ighlight | `report.py` | Annotated PDF + colour-coded Excel + JSON |

## Where the assumptions live

| Assumption (register ID) | Where it is encoded | Change it by |
|---|---|---|
| D8/D9 parts-table columns, B2 BOM columns | `header_synonyms` | edit config |
| Wording / abbreviations (B4) | `abbreviations`, `phrase_synonyms` | build from the client's real BOM vocabulary |
| T1 match thresholds | `thresholds` (match 85, possible 70, duplicate 78, token_fuzzy 80) | tune on real pairs |
| C1 hotspots / implied items | `implied_rules`, `non_drawn_patterns` | add rules from the client's past-omission list |
| "typical ×N per bank" | `multiplier` | edit regex |
| OCR (D2/D6) | `ocr` (dpi, psm), `thresholds.ocr_min_conf` | edit config |
| Reliability gate | `reliability` | edit config |

## Results so far (synthetic data — see `EVALUATION.md` for full output)

| Test | Result |
|---|---|
| 15 vector drawings (tuning set) | 78/78 injected errors found, 0 false alarms |
| 30 vector drawings (held-out seeds) | 153/153 found, 0 false alarms |
| 5 mildly degraded scans (OCR) | 28/28 found, 0 false alarms |
| **Stress A**: BOM wording with typos/abbreviations/reordering | 153/153 found, **8 false alarms (precision 95%)** — after adding an abbreviation list that overlaps the test's own abbreviations, so optimistic |
| **Stress B**: heavily degraded scans (150 dpi, skew, blur, noise, shadow, JPEG q40) | **Extraction fails** (recall 34%, precision 9%) — but all 10 runs are flagged **LOW reliability** with a banner on every output |

## Known limits (be upfront about these)

1. **Only tested on data I generated.** Real drawings will differ (fonts, wrapped cells, merged cells, multi-page tables, non-English text, CAD-specific table layouts).
2. **Vector tables assume left-aligned text under headers** (scans use the ruled grid instead). Fix = use the vertical rules for vector PDFs too.
3. **Poor scans fail.** OCR with Tesseract collapses on heavily degraded sheets; the tool says so (LOW reliability) rather than guessing. A stronger OCR engine and real scan samples are needed.
4. **Balloon detection on scans is partial (~60%)**; on scans the parts table is the primary source.
5. **Implied items (blue) are rule-based suggestions**, not proof. Rules are generic placeholders until the client's past omissions are known.
6. **Not handled yet:** DWG/DXF input, multi-page continued tables, rotated pages, 3D/CAD-model BOMs, multi-drawing → one BOM mapping, non-English, unit conversion beyond simple numbers.
7. **Quantities from calculated items** (tube lengths, insulation area) are only compared as numbers; no geometry-based quantity checking.

## Regenerating test data

```bash
python samples/generate_samples.py 15                          # tuning set into samples/
python samples/generate_samples.py 30 samples_heldout 7000     # held-out set
python evaluate.py --samples samples                           # metrics on vector PDFs
python evaluate.py --samples samples --scans                   # metrics on the scanned copies
python stress_test.py                                          # harder conditions
```

## What I need from the company after the MoU (in priority order)

1. 3–5 real drawing + BOM pairs (masked or closed projects) — decides everything below.
2. Whether drawings are vector, scanned, or CAD, and whether parts table + balloons exist.
3. A blank BOM template (column names) and the BOM abbreviation/vocabulary list.
4. The list of past omissions (even as plain text) to build realistic implied-item rules.
5. What "success" means (e.g. recall target, max false alarms per drawing).

## Project layout

```
run_check.py            CLI entry point
bomcheck/config.json    ALL assumptions (edit here)
bomcheck/extract.py     drawing (vector + OCR) and BOM readers
bomcheck/match.py       normalise, match, compare, rules, reliability gate
bomcheck/report.py      annotated PDF, colour-coded Excel, JSON
samples/                synthetic drawings + BOMs + ground truth (15 pairs)
evaluate.py             precision/recall on injected errors
stress_test.py          noisy-wording and heavy-scan tests
demo/                   example outputs (vector, scanned, unreliable-flagged)
EVALUATION.md           full evaluation output
```
