#!/usr/bin/env python3
"""
Evaluate the checker on the synthetic proxy set (known injected errors).

   python evaluate.py [--samples samples] [--scans]

Reports, per error type: injected / found (recall), and false alarms (precision).
NOTE: results are on SYNTHETIC data that the author generated - they show the pipeline works,
NOT how it will perform on the client's real drawings. Re-run on real pairs after the MoU.
"""
import argparse, glob, json, os, warnings
warnings.filterwarnings("ignore")
from bomcheck.extract import load_config, read_drawing, read_bom
from bomcheck.match import compare

HARD = ["omission", "quantity", "spec", "multiplier", "extra", "duplicate", "revision", "implied"]


def key_of(iss):
    t = iss["type"]
    if t in ("omission", "quantity", "spec", "multiplier"):
        return (t, iss.get("drawing_id"))
    if t in ("extra", "duplicate"):
        return (t, iss.get("bom_desc"))
    if t == "implied":
        return (t, iss.get("label"))
    return (t, "revision")


def evaluate(samples_dir, scans=False, cfg=None):
    cfg = cfg or load_config()
    stats = {t: dict(truth=0, tp=0, fp=0) for t in HARD}
    any_detect = dict(truth=0, hit=0)
    per_sample, confirm_total, n = [], 0, 0
    rel = {'HIGH': 0, 'MEDIUM': 0, 'LOW': 0}
    pat = "sample_*_drawing_scan.pdf" if scans else "sample_*_drawing.pdf"
    for pdf in sorted(glob.glob(os.path.join(samples_dir, pat))):
        idx = os.path.basename(pdf).split("_")[1]
        truth = json.load(open(os.path.join(samples_dir, f"sample_{idx}_truth.json")))
        bom = read_bom(os.path.join(samples_dir, f"sample_{idx}_bom.xlsx"), cfg)
        drawing = read_drawing(pdf, cfg)
        res = compare(drawing, bom, cfg)
        tkeys = {(e["type"], e["key"] if e["type"] != "revision" else "revision") for e in truth["errors"]}
        fkeys = {key_of(i) for i in res["issues"] if i["type"] in HARD}
        # A duplicated BOM line is two identical lines: which of the pair is "the duplicate" is arbitrary,
        # so score duplicates at ITEM level (any duplicate flag on the same drawing item counts).
        for e in truth["errors"]:
            if e["type"] == "duplicate":
                hit = next((i for i in res["issues"] if i["type"] == "duplicate" and i.get("drawing_id") == e["item"]), None)
                if hit:
                    fkeys.discard(("duplicate", hit.get("bom_desc"))); fkeys.add(("duplicate", e["key"]))
        for t, k in tkeys:
            stats[t]["truth"] += 1
            if (t, k) in fkeys:
                stats[t]["tp"] += 1
        for t, k in fkeys - tkeys:
            stats[t]["fp"] += 1
        # "any flag on the right item", whatever the type
        flagged_items = {i.get("drawing_id") for i in res["issues"]} | {i.get("bom_desc") for i in res["issues"]} | {i.get("label") for i in res["issues"]}
        dup_item = {e["key"]: e["item"] for e in truth["errors"] if e["type"] == "duplicate"}
        for t, k in tkeys:
            any_detect["truth"] += 1
            if k in flagged_items or dup_item.get(k) in flagged_items or (t == "revision" and any(i["type"] == "revision" for i in res["issues"])):
                any_detect["hit"] += 1
        rel[res['reliability']] += 1
        c = sum(1 for i in res["issues"] if i["type"] == "confirm")
        confirm_total += c; n += 1
        per_sample.append(dict(sample=idx, truth=len(tkeys), found=len(tkeys & fkeys), false_alarms=len(fkeys - tkeys), confirm=c,
                               missed=sorted(map(str, tkeys - fkeys)), fps=sorted(map(str, fkeys - tkeys))))
    stats['_rel'] = rel
    return stats, any_detect, per_sample, confirm_total, n


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--samples", default="samples"); ap.add_argument("--scans", action="store_true")
    a = ap.parse_args()
    stats, anyd, per, conf, n = evaluate(a.samples, a.scans)
    rel = stats.pop('_rel')
    print(f"{'SCANNED (OCR)' if a.scans else 'VECTOR'} drawings: {n} samples   run reliability: {rel}\n")
    print(f"{'type':11}{'injected':>9}{'found':>7}{'recall':>8}{'false alarms':>14}{'precision':>11}")
    T = TP = FP = 0
    for t in HARD:
        s = stats[t]
        if s["truth"] == 0 and s["fp"] == 0:
            continue
        rec = f"{100*s['tp']/s['truth']:.0f}%" if s["truth"] else "-"
        prec = f"{100*s['tp']/(s['tp']+s['fp']):.0f}%" if (s["tp"] + s["fp"]) else "-"
        print(f"{t:11}{s['truth']:>9}{s['tp']:>7}{rec:>8}{s['fp']:>14}{prec:>11}")
        T += s["truth"]; TP += s["tp"]; FP += s["fp"]
    print(f"{'ALL':11}{T:>9}{TP:>7}{100*TP/T:>7.0f}%{FP:>14}{100*TP/(TP+FP):>10.0f}%")
    print(f"\nRight item flagged (any type): {anyd['hit']}/{anyd['truth']} = {100*anyd['hit']/anyd['truth']:.0f}%")
    print(f"Review burden: {FP/n:.2f} false alarms + {conf/n:.2f} 'confirm' flags per drawing")
    bad = [p for p in per if p["missed"] or p["fps"]]
    if bad:
        print("\nDetail (samples with misses / false alarms):")
        for p in bad:
            print(f"  sample {p['sample']}: found {p['found']}/{p['truth']}  missed={p['missed']}  false_alarms={p['fps']}")
