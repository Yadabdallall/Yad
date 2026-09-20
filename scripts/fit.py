#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ڕێکخستنی درێژی کتێبەکە بۆ ژمارەیەکی دیاریکراوی لاپەڕە.

گەڕانێکی دووبەشی لەسەر بۆشاییی نێوان بڕگەکان ئەنجام دەدات تا کتێبەکە
ڕێک دەکەوێتە سەر ئەو ژمارە لاپەڕەیەی داوا کراوە.

بەکارهێنان:  python3 scripts/fit.py 206
"""
import os
import sys
import json
import subprocess

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "build")
TARGET = int(sys.argv[1]) if len(sys.argv) > 1 else 206


def run(after):
    """بەڵگەنامەکە دروست دەکات و ژمارەی لاپەڕەکانی دەگەڕێنێتەوە."""
    env = dict(os.environ, BOOK_BODY_AFTER=f"{after:.3f}")
    for _ in range(2):                      # دوو خول: پێڕست جێگیر دەبێت
        subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "build.py")],
                       check=True, capture_output=True, env=env)
        r = subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "measure.py")],
                           check=True, capture_output=True, text=True, env=env)
        data = json.loads(r.stdout.strip().splitlines()[-1])
    return data["total_pages"], data


def main():
    lo, hi = 3.0, 14.0                      # کەمترین و زۆرترین بۆشایی
    seen = {}

    n, d = run(7.0)
    seen[7.0] = n
    print(f"  بۆشایی ٧.٠ → {n} لاپەڕە (ئامانج: {TARGET})")
    if n == TARGET:
        print(f"✓ ڕێک لەسەر {TARGET} لاپەڕە")
        return 7.0

    for i in range(18):
        mid = (lo + hi) / 2
        n, d = run(mid)
        seen[round(mid, 3)] = n
        print(f"  بۆشایی {mid:.3f} → {n} لاپەڕە")
        if n == TARGET:
            print(f"✓ ڕێک لەسەر {TARGET} لاپەڕە (بۆشایی {mid:.3f})")
            with open(os.path.join(OUT, "fit.json"), "w") as fh:
                json.dump({"body_after": round(mid, 3), "pages": n}, fh)
            return mid
        if n > TARGET:
            hi = mid
        else:
            lo = mid
        if hi - lo < 0.002:
            break

    print(f"✗ نەگەیشت بە {TARGET}. نزیکترینەکان: "
          f"{sorted(seen.items(), key=lambda kv: abs(kv[1]-TARGET))[:4]}")
    return None


if __name__ == "__main__":
    main()
