"""
build_historical_report.py
─────────────────────────────────────────────────────────────────────────────
Generates the Atlas-style "Fulfillment Distress Analysis" HTML + PDF report as a
companion to the Excel distress report — automatically, from the SAME data that
build_domain_report.py has already loaded. No second data drop, no re-parsing.

The heavy HTML/PDF logic lives in the vendored template complementary_report/
generate.py (with complementary_report/analyze.py for the Distress Universe).
This module imports those files and feeds them data sourced from the Distress
Report project so the two reports stay perfectly in line:

  • Analysis window      = the Excel report's window (WINDOW_START .. WINDOW_END)
  • "Sold properties" (N) = Column D of the Excel report (sold-since-window-start
                            ∩ fulfillment), so the headline number reconciles.
  • Distress Universe     = domain, deduped by FOLIO, BUYBOX SCORE > 0 (Column B scope)

Call generate_historical(...) at the end of build_domain_report.py. Failures are
caught by the caller so the Excel report is never blocked by report generation.
"""

from __future__ import annotations

import importlib.util
import re
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path(__file__).parent
# Report template + signal helpers, vendored into this repo (self-contained —
# no cross-repo dependency). Copied from the former historical_distress_report
# pipeline; the standalone merge/main scaffolding is not needed here because
# build_domain_report.py drives generation from its already-loaded data.
COMPANION_DIR = BASE / "complementary_report"
GENERATE_PY = COMPANION_DIR / "generate.py"
ANALYZE_PY = COMPANION_DIR / "analyze.py"
# Brand assets (CSS + logos) already tracked in this repo. generate.py resolves
# these relative to its own location, so we override its globals after import.
ASSET_CS = BASE / "8020REI-skills-main" / "8020REI-skills-main" / "customer_success"


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _slug(client_name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", client_name.lower())


def _window_label(window_start: pd.Timestamp, window_end: pd.Timestamp) -> str:
    abbr = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
            7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
    sm, sy = abbr[window_start.month], window_start.year
    em, ey = abbr[window_end.month], window_end.year
    return f"{sm} – {em} {sy}" if sy == ey else f"{sm} {sy} – {em} {ey}"


def _build_analysis(gen, ff: pd.DataFrame, dom: pd.DataFrame,
                    window_start: pd.Timestamp,
                    ff_distress_map: dict | None = None,
                    ff_abs=None,
                    dist_map: dict | None = None) -> dict:
    """Build the companion analysis from the shared in-memory frames.

    Signal determination MUST match the Excel audit (Column D). When the Excel's
    per-FOLIO distress data is supplied (ff_distress_map / ff_abs / dist_map), each
    sold property's active signals are taken from it — i.e. "flagged in ANY
    fulfillment month across the window" (union), with Default Risk from the domain
    flag and ABSENTEE the 3-way max (combined in-state + out-of-state). This is the
    same source build_domain_report.py uses for C/D/E/F/G, so counts reconcile
    exactly. If the maps are not supplied (standalone use), it falls back to the
    point-in-time read at the Marketing First Recommendation period.
    """
    ff = ff.copy()
    ff["FOLIO"] = ff["FOLIO"].astype(str).str.strip().str.upper()
    ff["PERIOD"] = ff["Month"].astype(str).str.strip()  # 'Month' is the YYYY-MM filename prefix
    # Per methodology, LAST SALE DATE + MFR come from the domain, signals from the
    # fulfillment. Our fulfillment export also carries these date columns, so drop
    # them here to avoid an _x/_y merge collision (the domain values must win).
    ff = ff.drop(columns=[c for c in ("LAST SALE DATE", "MARKETING FIRST RECOMMENDATION")
                          if c in ff.columns])

    # Domain sales index — already deduped (most-recent sale per FOLIO) upstream.
    sales = dom[["FOLIO", "LAST SALE DATE", "MARKETING FIRST RECOMMENDATION"]].copy()
    sales["FOLIO"] = sales["FOLIO"].astype(str).str.strip().str.upper()
    sales["LAST SALE DATE"] = pd.to_datetime(sales["LAST SALE DATE"], errors="coerce")
    sales = sales.drop_duplicates(subset=["FOLIO"], keep="last")

    df = ff.merge(sales, on="FOLIO", how="left")
    df["LAST SALE DATE"] = pd.to_datetime(df["LAST SALE DATE"], errors="coerce")
    df["MARKETING FIRST RECOMMENDATION"] = pd.to_datetime(
        df["MARKETING FIRST RECOMMENDATION"], errors="coerce"
    )
    df = df[df["LAST SALE DATE"].notna()]
    # Inclusion criterion — matches Column D: sold on/after the window start, no upper bound.
    df = df[df["LAST SALE DATE"] >= window_start]
    df["MFR_MONTH"] = df["MARKETING FIRST RECOMMENDATION"].dt.to_period("M").astype(str)

    # One row per FOLIO — selected at the MFR period purely for the row METADATA
    # (COUNTY, OWNER TYPE, sale date). Signal flags are set separately below.
    rows = []
    for _, group in df.groupby("FOLIO"):
        row, _ = gen.pick_signal_row(group)
        rows.append(row)
    adf = pd.DataFrame(rows).reset_index(drop=True)

    aligned = ff_distress_map is not None and dist_map is not None
    if aligned:
        # ── Excel-aligned signals: "ever-flagged across the window" union ──
        # Restrict to exactly the signals the Excel audit tracks (DOMAIN_DIST_MAP
        # labels + combined ABSENTEE), so the breakdown reconciles to Column D.
        signal_to_label = {c: lbl for c, lbl in dist_map.items() if c in gen.SIGNAL_COLS}
        available_signals = [c for c in gen.SIGNAL_COLS
                             if c == "ABSENTEE" or c in signal_to_label]
        # Normalise keys (strip + upper) to match adf["FOLIO"], regardless of how
        # the source dict/series was keyed.
        dmap = {str(k).strip().upper(): v for k, v in ff_distress_map.items()}
        amap = {str(k).strip().upper(): v for k, v in dict(ff_abs).items()} if ff_abs is not None else {}
        folios = adf["FOLIO"]
        for col in available_signals:
            if col == "ABSENTEE":
                flag = folios.map(lambda f: 1 if amap.get(f) in (1, 2) else 0)
            else:
                lbl = signal_to_label[col]
                flag = folios.map(lambda f, l=lbl: 1 if l in dmap.get(f, ()) else 0)
            adf[f"{col}_active"] = flag.astype(int)
            adf[col] = adf[f"{col}_active"]
    else:
        # ── Fallback: point-in-time read at MFR period (standalone use) ──
        available_signals = [c for c in gen.SIGNAL_COLS if c in df.columns]
        for col in available_signals:
            adf[col] = pd.to_numeric(adf[col], errors="coerce")
            adf[f"{col}_active"] = gen.is_active(adf[col], col).astype(int)

    adf["SIGNAL_COUNT"] = adf[[f"{c}_active" for c in available_signals]].sum(axis=1)

    def active_signals_str(row: pd.Series) -> str:
        parts = [gen.LABEL_MAP.get(col, col.title())
                 for col in available_signals if row.get(f"{col}_active", 0) == 1]
        return ", ".join(parts) if parts else "&#8212;"

    adf["ACTIVE_SIGNALS_STR"] = adf.apply(active_signals_str, axis=1)

    dov_total, dov_county = _build_overview(dom)
    all_buybox_counties = sorted(adf["COUNTY"].dropna().astype(str).unique().tolist())

    return {
        "adf": adf,
        "available_signals": available_signals,
        "total_rows": len(ff),
        "dov_total": dov_total,
        "dov_county": dov_county,
        "all_buybox_counties": all_buybox_counties,
    }


def _build_overview(dom: pd.DataFrame) -> tuple[pd.Series, pd.DataFrame]:
    """Distress Universe table (Page 4) — domain deduped + BUYBOX SCORE > 0,
    signals counted per county. Mirrors analyze.py so the universe scope equals
    the Excel report's Column B scope."""
    analyze = _load_module("hist_analyze", ANALYZE_PY)
    signals = analyze.STANDARD_SIGNALS
    absentee = analyze.ABSENTEE_COL

    d = dom.copy()
    d.columns = d.columns.str.strip().str.upper()
    d["BUYBOX SCORE"] = pd.to_numeric(d.get("BUYBOX SCORE", 0), errors="coerce").fillna(0)
    d = d[d["BUYBOX SCORE"] > 0]

    present = [c for c in signals if c in d.columns]
    for col in present + ([absentee] if absentee in d.columns else []):
        d[col] = pd.to_numeric(d[col], errors="coerce")

    def counts(group: pd.DataFrame) -> pd.Series:
        out = {c: int(analyze.is_active(group[c], c).sum()) for c in present}
        if absentee in d.columns:
            out[absentee] = int(analyze.is_active(group[absentee], absentee).sum())
        out["TOTAL PROPERTIES"] = len(group)
        return pd.Series(out)

    by_county = d.groupby("COUNTY").apply(counts).reset_index()

    total = {c: int(analyze.is_active(d[c], c).sum()) for c in present}
    if absentee in d.columns:
        total[absentee] = int(analyze.is_active(d[absentee], absentee).sum())
    total["TOTAL PROPERTIES"] = len(d)
    return pd.Series(total), by_county


def generate_historical(client_name: str,
                        window_start: pd.Timestamp,
                        window_end: pd.Timestamp,
                        ff: pd.DataFrame,
                        dom: pd.DataFrame,
                        out_dir: Path,
                        report_date: str | None = None,
                        ff_distress_map: dict | None = None,
                        ff_abs=None,
                        dist_map: dict | None = None) -> dict:
    """Build the HTML + PDF report next to the Excel report. Returns paths + N.

    Pass ff_distress_map / ff_abs / dist_map (from build_domain_report.py) so the
    sold-property distress breakdown reconciles exactly with the Excel audit.
    """
    if not GENERATE_PY.exists():
        raise FileNotFoundError(
            f"Companion report template not found: {GENERATE_PY}\n"
            "  Expected the vendored template at complementary_report/generate.py."
        )
    gen = _load_module("hist_generate", GENERATE_PY)
    # Point the template at the brand assets tracked in this repo (generate.py
    # otherwise resolves them relative to its own former pipeline location).
    gen.REPORT_CSS_FILE = ASSET_CS / "standards" / "report.css"
    gen.LOGO_FULL = ASSET_CS / "logos" / "logo-full-light.png"
    gen.LOGO_ICON = ASSET_CS / "logos" / "logo-icon-light.png"

    report_date = report_date or pd.Timestamp.today().strftime("%B %Y")
    report_month = pd.Timestamp.today().strftime("%Y-%m")
    window_label = _window_label(window_start, window_end)

    print("\nGenerating companion Fulfillment Distress Analysis (HTML + PDF)...")
    print(f"  Window: {window_label}  |  Report date: {report_date}")

    analysis = _build_analysis(gen, ff, dom, window_start,
                               ff_distress_map=ff_distress_map, ff_abs=ff_abs, dist_map=dist_map)
    n = len(analysis["adf"])
    print(f"  Matched sold properties (= Column D): {n:,}")
    if ff_distress_map is not None:
        print("  Signal source: Excel-aligned (ever-flagged across window + domain Default Risk)")
    print(f"  Counties with matched sales: "
          f"{len([c for c in analysis['all_buybox_counties']])}")

    html = gen.build_html(
        client_name=client_name,
        analysis=analysis,
        report_date=report_date,
        window_label=window_label,
    )

    stem = f"{client_name} - Fulfillment Distress Analysis - {report_month}"
    html_path = out_dir / f"{stem}.html"
    pdf_path = out_dir / f"{stem}.pdf"
    html_path.write_text(html, encoding="utf-8")
    print(f"  HTML: {html_path.name}")

    gen.export_pdf(html_path, pdf_path)
    if pdf_path.exists():
        print(f"  PDF : {pdf_path.name}")
    else:
        print("  PDF : skipped (Playwright not installed) — open the HTML and print to PDF.")

    return {"n": n, "html": html_path, "pdf": pdf_path if pdf_path.exists() else None}
