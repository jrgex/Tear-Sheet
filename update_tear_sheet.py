#!/usr/bin/env python3
"""
BEATS Tear Sheet Generator
Reads 'data/Beats data .xlsx', computes performance statistics,
and writes a styled Excel tear sheet to output/.
"""

import sys
import math
import os
from datetime import datetime, date
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter

# ─── Paths ────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).parent
DATA_FILE = ROOT / "data" / "Beats data .xlsx"
OUTPUT_DIR = ROOT / "output"

# ─── Fund constants (static) ──────────────────────────────────────────────────
FUND_NAME       = "GREEN RENEWABLE STORAGE ENERGY MARKET\nPARTICIPATION AND TRADING"
FUND_NAME_SHORT = "Green Renewable Storage Energy Market Participation and Trading"
MANAGER         = "Green Exchange"
MIN_INVESTMENT  = "50,000 AUD"
LIQUIDITY       = "Quarterly"
MGMT_FEE        = "2.00%"
PERF_FEE        = "15.00%"
HWM             = "Yes"
PHONE           = "+61 2 8004 0234"
EMAIL           = "jr@greenexchange.com.au"
LEGAL_NOTE = (
    "NOTE: This publication has been prepared on behalf of and issued by Green Exchange Pty Ltd "
    "(ACN 654 125 247) holding AFSL 559619. This is for educational purposes only. This is not "
    "an offer to deal in any financial product. This information might contain unsolicited general "
    "information only, without regard to any investor's individual objectives, financial situation "
    "or needs. It is not specific advice for any particular investor. Before making any decision "
    "about the information provided, you must consider the appropriateness of the information in "
    "this document, having regard to your objectives, financial situation and needs and consult "
    "your adviser. This may not be passed on to anyone but the addressee. Past performance of "
    "financial products is no assurance of future performance."
)
STRATEGY_DESC = (
    "BEATS  Battery Electronic Automated Trading System\n\n"
    "Deployment of strategic Battery assets across Eastern Australia. The Energy landscape in "
    "Australia presents an Investment opportunity in BESS. This offers multiple revenue streams "
    "across the entire value chain. Installation of hardware with a 10+ year lifespan and market "
    "participation through proprietary software enable IRR of 26%+ with a payback period of 2-3 years.\n\n"
    "Certificate Generation and Trading coupled with creative commercial models further add to yields."
)
KEY_HIGHLIGHTS = [
    "Battery System Installation (BESS)",
    "System Integration",
    "Certificate Generation",
    "Trading",
    "Commercial Models",
]

# ─── Colours ──────────────────────────────────────────────────────────────────
C_HEADER_BG   = "#0A1828"   # very dark navy
C_HEADER_FG   = "#FFFFFF"
C_SECTION_BG  = "#1C3A2A"   # dark forest green
C_SECTION_FG  = "#FFFFFF"
C_ACCENT      = "#2D6A4F"   # mid green
C_LABEL_BG    = "#F0F4F0"   # very light green-grey
C_ROW_ALT     = "#F7FAF7"
C_POS         = "#006400"   # positive return
C_NEG         = "#CC0000"   # negative return
C_BORDER      = "#CCCCCC"
C_TABLE_HDR   = "#1C3A2A"
C_WHITE       = "#FFFFFF"
C_LIGHT_GREY  = "#F5F5F5"

MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]


# ═══════════════════════════════════════════════════════════════════════════════
# Data loading & statistics
# ═══════════════════════════════════════════════════════════════════════════════

def load_data(filepath: Path) -> pd.DataFrame:
    # Find the header row: scan for the first row that has 'Month' or 'Beats'
    raw = pd.read_excel(filepath, header=None)
    header_row = 0
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() for v in row if pd.notna(v)]
        if any("month" in v or "beat" in v or "return" in v for v in vals):
            header_row = i
            break

    df = pd.read_excel(filepath, header=header_row)
    df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)
    df.columns = [str(c).strip() for c in df.columns]

    # Identify month and return columns by name pattern
    month_col = next(
        (c for c in df.columns if "month" in c.lower() or "date" in c.lower()), None
    )
    beats_col = next(
        (c for c in df.columns if "beat" in c.lower() or "return" in c.lower()), None
    )

    # Fallback: use positional columns if names not recognised
    if month_col is None or beats_col is None:
        cols = df.columns.tolist()
        month_col = cols[0] if month_col is None else month_col
        beats_col = cols[1] if beats_col is None else beats_col

    df = df[[month_col, beats_col]].copy()
    df.columns = ["month", "return"]
    df["month"] = pd.to_datetime(df["month"], errors="coerce")
    df["return"] = pd.to_numeric(df["return"], errors="coerce")
    df = df.dropna().sort_values("month").reset_index(drop=True)
    return df


def build_stats(df: pd.DataFrame) -> dict:
    rets = df["return"].values          # decimal monthly returns
    n = len(rets)
    dates = df["month"].tolist()

    # Cumulative / periodic returns
    total_return   = float(np.prod(1 + rets) - 1)
    latest         = df["month"].iloc[-1]
    ytd_mask       = df["month"].dt.year == latest.year
    ytd_rets       = df.loc[ytd_mask, "return"].values
    ytd_return     = float(np.prod(1 + ytd_rets) - 1) if len(ytd_rets) else 0.0
    three_mo_rets  = rets[-3:] if n >= 3 else rets
    three_mo_ror   = float(np.prod(1 + three_mo_rets) - 1)
    six_mo_rets    = rets[-6:] if n >= 6 else rets
    six_mo_ror     = float(np.prod(1 + six_mo_rets) - 1)

    # Annualised
    ann_return = (1 + total_return) ** (12 / n) - 1
    ann_std    = float(np.std(rets, ddof=1)) * math.sqrt(12)
    sharpe     = ann_return / ann_std if ann_std else 0.0

    winning_pct = float(np.sum(rets > 0) / n * 100)

    # VAMI (Value Added Monthly Index, starts at 1000)
    vami = [1000.0]
    for r in rets:
        vami.append(vami[-1] * (1 + r))
    vami_dates = [df["month"].iloc[0]] + dates  # one extra point at start

    # Monthly performance table  {year: {month: pct_return}, 'annual': pct}
    monthly_table: dict[int, dict] = {}
    for _, row in df.iterrows():
        y = row["month"].year
        m = row["month"].month
        monthly_table.setdefault(y, {})[m] = row["return"] * 100  # → %

    annual_rets: dict[int, float] = {}
    for y, months in monthly_table.items():
        r_list = [months[m] / 100 for m in sorted(months)]
        annual_rets[y] = (np.prod([1 + r for r in r_list]) - 1) * 100

    return {
        "total_return":   total_return * 100,
        "ytd":            ytd_return * 100,
        "three_mo_ror":   three_mo_ror * 100,
        "six_mo_ror":     six_mo_ror * 100,
        "sharpe":         sharpe,
        "winning_pct":    winning_pct,
        "ann_return":     ann_return * 100,
        "vami":           vami,
        "vami_dates":     vami_dates,
        "monthly_table":  monthly_table,
        "annual_rets":    annual_rets,
        "latest_date":    latest,
        "start_date":     df["month"].iloc[0],
        "all_returns":    rets.tolist(),
        "all_dates":      dates,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# Excel generation
# ═══════════════════════════════════════════════════════════════════════════════

def write_excel(df: pd.DataFrame, stats: dict, output_path: Path):
    wb = xlsxwriter.Workbook(str(output_path), {"nan_inf_to_errors": True})

    # ── hidden data sheet (for charts) ────────────────────────────────────────
    ds = wb.add_worksheet("_data")
    ds.hide()
    _write_chart_data(wb, ds, df, stats)

    # ── main sheet ────────────────────────────────────────────────────────────
    ws = wb.add_worksheet("Tear Sheet")
    ws.set_zoom(85)
    ws.set_paper(9)       # A4
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.hide_gridlines(2)
    ws.set_margins(left=0.4, right=0.4, top=0.4, bottom=0.4)

    # Column widths (total ~26 cols)
    col_widths = (
        [1.5] +          # A: spacer
        [9.0] * 5 +      # B-F: left label area
        [2.5] * 3 +      # G-I: left value / mid
        [1.5] +          # J: gutter
        [1.0] +          # K: spacer
        [8.0] * 5 +      # L-P: right label area
        [6.0] * 4 +      # Q-T: right values
        [1.5]            # U: right edge
    )
    for ci, w in enumerate(col_widths):
        ws.set_column(ci, ci, w)

    # ── formats ───────────────────────────────────────────────────────────────
    fmt = _formats(wb)

    row = 0
    latest = stats["latest_date"]
    date_str = latest.strftime("%B %Y")

    # ── Row 0-1: master header ────────────────────────────────────────────────
    ws.set_row(row, 22)
    ws.merge_range(row, 0, row, 9,  date_str, fmt["hdr_date"])
    ws.merge_range(row, 10, row, 21, FUND_NAME.replace("\n", "  "), fmt["hdr_title"])
    row += 1
    ws.set_row(row, 4)
    ws.merge_range(row, 0, row, 21, "", fmt["hdr_bar"])
    row += 1

    # ── Row 2: section labels row ─────────────────────────────────────────────
    ws.set_row(row, 16)
    ws.merge_range(row, 0, row, 9,  "KEY HIGHLIGHTS",        fmt["sec_label"])
    ws.merge_range(row, 10, row, 21, "MANAGER",              fmt["sec_label"])
    row += 1

    # ── Key Highlights + Manager ───────────────────────────────────────────────
    for i, kh in enumerate(KEY_HIGHLIGHTS):
        ws.set_row(row, 14)
        ws.merge_range(row, 0, row, 9,  f"  • {kh}", fmt["bullet"])
        if i == 0:
            ws.merge_range(row, 10, row, 21, f"  {MANAGER}", fmt["manager_name"])
        else:
            ws.merge_range(row, 10, row, 21, "", fmt["body"])
        row += 1

    # ── Performance Statistics label ───────────────────────────────────────────
    ws.set_row(row, 16)
    ws.merge_range(row, 0, row, 9,  "", fmt["sec_divider"])
    ws.merge_range(row, 10, row, 21, "PERFORMANCE STATISTICS", fmt["sec_label"])
    row += 1

    # Stats 2×2 grid
    ws.set_row(row, 13)
    ws.merge_range(row, 10, row, 15, "3 Month ROR",     fmt["stat_label"])
    ws.merge_range(row, 16, row, 21, "Year To Date",    fmt["stat_label"])
    row += 1
    ws.set_row(row, 18)
    ws.merge_range(row, 10, row, 15, f"{stats['three_mo_ror']:.2f}%", fmt["stat_big"])
    ws.merge_range(row, 16, row, 21, f"{stats['ytd']:.2f}%",          fmt["stat_big"])
    row += 1
    ws.set_row(row, 13)
    ws.merge_range(row, 10, row, 15, "Total Return Cumulative", fmt["stat_label"])
    ws.merge_range(row, 16, row, 21, "6 Month ROR",             fmt["stat_label"])
    row += 1
    ws.set_row(row, 18)
    ws.merge_range(row, 10, row, 15, f"{stats['total_return']:.2f}%", fmt["stat_big"])
    ws.merge_range(row, 16, row, 21, f"{stats['six_mo_ror']:.2f}%",   fmt["stat_big"])
    row += 1

    # ── Strategy Description ───────────────────────────────────────────────────
    ws.set_row(row, 16)
    ws.merge_range(row, 0, row, 9,  "STRATEGY DESCRIPTION", fmt["sec_label"])
    ws.merge_range(row, 10, row, 21, "GENERAL INFORMATION",  fmt["sec_label"])
    row += 1

    strat_lines = STRATEGY_DESC.split("\n")
    gi_rows = [
        ("Minimum Investment", MIN_INVESTMENT),
        ("Liquidity",          LIQUIDITY),
        ("Management Fee",     MGMT_FEE),
        ("Performance Fee",    PERF_FEE),
        ("Highwater Mark",     HWM),
    ]
    max_gi = max(len(strat_lines), len(gi_rows))
    for i in range(max_gi):
        ws.set_row(row, 13)
        left_text = f"  {strat_lines[i]}" if i < len(strat_lines) else ""
        ws.merge_range(row, 0, row, 9, left_text, fmt["body_small"])
        if i < len(gi_rows):
            label, val = gi_rows[i]
            ws.merge_range(row, 10, row, 15, f"  {label}", fmt["gi_label"])
            ws.merge_range(row, 16, row, 21, val,          fmt["gi_value"])
        else:
            ws.merge_range(row, 10, row, 21, "", fmt["body"])
        row += 1

    # ── VAMI chart + Statistics ────────────────────────────────────────────────
    ws.set_row(row, 16)
    ws.merge_range(row, 0, row, 9,  "PERFORMANCE (VAMI)",  fmt["sec_label"])
    ws.merge_range(row, 10, row, 21, "STATISTICS",          fmt["sec_label"])
    chart_top_row = row
    row += 1

    stats_items = [
        ("Total Return Cumulative", f"{stats['total_return']:.2f}%"),
        ("Sharpe Ratio",            f"{stats['sharpe']:.2f}"),
        ("Winning Months (%)",      f"{stats['winning_pct']:.2f}%"),
        ("Alpha Annualized",        f"{stats['ann_return']:.2f}%"),
    ]
    for label, val in stats_items:
        ws.set_row(row, 14)
        ws.merge_range(row, 10, row, 17, f"  {label}", fmt["stat_row_label"])
        ws.merge_range(row, 18, row, 21, val,           fmt["stat_row_value"])
        row += 1

    # VAMI chart (14 rows tall × 10 cols wide, anchored at chart_top_row+1)
    vami_chart = _make_vami_chart(wb, ds, stats)
    ws.insert_chart(chart_top_row + 1, 0, vami_chart,
                    {"x_offset": 2, "y_offset": 2, "x_scale": 0.97, "y_scale": 1.2})
    chart_bottom = chart_top_row + 16
    row = max(row, chart_bottom) + 1

    # ── Monthly Returns chart ──────────────────────────────────────────────────
    ws.set_row(row, 16)
    ws.merge_range(row, 0, row, 21, "MONTHLY RETURNS", fmt["sec_label"])
    mr_chart_row = row
    row += 1

    mr_chart = _make_monthly_returns_chart(wb, ds, stats)
    ws.insert_chart(mr_chart_row + 1, 0, mr_chart,
                    {"x_offset": 2, "y_offset": 2, "x_scale": 1.96, "y_scale": 1.1})
    row = mr_chart_row + 14

    # ── Monthly Performance table ──────────────────────────────────────────────
    ws.set_row(row, 16)
    ws.merge_range(row, 0, row, 21, "MONTHLY PERFORMANCE", fmt["sec_label"])
    row += 1

    # Header row
    ws.set_row(row, 14)
    ws.write(row, 0, "Year", fmt["tbl_hdr"])
    for mi, mn in enumerate(MONTH_NAMES):
        ws.write(row, 1 + mi, mn, fmt["tbl_hdr"])
    ws.write(row, 13, "Year Return", fmt["tbl_hdr"])
    row += 1

    mt = stats["monthly_table"]
    ar = stats["annual_rets"]
    for yi, year in enumerate(sorted(mt.keys(), reverse=True)):
        ws.set_row(row, 13)
        bg = C_ROW_ALT if yi % 2 else C_WHITE
        ws.write(row, 0, year, fmt["tbl_year"])
        for m_idx in range(1, 13):
            val = mt[year].get(m_idx)
            if val is not None:
                cell_fmt = _return_fmt(wb, val, bg)
                ws.write(row, m_idx, round(val, 2), cell_fmt)
            else:
                ws.write_blank(row, m_idx, None, fmt["tbl_empty"])
        ann = ar.get(year)
        if ann is not None:
            ws.write(row, 13, round(ann, 2), _return_fmt(wb, ann, bg))
        else:
            ws.write_blank(row, 13, None, fmt["tbl_empty"])
        row += 1

    # ── Footer ─────────────────────────────────────────────────────────────────
    row += 1
    ws.set_row(row, 28)
    ws.merge_range(row, 0, row, 21, LEGAL_NOTE, fmt["footer"])
    row += 1
    ws.set_row(row, 14)
    ws.merge_range(row, 0, row, 21,
                   f"  {MANAGER}   Phone: {PHONE} | {EMAIL}",
                   fmt["footer_contact"])

    wb.close()
    print(f"✓ Saved: {output_path}")


# ─── Chart data ────────────────────────────────────────────────────────────────

def _write_chart_data(wb, ds, df: pd.DataFrame, stats: dict):
    ds.write(0, 0, "Date")
    ds.write(0, 1, "VAMI")
    for i, (d, v) in enumerate(zip(stats["vami_dates"], stats["vami"]), start=1):
        ds.write_datetime(i, 0, d.to_pydatetime(),
                          wb.add_format({"num_format": "mmm yyyy"}))
        ds.write(i, 1, round(v, 2))

    ds.write(0, 3, "Date")
    ds.write(0, 4, "Monthly Return (%)")
    for i, row in df.iterrows():
        ds.write_datetime(i + 1, 3, row["month"].to_pydatetime(),
                          wb.add_format({"num_format": "mmm yyyy"}))
        ds.write(i + 1, 4, round(row["return"] * 100, 4))

    ds.write(0, 6, "Date (str)")
    ds.write(0, 7, "VAMI")
    for i, (d, v) in enumerate(zip(stats["vami_dates"], stats["vami"]), start=1):
        ds.write(i, 6, d.strftime("%b %Y"))
        ds.write(i, 7, round(v, 2))

    ds.write(0, 9, "Date (str)")
    ds.write(0, 10, "Monthly Return (%)")
    for i, row in df.iterrows():
        ds.write(i + 1, 9, row["month"].strftime("%b %Y"))
        ds.write(i + 1, 10, round(row["return"] * 100, 4))


def _make_vami_chart(wb, ds, stats: dict):
    n = len(stats["vami"])
    chart = wb.add_chart({"type": "line"})
    chart.add_series({
        "name":       "VAMI",
        "categories": ["_data", 1, 6, n, 6],
        "values":     ["_data", 1, 7, n, 7],
        "line":       {"color": C_ACCENT, "width": 2},
        "marker":     {"type": "none"},
    })
    chart.set_title({"none": True})
    chart.set_x_axis({
        "name": "", "num_font": {"size": 7},
        "line": {"color": C_BORDER},
        "major_gridlines": {"visible": False},
    })
    chart.set_y_axis({
        "name": "Performance (VAMI)", "name_font": {"size": 8},
        "num_font": {"size": 7},
        "line": {"color": C_BORDER},
        "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0", "dash_type": "dash"}},
    })
    chart.set_legend({"none": True})
    chart.set_chartarea({"border": {"color": C_BORDER}, "fill": {"color": C_WHITE}})
    chart.set_plotarea({"fill": {"color": "#FAFAFA"}})
    chart.set_size({"width": 380, "height": 200})
    return chart


def _make_monthly_returns_chart(wb, ds, stats: dict):
    n = len(stats["all_returns"])
    chart = wb.add_chart({"type": "bar"})
    chart.add_series({
        "name":       "Monthly Return (%)",
        "categories": ["_data", 1, 9, n, 9],
        "values":     ["_data", 1, 10, n, 10],
        "fill":       {"color": C_ACCENT},
        "border":     {"color": C_ACCENT},
        "gap":        30,
    })
    chart.set_title({"none": True})
    chart.set_x_axis({
        "name": "", "num_font": {"size": 6},
        "line": {"color": C_BORDER},
        "major_gridlines": {"visible": False},
    })
    chart.set_y_axis({
        "name": "Monthly Return (%)", "name_font": {"size": 8},
        "num_font": {"size": 7},
        "line": {"color": C_BORDER},
        "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0", "dash_type": "dash"}},
        "num_format": "0.00%",
    })
    chart.set_legend({"none": True})
    chart.set_chartarea({"border": {"color": C_BORDER}, "fill": {"color": C_WHITE}})
    chart.set_plotarea({"fill": {"color": "#FAFAFA"}})
    chart.set_size({"width": 760, "height": 170})
    return chart


# ─── Formats ───────────────────────────────────────────────────────────────────

def _formats(wb) -> dict:
    def f(**kw):
        return wb.add_format(kw)

    return {
        "hdr_date": f(
            bold=True, font_size=13, font_color=C_HEADER_FG,
            bg_color=C_HEADER_BG, align="left", valign="vcenter",
            left=1, bottom=1, top=1, border_color=C_HEADER_BG,
        ),
        "hdr_title": f(
            bold=True, font_size=13, font_color=C_HEADER_FG,
            bg_color=C_HEADER_BG, align="center", valign="vcenter",
            right=1, bottom=1, top=1, border_color=C_HEADER_BG,
        ),
        "hdr_bar": f(bg_color=C_ACCENT),
        "sec_label": f(
            bold=True, font_size=9, font_color=C_SECTION_FG,
            bg_color=C_SECTION_BG, align="left", valign="vcenter",
            left=1, right=1, top=1, bottom=1, border_color=C_SECTION_BG,
        ),
        "sec_divider": f(bg_color=C_LABEL_BG, bottom=1, border_color=C_BORDER),
        "bullet": f(
            font_size=9, font_color="#333333", bg_color=C_LABEL_BG,
            align="left", valign="vcenter",
        ),
        "manager_name": f(
            bold=True, font_size=11, font_color=C_ACCENT,
            bg_color=C_WHITE, align="left", valign="vcenter",
        ),
        "body": f(font_size=9, bg_color=C_WHITE, valign="vcenter"),
        "body_small": f(
            font_size=8, font_color="#444444", bg_color=C_LABEL_BG,
            valign="vcenter", text_wrap=True,
        ),
        "stat_label": f(
            font_size=8, font_color="#666666", bg_color=C_WHITE,
            align="center", valign="vcenter", bottom=1, border_color=C_BORDER,
        ),
        "stat_big": f(
            bold=True, font_size=16, font_color=C_ACCENT,
            bg_color=C_WHITE, align="center", valign="vcenter",
            left=1, right=1, bottom=1, border_color=C_BORDER,
        ),
        "gi_label": f(
            font_size=9, font_color="#555555", bg_color=C_WHITE,
            align="left", valign="vcenter",
        ),
        "gi_value": f(
            bold=True, font_size=9, font_color="#222222",
            bg_color=C_WHITE, align="right", valign="vcenter",
        ),
        "stat_row_label": f(
            font_size=9, font_color="#555555", bg_color=C_WHITE,
            align="left", valign="vcenter", bottom=1, border_color="#EEEEEE",
        ),
        "stat_row_value": f(
            bold=True, font_size=9, font_color=C_ACCENT,
            bg_color=C_WHITE, align="right", valign="vcenter",
            bottom=1, border_color="#EEEEEE",
        ),
        "tbl_hdr": f(
            bold=True, font_size=9, font_color=C_SECTION_FG,
            bg_color=C_TABLE_HDR, align="center", valign="vcenter",
            left=1, right=1, top=1, bottom=1, border_color=C_TABLE_HDR,
        ),
        "tbl_year": f(
            bold=True, font_size=9, font_color="#222222",
            bg_color=C_LABEL_BG, align="center", valign="vcenter",
            left=1, right=1, top=1, bottom=1, border_color=C_BORDER,
        ),
        "tbl_empty": f(
            font_size=9, bg_color=C_WHITE,
            align="center", valign="vcenter",
            left=1, right=1, top=1, bottom=1, border_color=C_BORDER,
        ),
        "footer": f(
            font_size=7, font_color="#666666", bg_color=C_LIGHT_GREY,
            align="left", valign="vcenter", text_wrap=True,
            left=1, right=1, top=1, bottom=1, border_color=C_BORDER,
        ),
        "footer_contact": f(
            font_size=8, font_color="#333333", bg_color=C_LABEL_BG,
            align="left", valign="vcenter",
        ),
    }


_return_fmt_cache: dict = {}

def _return_fmt(wb, value: float, bg: str):
    key = (round(value, 4) > 0, bg)
    if key not in _return_fmt_cache:
        fg = C_POS if value >= 0 else C_NEG
        _return_fmt_cache[key] = wb.add_format({
            "font_size": 9, "font_color": fg, "bg_color": bg,
            "align": "center", "valign": "vcenter", "num_format": "0.00",
            "left": 1, "right": 1, "top": 1, "bottom": 1, "border_color": C_BORDER,
        })
    return _return_fmt_cache[key]


# ═══════════════════════════════════════════════════════════════════════════════
# Entry point
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    if not DATA_FILE.exists():
        print(f"ERROR: data file not found: {DATA_FILE}", file=sys.stderr)
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    print(f"Loading  {DATA_FILE.name} …")
    df = load_data(DATA_FILE)
    print(f"  {len(df)} monthly records  ({df['month'].iloc[0].strftime('%b %Y')} → "
          f"{df['month'].iloc[-1].strftime('%b %Y')})")

    stats = build_stats(df)
    latest = stats["latest_date"]
    fname  = f"Tear_Sheet_{latest.strftime('%b%Y')}.xlsx"
    out    = OUTPUT_DIR / fname

    print("Building tear sheet …")
    write_excel(df, stats, out)

    print("\nKey statistics:")
    print(f"  Total Return Cumulative : {stats['total_return']:.2f}%")
    print(f"  YTD                     : {stats['ytd']:.2f}%")
    print(f"  3-Month ROR             : {stats['three_mo_ror']:.2f}%")
    print(f"  6-Month ROR             : {stats['six_mo_ror']:.2f}%")
    print(f"  Sharpe Ratio            : {stats['sharpe']:.2f}")
    print(f"  Winning Months          : {stats['winning_pct']:.2f}%")


if __name__ == "__main__":
    main()
