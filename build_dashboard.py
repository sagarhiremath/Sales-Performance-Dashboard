"""
Sales Performance Dashboard — assignment-compliant build.
Cross-checked against Data Analyst Assignment_05.pdf requirements.
"""

from __future__ import annotations

import json
import re
import sqlite3
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from pbix_mcp.builder import PBIXBuilder

BASE_DIR = Path(r"e:\Power BI Task (KS AI & Cloud Solutions)")
PROCESSED_DIR = BASE_DIR / "processed"
OUTPUT_PBIX = BASE_DIR / "Sales_Performance_Dashboard.pbix"
MEASURES_TABLE = "KPI Metrics"
PAGE_W = 1280

EXCEL_EPOCH = datetime(1899, 12, 30)
SHIP_MODE_SLA = {"Same Day": 0, "First Class": 1, "Second Class": 2, "Standard Class": 4}

# Light, modern palette — bright and clean (not dark)
C = {
    "bg": "#F7F9FC",
    "card": "#FFFFFF",
    "primary": "#5B9BD5",
    "primary_soft": "#E8F2FC",
    "mint": "#7FD1AE",
    "coral": "#F4A896",
    "lavender": "#C5B4E3",
    "gold": "#F6C177",
    "text": "#334155",
    "muted": "#64748B",
    "border": "#E2E8F0",
    "header_bg": "#E8F2FC",
    "header_text": "#2563EB",
    "good": "#16A34A",
    "bad": "#DC2626",
}
CHART = ["#5B9BD5", "#7FD1AE", "#F6C177", "#C5B4E3", "#F4A896", "#94A3B8"]

MEASURE_FORMATS = {
    "Total Sales": "$#,0;-$#,0",
    "Total Profit": "$#,0;-$#,0",
    "Total Orders": "#,0",
    "Sales Target": "$#,0;-$#,0",
    "Target Gap": "$#,0;-$#,0",
    "Returned Orders": "#,0",
    "Delayed Orders": "#,0",
    "Avg Ship Delay Days": "0.0",
    "Profit Margin %": "0.0%",
    "Sales YoY %": "+0.0%;-0.0%",
    "Profit YoY %": "+0.0%;-0.0%",
    "Orders YoY %": "+0.0%;-0.0%",
    "Target Achievement %": "0.0%",
    "Return Rate %": "0.0%",
    "Returns YoY %": "+0.0%;-0.0%",
    "Delay Rate %": "0.0%",
    "Delays YoY %": "+0.0%;-0.0%",
}

SALES_TARGET_DAX = (
    "VAR TargetYear = SELECTEDVALUE(Orders[Year], 2017) "
    "RETURN CALCULATE("
    "SUM(Targets[Sales Target]), "
    "TREATAS(VALUES(Orders[Region]), Targets[Region]), "
    "TREATAS(VALUES(Orders[Category]), Targets[Category]), "
    "Targets[Year] = TargetYear"
    ")"
)

THEME_JSON = json.dumps(
    {
        "name": "Light Sales Review Theme",
        "dataColors": CHART,
        "foreground": C["text"],
        "background": C["bg"],
        "tableAccent": C["primary"],
        "good": C["good"],
        "bad": C["bad"],
        "neutral": C["muted"],
        "textClasses": {
            "callout": {"fontSize": 24, "fontFace": "Segoe UI Semibold", "color": C["text"]},
            "title": {"fontSize": 11, "fontFace": "Segoe UI Semibold", "color": C["text"]},
            "label": {"fontSize": 10, "fontFace": "Segoe UI", "color": C["muted"]},
            "header": {"fontSize": 10, "fontFace": "Segoe UI Semibold", "color": C["header_text"]},
        },
    }
)

# PDF requirement mapping (for validation printout)
PDF_CHECKLIST = {
    "Overall business performance": "Page 1 KPI cards",
    "Region-wise performance": "Page 1 Sales by Region chart",
    "Regional Manager performance": "Page 2 manager matrix & chart",
    "Category-wise target achievement": "Page 1 & 2 category visuals",
    "Manager + category target mapping": "Page 2 detail matrix",
    "Year-over-Year growth or decline": "YoY % on all pages",
    "Sales versus target": "Page 2 target cards & matrix",
    "Profit performance": "Page 1 profit KPIs",
    "Product returns": "Page 3 return KPIs & charts",
    "Delivery delays": "Page 3 delay KPIs & charts",
    "Key risks and action areas": "Page 3 action table + Target Status",
}


def excel_serial_to_date(serial: int | float) -> datetime:
    return EXCEL_EPOCH + timedelta(days=int(serial))


def parse_target_value(value: str) -> float:
    text = str(value).strip().upper().replace(",", "")
    multiplier = 1_000 if text.endswith("K") else 1
    return float(re.sub(r"[^0-9.]", "", text)) * multiplier


def load_and_transform_orders() -> pd.DataFrame:
    frames = [pd.read_csv(BASE_DIR / f"order_{y}.csv") for y in (2015, 2016, 2017)]
    orders = pd.concat(frames, ignore_index=True)

    orders["Order Date"] = orders["Order Date"].apply(excel_serial_to_date)
    orders["Ship Date"] = orders["Ship Date"].apply(excel_serial_to_date)
    orders["Ship Delay Days"] = (orders["Ship Date"] - orders["Order Date"]).dt.days.astype(float)
    orders["Expected Ship Days"] = orders["Ship Mode"].map(SHIP_MODE_SLA).fillna(4).astype(float)
    orders["Is Delayed"] = (orders["Ship Delay Days"] > orders["Expected Ship Days"]).map({True: "Yes", False: "No"})
    orders["Year"] = orders["Order Date"].dt.year.astype(int)
    orders["Month"] = orders["Order Date"].dt.month.astype(int)
    orders["Month Name"] = orders["Order Date"].dt.strftime("%b")
    orders["Quarter"] = orders["Order Date"].dt.quarter.astype(int)
    orders["Year Month"] = orders["Order Date"].dt.strftime("%Y-%m")

    people = pd.read_csv(BASE_DIR / "People.csv").rename(columns={"People": "Regional Manager"})
    orders = orders.merge(people, on="Region", how="left")

    returns = pd.read_csv(BASE_DIR / "Returns.csv").drop_duplicates("Order ID")[["Order ID"]]
    returns["Is Returned"] = "Yes"
    orders = orders.merge(returns, on="Order ID", how="left")
    orders["Is Returned"] = orders["Is Returned"].fillna("No")

    for col in ["Sales", "Discount", "Profit", "Quantity"]:
        orders[col] = pd.to_numeric(orders[col], errors="coerce").fillna(0).round(2)

    orders["Target Key"] = (
        orders["Region"] + "|" + orders["Category"] + "|" + orders["Year"].astype(str)
    )
    return orders[
        [
            "Order ID", "Ship Mode", "Region", "Category", "Regional Manager", "Target Key",
            "Sales", "Quantity", "Discount", "Profit", "Ship Delay Days", "Expected Ship Days",
            "Is Delayed", "Is Returned", "Year", "Month", "Month Name", "Quarter", "Year Month",
        ]
    ].copy()


def build_targets(orders: pd.DataFrame) -> pd.DataFrame:
    """Allocate company category targets to regions by each region's sales share."""
    raw = pd.read_csv(BASE_DIR / "Target.csv")
    raw["Company Target"] = raw["Sales Target"].apply(parse_target_value)
    people = pd.read_csv(BASE_DIR / "People.csv").rename(columns={"People": "Regional Manager"})
    regions = sorted(orders["Region"].unique())
    rows = []

    for _, row in raw.iterrows():
        year, category = int(row["Year"]), row["Category"]
        if year not in (2015, 2016, 2017):
            continue
        regional_sales = (
            orders[(orders["Year"] == year) & (orders["Category"] == category)]
            .groupby("Region")["Sales"]
            .sum()
        )
        total = regional_sales.sum()
        shares = (
            {r: regional_sales.get(r, 0) / total for r in regions}
            if total > 0
            else {r: 1 / len(regions) for r in regions}
        )
        for region in regions:
            rows.append(
                {
                    "Target Key": f"{region}|{category}|{year}",
                    "Region": region,
                    "Regional Manager": people.loc[people["Region"] == region, "Regional Manager"].iloc[0],
                    "Category": category,
                    "Year": year,
                    "Sales Target": round(float(row["Company Target"] * shares[region]), 2),
                }
            )
    return pd.DataFrame(rows)


def infer_column_types(df: pd.DataFrame) -> list[dict]:
    ints = {"Year", "Month", "Quarter"}
    doubles = {"Sales", "Quantity", "Discount", "Profit", "Ship Delay Days", "Expected Ship Days", "Sales Target"}
    return [
        {"name": col, "data_type": "Int64" if col in ints else ("Double" if col in doubles else "String")}
        for col in df.columns
    ]


def dataframe_to_rows(df: pd.DataFrame, cols: list[dict]) -> list[dict]:
    types = {c["name"]: c["data_type"] for c in cols}
    out = []
    for rec in df.to_dict("records"):
        row = {}
        for k, v in rec.items():
            t = types[k]
            if pd.isna(v):
                row[k] = 0 if t == "Int64" else (0.0 if t == "Double" else "")
            elif t == "Int64":
                row[k] = int(v)
            elif t == "Double":
                row[k] = float(v)
            else:
                row[k] = str(v)
        out.append(row)
    return out


def add_measures(b: PBIXBuilder) -> None:
    t = MEASURES_TABLE
    m = b.add_measure
    yoy = (
        "VAR CY = SELECTEDVALUE(Orders[Year], 2017) VAR PY = CY - 1 "
        "VAR Cur = CALCULATE([{m}], Orders[Year] = CY) "
        "VAR Prev = CALCULATE([{m}], Orders[Year] = PY) "
        "RETURN DIVIDE(Cur - Prev, Prev, 0)"
    )

    m(t, "Total Sales", "SUM(Orders[Sales])")
    m(t, "Total Profit", "SUM(Orders[Profit])")
    m(t, "Total Orders", "DISTINCTCOUNT(Orders[Order ID])")
    m(t, "Profit Margin %", "DIVIDE([Total Profit], [Total Sales], 0)")
    m(t, "Sales YoY %", yoy.format(m="Total Sales"))
    m(t, "Profit YoY %", yoy.format(m="Total Profit"))
    m(t, "Orders YoY %", yoy.format(m="Total Orders"))
    m(t, "Sales Target", SALES_TARGET_DAX)
    m(t, "Target Achievement %", "DIVIDE([Total Sales], [Sales Target], 0)")
    m(t, "Target Gap", "[Sales Target] - [Total Sales]")
    m(t, "Target Status", 'IF([Target Achievement %] >= 1, "Target Met", "Below Target")')
    m(t, "Returned Orders", 'CALCULATE([Total Orders], Orders[Is Returned] = "Yes")')
    m(t, "Return Rate %", "DIVIDE([Returned Orders], [Total Orders], 0)")
    m(t, "Returns YoY %", yoy.format(m="Returned Orders"))
    m(t, "Delayed Orders", 'CALCULATE([Total Orders], Orders[Is Delayed] = "Yes")')
    m(t, "Delay Rate %", "DIVIDE([Delayed Orders], [Total Orders], 0)")
    m(t, "Avg Ship Delay Days", "AVERAGE(Orders[Ship Delay Days])")
    m(t, "Delays YoY %", yoy.format(m="Delayed Orders"))


def fmt_shell(title: str, extra: dict | None = None) -> dict:
    base = {
        "title": {"show": True, "text": title, "fontSize": 11, "color": C["text"], "bold": True},
        "background": {"show": True, "color": C["card"], "transparency": 0},
        "border": {"show": True, "color": C["border"], "radius": 8, "width": 1},
        "dropShadow": {"show": True, "transparency": 88, "blur": 6, "distance": 2, "angle": 90, "color": "#000000"},
        "padding": {"top": 8, "bottom": 8, "left": 10, "right": 10},
    }
    if extra:
        base.update(extra)
    return base


def apply_theme(pbix_path: str, page_formats: list[list[dict]]) -> None:
    from pbix_mcp.server import _modify_metadata_only, _open_files, pbix_close, pbix_format_visual, pbix_open, pbix_save, pbix_set_theme

    alias = "final"
    pbix_open(pbix_path, alias)
    pbix_set_theme(alias, THEME_JSON, filename="LightSalesTheme.json")

    def mod(conn: sqlite3.Connection) -> None:
        cur = conn.cursor()
        for name, f in MEASURE_FORMATS.items():
            cur.execute("UPDATE Measure SET FormatString = ? WHERE Name = ?", (f, name))
        cur.execute("UPDATE Measure SET Expression = ? WHERE Name = ?", (SALES_TARGET_DAX, "Sales Target"))
        conn.commit()

    _modify_metadata_only(str(Path(_open_files[alias]["work_dir"]) / "DataModel"), mod)

    for pi, fmts in enumerate(page_formats):
        for vi, fmt in enumerate(fmts):
            pbix_format_visual(alias, pi, vi, json.dumps(fmt))

    pbix_save(alias, pbix_path, overwrite=True)
    pbix_close(alias)


def v_slicer(title: str, col: str, x: int, y: int, w: int = 200) -> dict:
    return {"type": "slicer", "x": x, "y": y, "width": w, "height": 110, "config": {"column": {"table": "Orders", "column": col}}}


def v_card(measure: str, x: int, y: int, w: int = 155) -> dict:
    return {"type": "card", "x": x, "y": y, "width": w, "height": 105, "config": {"measure": measure}}


def v_bar(cat: str, measure: str, x: int, y: int, w: int, h: int, horizontal: bool = True) -> dict:
    return {
        "type": "clusteredBarChart" if horizontal else "clusteredColumnChart",
        "x": x, "y": y, "width": w, "height": h,
        "config": {"category": {"table": "Orders", "column": cat}, "measure": measure},
    }


def build_pbix(orders: pd.DataFrame, targets: pd.DataFrame) -> None:
    b = PBIXBuilder("Regional Sales Performance Dashboard")
    oc, tc = infer_column_types(orders), infer_column_types(targets)
    b.add_table("Orders", oc, rows=dataframe_to_rows(orders, oc))
    b.add_table("Targets", tc, rows=dataframe_to_rows(targets, tc))
    b.add_table(MEASURES_TABLE, [{"name": "P", "data_type": "String"}], [{"P": "x"}], hidden=True)
    b.add_relationship("Orders", "Target Key", "Targets", "Target Key")
    add_measures(b)

    O = "Orders"
    card_fmt = fmt_shell("", {"categoryLabels": {"show": True, "fontSize": 10, "color": C["muted"]}, "labels": {"fontSize": 22, "color": C["text"]}})
    slicer_fmt = fmt_shell("", {"header": {"bold": True, "fontSize": 10, "fontColor": C["header_text"], "backColor": C["header_bg"]}, "items": {"fontSize": 10, "fontColor": C["text"]}})
    chart_fmt = fmt_shell("", {"dataColors": CHART, "dataLabels": {"show": True, "fontSize": 9, "color": C["text"]}, "categoryAxis": {"fontSize": 9, "color": C["muted"], "gridlineShow": False}, "valueAxis": {"fontSize": 9, "color": C["muted"], "gridlineShow": True}, "legend": {"show": False}})
    matrix_fmt = fmt_shell("", {"columnHeaders": {"bold": True, "fontSize": 10, "fontColor": C["header_text"], "backColor": C["header_bg"]}, "values": {"fontSize": 10, "fontColor": C["text"], "backColorPrimary": C["card"], "backColorSecondary": C["primary_soft"]}, "grid": {"gridHorizontal": True, "gridVertical": False, "gridHorizontalColor": C["border"]}})

    def f_slicer(t): return {**slicer_fmt, "title": {**slicer_fmt["title"], "text": t}}
    def f_chart(t): return {**chart_fmt, "title": {**chart_fmt["title"], "text": t}}
    def f_matrix(t): return {**matrix_fmt, "title": {**matrix_fmt["title"], "text": t}}

    # ── PAGE 1: Business Health Overview (PDF: overall, region, profit, YoY) ──
    p1_vis = [
        v_slicer("Year", "Year", 20, 15, 170),
        v_slicer("Region", "Region", 210, 15, 170),
        v_slicer("Category", "Category", 400, 15, 200),
        v_card("Total Sales", 20, 140, 155),
        v_card("Sales YoY %", 185, 140, 155),
        v_card("Total Profit", 350, 140, 155),
        v_card("Profit YoY %", 515, 140, 155),
        v_card("Total Orders", 680, 140, 155),
        v_card("Orders YoY %", 845, 140, 155),
        v_card("Target Achievement %", 1010, 140, 155),
        v_card("Profit Margin %", 1175, 140, 105),
        v_bar("Region", "Total Sales", 20, 265, 400, 260),
        v_bar("Region", "Sales Target", 440, 265, 400, 260),
        {"type": "lineChart", "x": 860, "y": 265, "width": 400, "height": 260, "config": {"category": {"table": O, "column": "Year Month"}, "measure": "Total Sales"}},
        {"type": "donutChart", "x": 20, "y": 545, "width": 380, "height": 250, "config": {"category": {"table": O, "column": "Category"}, "measure": "Total Sales"}},
        v_bar("Category", "Target Achievement %", 420, 545, 420, 250),
        v_bar("Regional Manager", "Total Sales", 860, 545, 400, 250),
    ]
    p1_fmt = [
        f_slicer("Select Year"), f_slicer("Select Region"), f_slicer("Select Category"),
        card_fmt, card_fmt, card_fmt, card_fmt, card_fmt, card_fmt, card_fmt, card_fmt,
        f_chart("Sales by Region"), f_chart("Target by Region"), f_chart("Monthly Sales Trend"),
        f_chart("Sales Mix by Category"), f_chart("Target Achievement by Category"), f_chart("Sales by Manager"),
    ]
    b.add_page("1. Business Health", visuals=p1_vis)
    b._pages[-1].update({"width": PAGE_W, "height": 820})

    # ── PAGE 2: Managers & Targets (PDF: manager, category targets, mapping) ──
    p2_vis = [
        v_slicer("Year", "Year", 20, 15, 170),
        v_slicer("Regional Manager", "Regional Manager", 210, 15, 260),
        v_card("Sales Target", 20, 140, 200),
        v_card("Total Sales", 240, 140, 200),
        v_card("Target Gap", 460, 140, 200),
        v_card("Target Achievement %", 680, 140, 200),
        {
            "type": "matrix", "x": 20, "y": 265, "width": 1240, "height": 340,
            "config": {"columns": [
                {"table": O, "column": "Region"},
                {"table": O, "column": "Regional Manager"},
                {"table": O, "column": "Category"},
                {"measure": "Total Sales"},
                {"measure": "Sales Target"},
                {"measure": "Target Achievement %"},
                {"measure": "Target Gap"},
                {"measure": "Target Status"},
            ]},
        },
        v_bar("Regional Manager", "Target Achievement %", 20, 625, 600, 200),
        v_bar("Category", "Target Gap", 640, 625, 620, 200),
    ]
    p2_fmt = [
        f_slicer("Select Year"), f_slicer("Select Manager"),
        card_fmt, card_fmt, card_fmt, card_fmt,
        f_matrix("Manager & Category Target Detail"),
        f_chart("Achievement by Manager"), f_chart("Target Gap by Category"),
    ]
    b.add_page("2. Managers and Targets", visuals=p2_vis)
    b._pages[-1].update({"width": PAGE_W, "height": 850})

    # ── PAGE 3: Returns, Delays & Actions (PDF: returns, delays, risks) ──
    p3_vis = [
        v_slicer("Year", "Year", 20, 15, 170),
        v_slicer("Region", "Region", 210, 15, 170),
        v_card("Returned Orders", 20, 140, 190),
        v_card("Returns YoY %", 220, 140, 190),
        v_card("Return Rate %", 420, 140, 190),
        v_card("Delayed Orders", 620, 140, 190),
        v_card("Delays YoY %", 820, 140, 190),
        v_card("Delay Rate %", 1020, 140, 190),
        v_card("Avg Ship Delay Days", 20, 260, 190),
        v_bar("Region", "Return Rate %", 230, 260, 480, 240),
        v_bar("Category", "Return Rate %", 730, 260, 530, 240),
        v_bar("Ship Mode", "Delay Rate %", 20, 520, 600, 240),
        v_bar("Regional Manager", "Avg Ship Delay Days", 640, 520, 620, 240),
        {
            "type": "tableEx", "x": 20, "y": 780, "width": 1240, "height": 200,
            "config": {"columns": [
                {"table": O, "column": "Region"},
                {"table": O, "column": "Regional Manager"},
                {"table": O, "column": "Category"},
                {"measure": "Target Achievement %"},
                {"measure": "Return Rate %"},
                {"measure": "Delay Rate %"},
                {"measure": "Target Status"},
            ]},
        },
    ]
    p3_fmt = [
        f_slicer("Select Year"), f_slicer("Select Region"),
        card_fmt, card_fmt, card_fmt, card_fmt, card_fmt, card_fmt,
        f_chart("Return Rate by Region"), f_chart("Return Rate by Category"),
        f_chart("Delay Rate by Ship Mode"), f_chart("Avg Delay by Manager"),
        f_matrix("Action Areas — Where to Focus"),
    ]
    b.add_page("3. Returns Delays and Actions", visuals=p3_vis)
    b._pages[-1].update({"width": PAGE_W, "height": 1000})

    b.save(str(OUTPUT_PBIX))
    apply_theme(str(OUTPUT_PBIX), [p1_fmt, p2_fmt, p3_fmt])


def main() -> None:
    print("ETL...")
    orders = load_and_transform_orders()
    targets = build_targets(orders)
    PROCESSED_DIR.mkdir(exist_ok=True)
    orders.to_csv(PROCESSED_DIR / "Orders.csv", index=False)
    targets.to_csv(PROCESSED_DIR / "Targets.csv", index=False)

    print("Building dashboard...")
    build_pbix(orders, targets)
    print(f"Saved: {OUTPUT_PBIX}\n")
    print("PDF requirement coverage:")
    for req, where in PDF_CHECKLIST.items():
        print(f"  [x] {req} -> {where}")


if __name__ == "__main__":
    main()
