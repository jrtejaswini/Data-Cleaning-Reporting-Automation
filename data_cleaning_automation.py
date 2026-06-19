"""
================================================================
             Data Cleaning & Reporting Automation
================================================================
Automates data cleaning and generates visual summary reports
from raw sales data with common quality issues.

Key Features:
  - Handles missing values, duplicates, inconsistent formatting
  - Auto-detects and standardizes date formats
  - Generates automated HTML report with charts and stats
  - Exports cleaned data to Excel/CSV

Author  : Tejas
================================================================
"""

import pandas as pd
import numpy as np
import os
import json
import warnings
from datetime import datetime
from collections import defaultdict

warnings.filterwarnings("ignore")


INPUT_FILE   = "sample_dirty_data.csv"
CLEAN_CSV    = "cleaned_data.csv"
CLEAN_EXCEL  = "cleaned_data.xlsx"
REPORT_HTML  = "data_report.html"


VALID_STATUSES = {"Completed", "Pending", "Cancelled"}
VALID_REGIONS  = {"Luzon", "Visayas", "Mindanao"}


def load_data(filepath: str) -> pd.DataFrame:
    """Load CSV and record raw shape."""
    df = pd.read_csv(filepath)
    print(f"\n[LOAD] Loaded {len(df)} rows × {len(df.columns)} cols from '{filepath}'")
    return df


def audit_data(df: pd.DataFrame) -> dict:
    """Return a dict summarising data quality issues."""
    report = {}
    report["total_rows"]       = len(df)
    report["total_columns"]    = len(df.columns)
    report["duplicate_rows"]   = int(df.duplicated().sum())
    report["missing_per_col"]  = df.isnull().sum().to_dict()
    report["missing_total"]    = int(df.isnull().sum().sum())

    
    region_raw   = df["Region"].dropna().unique().tolist()
    category_raw = df["Category"].dropna().unique().tolist()
    report["region_variants"]   = region_raw
    report["category_variants"] = category_raw

    
    sample_dates = df["OrderDate"].dropna().unique().tolist()
    report["date_samples"] = sample_dates[:6]

    print("\n[AUDIT] Data Quality Summary:")
    print(f"  Total rows       : {report['total_rows']}")
    print(f"  Duplicate rows   : {report['duplicate_rows']}")
    print(f"  Missing values   : {report['missing_total']}")
    print(f"  Region variants  : {region_raw}")
    print(f"  Category variants: {category_raw}")

    return report


def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Apply cleaning steps and return (cleaned_df, cleaning_log).
    """
    log = defaultdict(int)
    df  = df.copy()

    
    before = len(df)
    df.drop_duplicates(inplace=True)
    log["duplicates_removed"] = before - len(df)
    print(f"\n[CLEAN] Duplicates removed : {log['duplicates_removed']}")

    
    str_cols = df.select_dtypes(include="object").columns
    for col in str_cols:
        df[col] = df[col].str.strip()

    
    df["Region"]       = df["Region"].str.title()
    df["Category"]     = df["Category"].str.title()
    df["Status"]       = df["Status"].str.title()
    df["CustomerName"] = df["CustomerName"].str.title()
    log["casing_standardised"] = len(str_cols)
    print(f"[CLEAN] Casing standardised for {len(str_cols)} columns")


    invalid_region = ~df["Region"].isin(VALID_REGIONS) & df["Region"].notna()
    log["invalid_regions"] = int(invalid_region.sum())
    df.loc[invalid_region, "Region"] = np.nan   

    invalid_status = ~df["Status"].isin(VALID_STATUSES) & df["Status"].notna()
    log["invalid_statuses"] = int(invalid_status.sum())
    df.loc[invalid_status, "Status"] = np.nan

    
    def parse_date(val):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(str(val).strip(), fmt)
            except (ValueError, TypeError):
                continue
        return pd.NaT

    before_nulls = df["OrderDate"].isnull().sum()
    df["OrderDate"] = df["OrderDate"].apply(parse_date)
    after_nulls     = df["OrderDate"].isnull().sum()
    log["dates_fixed"] = int(after_nulls - before_nulls) if after_nulls < before_nulls else 0
    print(f"[CLEAN] Date column parsed  — unparseable remaining: {int(after_nulls)}")

    
    qty_missing = int(df["Quantity"].isnull().sum())
    df["Quantity"] = df.groupby("Category")["Quantity"].transform(
        lambda x: x.fillna(x.median())
    )
    df["Quantity"] = df["Quantity"].round(0).astype("Int64")
    log["quantity_imputed"] = qty_missing
    print(f"[CLEAN] Quantity imputed    : {qty_missing} values (category median)")

    
    price_missing = int(df["UnitPrice"].isnull().sum())
    df["UnitPrice"] = df.groupby("Product")["UnitPrice"].transform(
        lambda x: x.fillna(x.median())
    )
    log["unitprice_imputed"] = price_missing
    print(f"[CLEAN] UnitPrice imputed   : {price_missing} values (product median)")

    
    name_missing = int(df["CustomerName"].isnull().sum())
    df["CustomerName"].fillna("Unknown", inplace=True)
    log["names_filled"] = name_missing

    
    df["Revenue"] = df["Quantity"].astype(float) * df["UnitPrice"]


    df.reset_index(drop=True, inplace=True)

    print(f"\n[CLEAN] Done — {len(df)} clean rows remain")
    return df, dict(log)


def build_aggregates(df: pd.DataFrame) -> dict:
    """Compute summary stats used in the HTML report."""
    agg = {}

    agg["total_orders"]   = len(df)
    agg["total_revenue"]  = float(df["Revenue"].sum())
    agg["avg_order_val"]  = float(df["Revenue"].mean())
    agg["total_units"]    = int(df["Quantity"].sum())

    
    agg["revenue_region"] = (
        df.groupby("Region")["Revenue"].sum()
          .sort_values(ascending=False)
          .to_dict()
    )


    agg["revenue_category"] = (
        df.groupby("Category")["Revenue"].sum()
          .sort_values(ascending=False)
          .to_dict()
    )

    
    agg["orders_status"] = df["Status"].value_counts().to_dict()

    
    agg["top_products"] = (
        df.groupby("Product")["Revenue"].sum()
          .sort_values(ascending=False)
          .head(5)
          .to_dict()
    )

    
    daily = (
        df.groupby(df["OrderDate"].dt.date)["Revenue"]
          .sum()
          .reset_index()
    )
    daily.columns = ["date", "revenue"]
    daily["date"] = daily["date"].astype(str)
    agg["daily_revenue"] = daily.to_dict(orient="records")

    return agg


def generate_report(
    raw_audit:    dict,
    clean_log:    dict,
    agg:          dict,
    output_path:  str
) -> None:
    """Write a self-contained HTML report with embedded Chart.js visuals."""

    ts = datetime.now().strftime("%B %d, %Y — %I:%M %p")

    # Helper: format currency PHP
    def php(val): return f"₱{val:,.2f}"

    
    region_labels  = json.dumps(list(agg["revenue_region"].keys()))
    region_values  = json.dumps([round(v, 2) for v in agg["revenue_region"].values()])

    category_labels = json.dumps(list(agg["revenue_category"].keys()))
    category_values = json.dumps([round(v, 2) for v in agg["revenue_category"].values()])

    status_labels  = json.dumps(list(agg["orders_status"].keys()))
    status_values  = json.dumps(list(agg["orders_status"].values()))

    product_labels = json.dumps(list(agg["top_products"].keys()))
    product_values = json.dumps([round(v, 2) for v in agg["top_products"].values()])

    daily_labels   = json.dumps([r["date"] for r in agg["daily_revenue"]])
    daily_values   = json.dumps([round(r["revenue"], 2) for r in agg["daily_revenue"]])

    
    def log_row(label, val, badge_class="badge-info"):
        return f'<tr><td>{label}</td><td><span class="badge {badge_class}">{val}</span></td></tr>'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Data Cleaning & Reporting — Thiranex Project 4</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg:       #0f1117;
    --surface:  #1a1d27;
    --card:     #21253a;
    --accent:   #4f8ef7;
    --accent2:  #38d9a9;
    --accent3:  #f7c948;
    --danger:   #f76e6e;
    --text:     #e2e8f0;
    --muted:    #8892aa;
    --border:   #2d3352;
    --radius:   12px;
    --font:     'Segoe UI', system-ui, sans-serif;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: var(--bg);
    color: var(--text);
    font-family: var(--font);
    line-height: 1.6;
  }}

  /* ── Header ── */
  header {{
    background: linear-gradient(135deg, #131729 0%, #1c2444 100%);
    border-bottom: 1px solid var(--border);
    padding: 32px 40px 28px;
  }}
  header .eyebrow {{
    font-size: 11px;
    letter-spacing: 2.5px;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 8px;
  }}
  header h1 {{
    font-size: 28px;
    font-weight: 700;
    color: #fff;
    margin-bottom: 6px;
  }}
  header .meta {{
    font-size: 13px;
    color: var(--muted);
  }}

  /* ── Layout ── */
  main {{ padding: 32px 40px; max-width: 1200px; margin: 0 auto; }}
  section {{ margin-bottom: 40px; }}
  h2 {{
    font-size: 13px;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 16px;
    padding-bottom: 8px;
    border-bottom: 1px solid var(--border);
  }}

  /* ── KPI Cards ── */
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 16px;
    margin-bottom: 0;
  }}
  .kpi-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px 22px;
    position: relative;
    overflow: hidden;
  }}
  .kpi-card::before {{
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: var(--accent);
  }}
  .kpi-card.green::before  {{ background: var(--accent2); }}
  .kpi-card.yellow::before {{ background: var(--accent3); }}
  .kpi-card.red::before    {{ background: var(--danger);  }}
  .kpi-label {{ font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 1.5px; margin-bottom: 8px; }}
  .kpi-value {{ font-size: 26px; font-weight: 700; color: #fff; }}
  .kpi-sub   {{ font-size: 12px; color: var(--muted); margin-top: 4px; }}

  /* ── Chart Grid ── */
  .chart-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 20px;
  }}
  .chart-grid.three {{ grid-template-columns: 1fr 1fr 1fr; }}
  @media (max-width: 900px) {{
    .chart-grid, .chart-grid.three {{ grid-template-columns: 1fr; }}
  }}
  .chart-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px 22px;
  }}
  .chart-card h3 {{
    font-size: 13px;
    font-weight: 600;
    color: var(--text);
    margin-bottom: 16px;
  }}
  .chart-card canvas {{ width: 100% !important; }}

  /* ── Cleaning Log Table ── */
  .log-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 14px;
  }}
  .log-table th, .log-table td {{
    padding: 10px 14px;
    text-align: left;
    border-bottom: 1px solid var(--border);
  }}
  .log-table th {{
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--muted);
  }}
  .log-table tr:last-child td {{ border-bottom: none; }}
  .badge {{
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 600;
  }}
  .badge-info    {{ background: rgba(79,142,247,.18); color: var(--accent);  }}
  .badge-success {{ background: rgba(56,217,169,.18); color: var(--accent2); }}
  .badge-warn    {{ background: rgba(247,201,72,.18); color: var(--accent3); }}
  .badge-danger  {{ background: rgba(247,110,110,.18); color: var(--danger); }}

  /* ── Issues Panel ── */
  .issue-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 12px;
  }}
  .issue-item {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 14px 16px;
    display: flex;
    align-items: center;
    gap: 12px;
  }}
  .issue-icon {{
    font-size: 22px;
    flex-shrink: 0;
  }}
  .issue-label {{ font-size: 13px; color: var(--muted); }}
  .issue-count {{ font-size: 20px; font-weight: 700; color: #fff; }}

  footer {{
    text-align: center;
    padding: 24px 40px;
    font-size: 12px;
    color: var(--muted);
    border-top: 1px solid var(--border);
    margin-top: 20px;
  }}
</style>
</head>
<body>

<header>
  <p class="eyebrow">Thiranex · Project 4 · Data Cleaning &amp; Reporting Automation</p>
  <h1>Sales Data Report</h1>
  <p class="meta">Generated: {ts} &nbsp;|&nbsp; Source: {INPUT_FILE} &nbsp;|&nbsp; Cleaned records: {agg['total_orders']}</p>
</header>

<main>

  <!-- KPIs -->
  <section>
    <h2>Key Metrics — Cleaned Dataset</h2>
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-label">Total Orders</div>
        <div class="kpi-value">{agg['total_orders']}</div>
      </div>
      <div class="kpi-card green">
        <div class="kpi-label">Total Revenue</div>
        <div class="kpi-value">{php(agg['total_revenue'])}</div>
      </div>
      <div class="kpi-card yellow">
        <div class="kpi-label">Avg Order Value</div>
        <div class="kpi-value">{php(agg['avg_order_val'])}</div>
      </div>
      <div class="kpi-card red">
        <div class="kpi-label">Total Units Sold</div>
        <div class="kpi-value">{agg['total_units']:,}</div>
      </div>
    </div>
  </section>

  <!-- Data Quality Audit -->
  <section>
    <h2>Data Quality Issues Found</h2>
    <div class="issue-grid">
      <div class="issue-item">
        <span class="issue-icon">📋</span>
        <div>
          <div class="issue-label">Raw Rows Loaded</div>
          <div class="issue-count">{raw_audit['total_rows']}</div>
        </div>
      </div>
      <div class="issue-item">
        <span class="issue-icon">🔁</span>
        <div>
          <div class="issue-label">Duplicate Rows</div>
          <div class="issue-count">{raw_audit['duplicate_rows']}</div>
        </div>
      </div>
      <div class="issue-item">
        <span class="issue-icon">❓</span>
        <div>
          <div class="issue-label">Missing Values</div>
          <div class="issue-count">{raw_audit['missing_total']}</div>
        </div>
      </div>
      <div class="issue-item">
        <span class="issue-icon">📅</span>
        <div>
          <div class="issue-label">Date Format Variants</div>
          <div class="issue-count">2</div>
        </div>
      </div>
      <div class="issue-item">
        <span class="issue-icon">🔡</span>
        <div>
          <div class="issue-label">Region Casing Issues</div>
          <div class="issue-count">{len(raw_audit['region_variants'])}</div>
        </div>
      </div>
      <div class="issue-item">
        <span class="issue-icon">🏷️</span>
        <div>
          <div class="issue-label">Category Variants</div>
          <div class="issue-count">{len(raw_audit['category_variants'])}</div>
        </div>
      </div>
    </div>
  </section>

  <!-- Cleaning Log -->
  <section>
    <h2>Cleaning Steps Applied</h2>
    <div class="chart-card">
      <table class="log-table">
        <thead>
          <tr><th>Cleaning Action</th><th>Records Affected</th></tr>
        </thead>
        <tbody>
          {log_row("Duplicate rows removed",             clean_log.get("duplicates_removed", 0), "badge-danger")}
          {log_row("Text columns casing standardised",   clean_log.get("casing_standardised", 0), "badge-info")}
          {log_row("Date formats normalised (→ YYYY-MM-DD)", "All", "badge-info")}
          {log_row("Quantity imputed (category median)", clean_log.get("quantity_imputed", 0),   "badge-warn")}
          {log_row("UnitPrice imputed (product median)", clean_log.get("unitprice_imputed", 0),  "badge-warn")}
          {log_row("Missing CustomerName → 'Unknown'",   clean_log.get("names_filled", 0),       "badge-warn")}
          {log_row("Revenue column derived",              agg['total_orders'],                    "badge-success")}
        </tbody>
      </table>
    </div>
  </section>

  <!-- Charts Row 1 -->
  <section>
    <h2>Revenue Analysis</h2>
    <div class="chart-grid">
      <div class="chart-card">
        <h3>Revenue by Region</h3>
        <canvas id="regionChart" height="220"></canvas>
      </div>
      <div class="chart-card">
        <h3>Revenue by Category</h3>
        <canvas id="categoryChart" height="220"></canvas>
      </div>
    </div>
  </section>

  <!-- Charts Row 2 -->
  <section>
    <h2>Orders &amp; Products</h2>
    <div class="chart-grid">
      <div class="chart-card">
        <h3>Order Status Breakdown</h3>
        <canvas id="statusChart" height="220"></canvas>
      </div>
      <div class="chart-card">
        <h3>Top 5 Products by Revenue</h3>
        <canvas id="productChart" height="220"></canvas>
      </div>
    </div>
  </section>

  <!-- Daily Trend -->
  <section>
    <h2>Daily Revenue Trend</h2>
    <div class="chart-card">
      <h3>Revenue Over Time</h3>
      <canvas id="trendChart" height="120"></canvas>
    </div>
  </section>

</main>

<footer>
  Data Cleaning &amp; Reporting Automation &nbsp;|&nbsp; Auto-generated report
</footer>

<script>
const COLORS = ['#4f8ef7','#38d9a9','#f7c948','#f76e6e','#b07ef7','#f79e4f'];
const GRID   = 'rgba(255,255,255,0.06)';
const TEXT   = '#8892aa';
const defaults = {{
  responsive: true,
  plugins: {{ legend: {{ labels: {{ color: TEXT, font: {{ size: 12 }} }} }} }},
}};

// Region Bar
new Chart(document.getElementById('regionChart'), {{
  type: 'bar',
  data: {{
    labels: {region_labels},
    datasets: [{{ label: 'Revenue (₱)', data: {region_values},
      backgroundColor: COLORS, borderRadius: 6, borderSkipped: false }}]
  }},
  options: {{ ...defaults, scales: {{
    x: {{ ticks: {{ color: TEXT }}, grid: {{ color: GRID }} }},
    y: {{ ticks: {{ color: TEXT, callback: v => '₱' + v.toLocaleString() }}, grid: {{ color: GRID }} }}
  }}, plugins: {{ ...defaults.plugins, legend: {{ display: false }} }} }}
}});

// Category Bar
new Chart(document.getElementById('categoryChart'), {{
  type: 'bar',
  data: {{
    labels: {category_labels},
    datasets: [{{ label: 'Revenue (₱)', data: {category_values},
      backgroundColor: ['#4f8ef7','#38d9a9'], borderRadius: 6, borderSkipped: false }}]
  }},
  options: {{ ...defaults, indexAxis: 'y', scales: {{
    x: {{ ticks: {{ color: TEXT, callback: v => '₱' + v.toLocaleString() }}, grid: {{ color: GRID }} }},
    y: {{ ticks: {{ color: TEXT }}, grid: {{ color: GRID }} }}
  }}, plugins: {{ ...defaults.plugins, legend: {{ display: false }} }} }}
}});

// Status Doughnut
new Chart(document.getElementById('statusChart'), {{
  type: 'doughnut',
  data: {{
    labels: {status_labels},
    datasets: [{{ data: {status_values}, backgroundColor: ['#38d9a9','#f7c948','#f76e6e'],
      borderWidth: 0, hoverOffset: 6 }}]
  }},
  options: {{ ...defaults, cutout: '65%' }}
}});

// Product Bar
new Chart(document.getElementById('productChart'), {{
  type: 'bar',
  data: {{
    labels: {product_labels},
    datasets: [{{ label: 'Revenue (₱)', data: {product_values},
      backgroundColor: COLORS, borderRadius: 6, borderSkipped: false }}]
  }},
  options: {{ ...defaults, scales: {{
    x: {{ ticks: {{ color: TEXT }}, grid: {{ color: GRID }} }},
    y: {{ ticks: {{ color: TEXT, callback: v => '₱' + v.toLocaleString() }}, grid: {{ color: GRID }} }}
  }}, plugins: {{ ...defaults.plugins, legend: {{ display: false }} }} }}
}});

// Daily Trend Line
new Chart(document.getElementById('trendChart'), {{
  type: 'line',
  data: {{
    labels: {daily_labels},
    datasets: [{{
      label: 'Daily Revenue (₱)',
      data: {daily_values},
      borderColor: '#4f8ef7',
      backgroundColor: 'rgba(79,142,247,0.12)',
      fill: true,
      tension: 0.35,
      pointRadius: 4,
      pointBackgroundColor: '#4f8ef7'
    }}]
  }},
  options: {{ ...defaults, scales: {{
    x: {{ ticks: {{ color: TEXT, maxRotation: 45 }}, grid: {{ color: GRID }} }},
    y: {{ ticks: {{ color: TEXT, callback: v => '₱' + v.toLocaleString() }}, grid: {{ color: GRID }} }}
  }} }}
}});
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"\n[REPORT] HTML report saved → '{output_path}'")


def export_cleaned(df: pd.DataFrame) -> None:
    """Save cleaned data to CSV and Excel."""
    df.to_csv(CLEAN_CSV, index=False)
    print(f"[EXPORT] CSV  saved → '{CLEAN_CSV}'")

    try:
        df.to_excel(CLEAN_EXCEL, index=False, engine="openpyxl")
        print(f"[EXPORT] XLSX saved → '{CLEAN_EXCEL}'")
    except ImportError:
        print("[EXPORT] openpyxl not installed — skipping Excel export (CSV still saved)")



def main():
    print("=" * 60)
    print(" Data Cleaning & Reporting Automation")
    print("=" * 60)

    
    df_raw = load_data(INPUT_FILE)

    
    raw_audit = audit_data(df_raw)

    
    df_clean, clean_log = clean_data(df_raw)

    
    agg = build_aggregates(df_clean)

    
    generate_report(raw_audit, clean_log, agg, REPORT_HTML)

    
    export_cleaned(df_clean)

    print("\n" + "=" * 60)
    print(" All done! Outputs:")
    print(f"     • {REPORT_HTML}  ← Open in browser")
    print(f"     • {CLEAN_CSV}")
    print(f"     • {CLEAN_EXCEL}")
    print("=" * 60)


if __name__ == "__main__":
    main()
