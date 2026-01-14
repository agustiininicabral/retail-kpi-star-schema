from __future__ import annotations
import os
import sqlite3
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt

REQUIRED_COLS = ["fecha", "cliente", "producto", "cantidad", "precio_unitario", "costo_unitario"]

def _ensure_dirs(db_path: str, out_dir: str):
    os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(os.path.join(out_dir, "charts"), exist_ok=True)

def _read_excel(path: str, sheet) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    df.columns = [str(c).strip().lower() for c in df.columns]
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}. Found: {list(df.columns)}")
    return df

def _clean(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.dropna(how="all").copy()
    # parse fecha
    df2["fecha"] = pd.to_datetime(df2["fecha"], errors="coerce")
    # numeric
    for c in ["cantidad", "precio_unitario", "costo_unitario"]:
        df2[c] = pd.to_numeric(df2[c], errors="coerce")
    # trim text
    for c in ["cliente", "producto"]:
        df2[c] = df2[c].astype(str).str.strip()
    # drop rows with critical nulls
    df2 = df2.dropna(subset=["fecha", "cliente", "producto", "cantidad", "precio_unitario", "costo_unitario"])
    # filter invalid quantities
    df2 = df2[df2["cantidad"] > 0]
    df2 = df2[df2["precio_unitario"] >= 0]
    df2 = df2[df2["costo_unitario"] >= 0]
    return df2

def _build_star_schema(df: pd.DataFrame, con: sqlite3.Connection):
    # dim_date
    dim_date = (
        df[["fecha"]].drop_duplicates().assign(
            anio=lambda x: x["fecha"].dt.year,
            mes=lambda x: x["fecha"].dt.month,
            mes_nombre=lambda x: x["fecha"].dt.strftime("%b")
        ).sort_values("fecha").reset_index(drop=True)
    )
    dim_date["date_key"] = range(1, len(dim_date) + 1)

    # dim_product
    dim_product = df[["producto"]].drop_duplicates().sort_values("producto").reset_index(drop=True)
    dim_product["product_key"] = range(1, len(dim_product) + 1)

    # dim_customer
    dim_customer = df[["cliente"]].drop_duplicates().sort_values("cliente").reset_index(drop=True)
    dim_customer["customer_key"] = range(1, len(dim_customer) + 1)

    # fact
    f = df.merge(dim_date[["fecha","date_key"]], on="fecha") \
          .merge(dim_product, on="producto") \
          .merge(dim_customer, on="cliente")

    f["venta"] = f["cantidad"] * f["precio_unitario"]
    f["costo"] = f["cantidad"] * f["costo_unitario"]
    f["margen"] = f["venta"] - f["costo"]

    fact_sales = f[["date_key","product_key","customer_key","cantidad","precio_unitario","costo_unitario","venta","costo","margen"]]

    # write to sqlite
    dim_date.to_sql("dim_date", con, if_exists="replace", index=False)
    dim_product.to_sql("dim_product", con, if_exists="replace", index=False)
    dim_customer.to_sql("dim_customer", con, if_exists="replace", index=False)
    fact_sales.to_sql("fact_sales", con, if_exists="replace", index=False)

def _run_multi_select(con: sqlite3.Connection, sql_text: str):
    cur = con.cursor()
    statements = [s.strip() for s in sql_text.split(";") if s.strip()]
    frames = []
    for stmt in statements:
        cur.execute(stmt)
        cols = [d[0] for d in cur.description] if cur.description else []
        rows = cur.fetchall()
        frames.append(pd.DataFrame(rows, columns=cols))
    return frames

def _save_kpis(frames: list[pd.DataFrame], out_dir: str):
    # Map outputs by index following kpis.sql order
    kpi_total = frames[0]
    ticket = frames[1]
    ventas_mes = frames[2]
    top_prod = frames[3]
    top_cli = frames[4]

    # CSV summary (simple)
    summary = pd.concat([
        kpi_total.assign(section="totales"),
        ticket.assign(section="ticket_promedio")
    ], axis=0, ignore_index=True)

    summary_path = os.path.join(out_dir, "kpis_summary.csv")
    summary.to_csv(summary_path, index=False, encoding="utf-8")

    # Excel output
    xlsx_path = os.path.join(out_dir, "kpis.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        kpi_total.to_excel(writer, sheet_name="Totales", index=False)
        ticket.to_excel(writer, sheet_name="Ticket", index=False)
        ventas_mes.to_excel(writer, sheet_name="VentasPorMes", index=False)
        top_prod.to_excel(writer, sheet_name="TopProductos", index=False)
        top_cli.to_excel(writer, sheet_name="TopClientes", index=False)

    return {"summary_csv": summary_path, "kpis_xlsx": xlsx_path, "ventas_mes": ventas_mes, "top_prod": top_prod}

def _chart_ventas_por_mes(ventas_mes: pd.DataFrame, out_dir: str):
    path = os.path.join(out_dir, "charts", "ventas_por_mes.png")
    # ensure month order
    ventas_mes = ventas_mes.copy()
    ventas_mes["anio_mes"] = ventas_mes["anio"].astype(str) + "-" + ventas_mes["mes"].astype(str).str.zfill(2)
    plt.figure()
    plt.plot(ventas_mes["anio_mes"], ventas_mes["ventas"], marker="o")
    plt.xticks(rotation=45, ha="right")
    plt.title("Ventas por mes")
    plt.xlabel("Año-Mes")
    plt.ylabel("Ventas")
    plt.tight_layout()
    plt.savefig(path, dpi=160)
    plt.close()
    return path

def _chart_top_productos(top_prod: pd.DataFrame, out_dir: str):
    path = os.path.join(out_dir, "charts", "top_productos.png")
    plt.figure()
    plt.bar(top_prod["producto"], top_prod["ventas"])
    plt.xticks(rotation=45, ha="right")
    plt.title("Top 10 productos por ventas")
    plt.xlabel("Producto")
    plt.ylabel("Ventas")
    plt.tight_layout()
    plt.savefig(path, dpi=160)
    plt.close()
    return path

def _write_report_md(meta: dict, out_dir: str, artifacts: dict):
    md_path = os.path.join(out_dir, "REPORT.md")
    lines = []
    lines.append("# Retail KPI Star Schema Report")
    lines.append("")
    lines.append(f"- Generated at: **{meta['generated_at']}**")
    lines.append(f"- Source Excel: **{meta['excel_path']}**")
    lines.append(f"- Rows processed: **{meta['rows']}**")
    lines.append(f"- SQLite DB: **{meta['db_path']}**")
    lines.append("")
    lines.append("## Artifacts")
    lines.append(f"- KPIs Excel: `{os.path.basename(artifacts['kpis_xlsx'])}`")
    lines.append(f"- KPIs Summary CSV: `{os.path.basename(artifacts['summary_csv'])}`")
    lines.append("- Charts:")
    lines.append(f"  - `charts/{os.path.basename(artifacts['chart_ventas_mes'])}`")
    lines.append(f"  - `charts/{os.path.basename(artifacts['chart_top_prod'])}`")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return md_path

def run(excel_path: str, sheet, db_path: str, sql_path: str, out_dir: str):
    _ensure_dirs(db_path, out_dir)

    df = _read_excel(excel_path, sheet)
    df = _clean(df)

    with sqlite3.connect(db_path) as con:
        _build_star_schema(df, con)
        with open(sql_path, "r", encoding="utf-8") as f:
            sql_text = f.read()
        frames = _run_multi_select(con, sql_text)

    artifacts = _save_kpis(frames, out_dir)
    chart1 = _chart_ventas_por_mes(artifacts["ventas_mes"], out_dir)
    chart2 = _chart_top_productos(artifacts["top_prod"], out_dir)
    artifacts["chart_ventas_mes"] = chart1
    artifacts["chart_top_prod"] = chart2

    meta = {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "excel_path": excel_path,
        "rows": int(len(df)),
        "db_path": db_path
    }
    _write_report_md(meta, out_dir, artifacts)

    print("OK - Retail KPI pipeline finished")
    print(f"- DB: {db_path}")
    print(f"- Outputs: {out_dir}")
