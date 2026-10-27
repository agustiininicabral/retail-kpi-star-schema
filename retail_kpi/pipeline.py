from __future__ import annotations

import os
import sqlite3
from datetime import datetime, timezone
from typing import Any

import matplotlib.pyplot as plt
import pandas as pd

REQUIRED_COLS = (
    "fecha",
    "cliente",
    "producto",
    "cantidad",
    "precio_unitario",
    "costo_unitario",
)


def _ensure_dirs(db_path: str, out_dir: str) -> None:
    os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)
    os.makedirs(out_dir, exist_ok=True)
    os.makedirs(os.path.join(out_dir, "charts"), exist_ok=True)


def _read_excel(path: str, sheet: str | int) -> pd.DataFrame:
    dataframe = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    dataframe.columns = [str(column).strip().lower() for column in dataframe.columns]
    missing = [column for column in REQUIRED_COLS if column not in dataframe.columns]
    if missing:
        raise ValueError(
            f"Missing required columns: {missing}. Found: {list(dataframe.columns)}"
        )
    return dataframe


def _clean(dataframe: pd.DataFrame) -> pd.DataFrame:
    cleaned = dataframe.dropna(how="all").copy()
    cleaned["fecha"] = pd.to_datetime(cleaned["fecha"], errors="coerce")

    for column in ("cantidad", "precio_unitario", "costo_unitario"):
        cleaned[column] = pd.to_numeric(cleaned[column], errors="coerce")

    for column in ("cliente", "producto"):
        cleaned[column] = cleaned[column].astype("string").str.strip()
        cleaned[column] = cleaned[column].replace("", pd.NA)

    cleaned = cleaned.dropna(subset=list(REQUIRED_COLS))
    cleaned = cleaned[
        (cleaned["cantidad"] > 0)
        & (cleaned["precio_unitario"] >= 0)
        & (cleaned["costo_unitario"] >= 0)
    ]
    return cleaned


def _build_star_schema(dataframe: pd.DataFrame, connection: sqlite3.Connection) -> None:
    dim_date = (
        dataframe[["fecha"]]
        .drop_duplicates()
        .assign(
            anio=lambda frame: frame["fecha"].dt.year,
            mes=lambda frame: frame["fecha"].dt.month,
            mes_nombre=lambda frame: frame["fecha"].dt.strftime("%b"),
        )
        .sort_values("fecha")
        .reset_index(drop=True)
    )
    dim_date["date_key"] = range(1, len(dim_date) + 1)

    dim_product = (
        dataframe[["producto"]]
        .drop_duplicates()
        .sort_values("producto")
        .reset_index(drop=True)
    )
    dim_product["product_key"] = range(1, len(dim_product) + 1)

    dim_customer = (
        dataframe[["cliente"]]
        .drop_duplicates()
        .sort_values("cliente")
        .reset_index(drop=True)
    )
    dim_customer["customer_key"] = range(1, len(dim_customer) + 1)

    fact = (
        dataframe.merge(dim_date[["fecha", "date_key"]], on="fecha")
        .merge(dim_product, on="producto")
        .merge(dim_customer, on="cliente")
    )
    fact["venta"] = fact["cantidad"] * fact["precio_unitario"]
    fact["costo"] = fact["cantidad"] * fact["costo_unitario"]
    fact["margen"] = fact["venta"] - fact["costo"]

    fact_sales = fact[
        [
            "date_key",
            "product_key",
            "customer_key",
            "cantidad",
            "precio_unitario",
            "costo_unitario",
            "venta",
            "costo",
            "margen",
        ]
    ]

    dim_date.to_sql("dim_date", connection, if_exists="replace", index=False)
    dim_product.to_sql("dim_product", connection, if_exists="replace", index=False)
    dim_customer.to_sql("dim_customer", connection, if_exists="replace", index=False)
    fact_sales.to_sql("fact_sales", connection, if_exists="replace", index=False)


def _run_multi_select(
    connection: sqlite3.Connection,
    sql_text: str,
) -> list[pd.DataFrame]:
    frames = []
    for statement in (part.strip() for part in sql_text.split(";")):
        if not statement:
            continue
        cursor = connection.execute(statement)
        columns = [description[0] for description in cursor.description or []]
        frames.append(pd.DataFrame(cursor.fetchall(), columns=columns))
    return frames


def _save_kpis(frames: list[pd.DataFrame], out_dir: str) -> dict[str, Any]:
    if len(frames) != 5:
        raise ValueError(f"Expected 5 KPI query results, got {len(frames)}.")

    kpi_total, ticket, ventas_mes, top_prod, top_cli = frames
    summary = pd.concat(
        [
            kpi_total.assign(section="totales"),
            ticket.assign(section="ticket_promedio"),
        ],
        ignore_index=True,
    )

    summary_path = os.path.join(out_dir, "kpis_summary.csv")
    summary.to_csv(summary_path, index=False, encoding="utf-8")

    xlsx_path = os.path.join(out_dir, "kpis.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        kpi_total.to_excel(writer, sheet_name="Totales", index=False)
        ticket.to_excel(writer, sheet_name="Ticket", index=False)
        ventas_mes.to_excel(writer, sheet_name="VentasPorMes", index=False)
        top_prod.to_excel(writer, sheet_name="TopProductos", index=False)
        top_cli.to_excel(writer, sheet_name="TopClientes", index=False)

    return {
        "summary_csv": summary_path,
        "kpis_xlsx": xlsx_path,
        "ventas_mes": ventas_mes,
        "top_prod": top_prod,
    }


def _chart_ventas_por_mes(ventas_mes: pd.DataFrame, out_dir: str) -> str:
    path = os.path.join(out_dir, "charts", "ventas_por_mes.png")
    monthly = ventas_mes.copy()
    monthly["anio_mes"] = (
        monthly["anio"].astype(str) + "-" + monthly["mes"].astype(str).str.zfill(2)
    )

    figure, axis = plt.subplots()
    axis.plot(monthly["anio_mes"], monthly["ventas"], marker="o")
    axis.tick_params(axis="x", labelrotation=45)
    axis.set_title("Ventas por mes")
    axis.set_xlabel("Año-Mes")
    axis.set_ylabel("Ventas")
    figure.tight_layout()
    figure.savefig(path, dpi=160)
    plt.close(figure)
    return path


def _chart_top_productos(top_prod: pd.DataFrame, out_dir: str) -> str:
    path = os.path.join(out_dir, "charts", "top_productos.png")
    figure, axis = plt.subplots()
    axis.bar(top_prod["producto"], top_prod["ventas"])
    axis.tick_params(axis="x", labelrotation=45)
    axis.set_title("Top 10 productos por ventas")
    axis.set_xlabel("Producto")
    axis.set_ylabel("Ventas")
    figure.tight_layout()
    figure.savefig(path, dpi=160)
    plt.close(figure)
    return path


def _write_report_md(
    meta: dict[str, Any],
    out_dir: str,
    artifacts: dict[str, Any],
) -> str:
    report_path = os.path.join(out_dir, "REPORT.md")
    lines = [
        "# Retail KPI Star Schema Report",
        "",
        f"- Generated at: **{meta['generated_at']}**",
        f"- Source Excel: **{meta['excel_path']}**",
        f"- Rows processed: **{meta['rows']}**",
        f"- SQLite DB: **{meta['db_path']}**",
        "",
        "## Artifacts",
        f"- KPIs Excel: `{os.path.basename(artifacts['kpis_xlsx'])}`",
        f"- KPIs Summary CSV: `{os.path.basename(artifacts['summary_csv'])}`",
        "- Charts:",
        f"  - `charts/{os.path.basename(artifacts['chart_ventas_mes'])}`",
        f"  - `charts/{os.path.basename(artifacts['chart_top_prod'])}`",
    ]
    with open(report_path, "w", encoding="utf-8") as report:
        report.write("\n".join(lines))
    return report_path


def run(
    excel_path: str,
    sheet: str | int,
    db_path: str,
    sql_path: str,
    out_dir: str,
) -> None:
    _ensure_dirs(db_path, out_dir)
    dataframe = _clean(_read_excel(excel_path, sheet))
    if dataframe.empty:
        raise ValueError(
            "No valid sales rows remain after cleaning the input workbook."
        )

    with sqlite3.connect(db_path) as connection:
        _build_star_schema(dataframe, connection)
        with open(sql_path, encoding="utf-8") as sql_file:
            frames = _run_multi_select(connection, sql_file.read())

    artifacts = _save_kpis(frames, out_dir)
    artifacts["chart_ventas_mes"] = _chart_ventas_por_mes(
        artifacts["ventas_mes"], out_dir
    )
    artifacts["chart_top_prod"] = _chart_top_productos(
        artifacts["top_prod"], out_dir
    )

    metadata = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "excel_path": excel_path,
        "rows": int(len(dataframe)),
        "db_path": db_path,
    }
    _write_report_md(metadata, out_dir, artifacts)

    print("OK - Retail KPI pipeline finished")
    print(f"- DB: {db_path}")
    print(f"- Outputs: {out_dir}")
