from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd
import pytest

from retail_kpi.pipeline import _clean_with_rejections, run


ROOT = Path(__file__).resolve().parents[1]


def _sample_sales() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Fecha": "2025-01-05",
                "Cliente": "Ana",
                "Producto": "Crema",
                "Cantidad": 2,
                "Precio_Unitario": 100,
                "Costo_Unitario": 60,
            },
            {
                "Fecha": "2025-02-10",
                "Cliente": "Beto",
                "Producto": "Jabón",
                "Cantidad": 1,
                "Precio_Unitario": 50,
                "Costo_Unitario": 20,
            },
            {
                "Fecha": "2025-02-10",
                "Cliente": None,
                "Producto": "Jabón",
                "Cantidad": 1,
                "Precio_Unitario": 50,
                "Costo_Unitario": 20,
            },
            {
                "Fecha": "invalid-date",
                "Cliente": "Carla",
                "Producto": "Perfume",
                "Cantidad": -1,
                "Precio_Unitario": 300,
                "Costo_Unitario": 100,
            },
        ]
    )


def test_pipeline_builds_star_schema_kpis_and_charts(tmp_path: Path) -> None:
    source = tmp_path / "sales.xlsx"
    _sample_sales().to_excel(source, sheet_name="Ventas", index=False)
    database = tmp_path / "data" / "warehouse.db"
    output = tmp_path / "outputs"

    run(
        excel_path=str(source),
        sheet="Ventas",
        db_path=str(database),
        sql_path=str(ROOT / "sql" / "kpis.sql"),
        out_dir=str(output),
    )

    assert database.is_file()
    assert (output / "kpis.xlsx").is_file()
    assert (output / "kpis_summary.csv").is_file()
    assert (output / "REPORT.md").is_file()
    assert (output / "rejected_rows.csv").is_file()
    assert (output / "charts" / "ventas_por_mes.png").is_file()
    assert (output / "charts" / "top_productos.png").is_file()

    with sqlite3.connect(database) as connection:
        assert connection.execute("SELECT COUNT(*) FROM fact_sales").fetchone()[0] == 2
        assert (
            connection.execute("SELECT COUNT(*) FROM dim_customer").fetchone()[0] == 2
        )
        assert connection.execute("SELECT COUNT(*) FROM dim_product").fetchone()[0] == 2
        total_sales = connection.execute(
            "SELECT SUM(venta) FROM fact_sales"
        ).fetchone()[0]
        assert total_sales == 250

    rejected = pd.read_csv(output / "rejected_rows.csv")
    assert rejected["source_excel_row"].tolist() == [4, 5]
    assert "cliente_missing_or_blank" in rejected.iloc[0]["rejection_reasons"]
    assert "fecha_missing_or_invalid" in rejected.iloc[1]["rejection_reasons"]
    assert "cantidad_not_positive" in rejected.iloc[1]["rejection_reasons"]
    report = (output / "REPORT.md").read_text(encoding="utf-8")
    assert "Rows rejected: **2**" in report
    assert "Rejected rows audit" in report

    with pd.ExcelFile(output / "kpis.xlsx") as workbook:
        assert workbook.sheet_names == [
            "Totales",
            "Ticket",
            "VentasPorMes",
            "TopProductos",
            "TopClientes",
        ]


def test_pipeline_rejects_workbook_without_required_columns(tmp_path: Path) -> None:
    source = tmp_path / "invalid.xlsx"
    pd.DataFrame({"fecha": ["2025-01-01"], "cliente": ["Ana"]}).to_excel(
        source,
        index=False,
    )

    with pytest.raises(ValueError, match="Missing required columns"):
        run(
            excel_path=str(source),
            sheet=0,
            db_path=str(tmp_path / "data" / "warehouse.db"),
            sql_path=str(ROOT / "sql" / "kpis.sql"),
            out_dir=str(tmp_path / "outputs"),
        )


def test_pipeline_rejects_when_no_valid_sales_rows_remain(tmp_path: Path) -> None:
    source = tmp_path / "empty.xlsx"
    _sample_sales().iloc[2:].to_excel(source, index=False)

    with pytest.raises(ValueError, match="No valid sales rows remain"):
        run(
            excel_path=str(source),
            sheet=0,
            db_path=str(tmp_path / "data" / "warehouse.db"),
            sql_path=str(ROOT / "sql" / "kpis.sql"),
            out_dir=str(tmp_path / "outputs"),
        )
    rejected = pd.read_csv(tmp_path / "outputs" / "rejected_rows.csv")
    assert len(rejected) == 2
    assert rejected["rejection_reasons"].str.len().gt(0).all()


def test_clean_with_rejections_returns_valid_rows_and_reason_codes() -> None:
    cleaned, rejected = _clean_with_rejections(_sample_sales())

    assert len(cleaned) == 2
    assert len(rejected) == 2
    assert rejected["source_excel_row"].tolist() == [4, 5]
    assert rejected["rejection_reasons"].str.contains("cliente_missing_or_blank").any()
    assert rejected["rejection_reasons"].str.contains("fecha_missing_or_invalid").any()


def test_pipeline_rejects_duplicate_normalized_headers(tmp_path: Path) -> None:
    source = tmp_path / "duplicate.xlsx"
    pd.DataFrame([["Crema", "Jabón"]], columns=["Producto", " producto "]).to_excel(
        source, index=False
    )
    with pytest.raises(ValueError, match="Duplicate column names after normalization"):
        run(
            excel_path=str(source),
            sheet=0,
            db_path=str(tmp_path / "data" / "warehouse.db"),
            sql_path=str(ROOT / "sql" / "kpis.sql"),
            out_dir=str(tmp_path / "outputs"),
        )
