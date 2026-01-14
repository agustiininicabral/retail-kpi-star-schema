import argparse
from .pipeline import run

def main():
    p = argparse.ArgumentParser(prog="retail_kpi", description="Excel -> Star Schema (SQLite) -> SQL KPIs -> Excel+Charts")
    sub = p.add_subparsers(dest="cmd", required=True)

    runp = sub.add_parser("run", help="Run the pipeline")
    runp.add_argument("--excel", required=True, help="Path to Excel file")
    runp.add_argument("--sheet", default=0, help="Sheet name/index (default 0)")
    runp.add_argument("--db", default="data/warehouse.db", help="SQLite DB path")
    runp.add_argument("--sql", default="sql/kpis.sql", help="KPI SQL file path")
    runp.add_argument("--out", default="outputs", help="Outputs directory")

    args = p.parse_args()
    if args.cmd == "run":
        run(excel_path=args.excel, sheet=args.sheet, db_path=args.db, sql_path=args.sql, out_dir=args.out)
