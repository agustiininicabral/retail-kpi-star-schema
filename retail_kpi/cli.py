import argparse

from .pipeline import run


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="retail_kpi",
        description="Excel -> Star Schema (SQLite) -> SQL KPIs -> Excel + charts",
    )
    subparsers = parser.add_subparsers(dest="cmd", required=True)

    run_parser = subparsers.add_parser("run", help="Run the pipeline")
    run_parser.add_argument("--excel", required=True, help="Path to Excel file")
    run_parser.add_argument("--sheet", default=0, help="Sheet name/index (default: 0)")
    run_parser.add_argument("--db", default="data/warehouse.db", help="SQLite DB path")
    run_parser.add_argument("--sql", default="sql/kpis.sql", help="KPI SQL file path")
    run_parser.add_argument("--out", default="outputs", help="Output directory")

    args = parser.parse_args()
    if args.cmd == "run":
        run(
            excel_path=args.excel,
            sheet=args.sheet,
            db_path=args.db,
            sql_path=args.sql,
            out_dir=args.out,
        )


if __name__ == "__main__":
    main()
