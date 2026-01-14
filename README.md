# retail-kpi-star-schema

Proyecto de **Data Analytics** para portfolio: transforma un Excel de ventas a un **modelo estrella (star schema)** en SQLite y genera:
- KPIs (ventas, margen, ticket promedio, top productos, top clientes)
- Consultas SQL reproducibles
- Export a Excel con tablas de resultados
- Gráficos PNG (matplotlib) para evidencias

Tecnologías: **Excel + Python (pandas) + SQLite + SQL + Reporting**.

## Estructura
- `src/` código
- `sql/` queries KPI
- `examples/` Excel demo
- `outputs/` resultados generados
- `data/warehouse.db` base SQLite

## Setup
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run (demo)
```bash
python -m retail_kpi run --excel examples/ventas_retail_demo.xlsx --sheet Ventas
```

Outputs:
- `data/warehouse.db`
- `outputs/kpis.xlsx`
- `outputs/kpis_summary.csv`
- `outputs/charts/ventas_por_mes.png`
- `outputs/charts/top_productos.png`

## Cómo adaptarlo a tu dataset
- Asegurate de tener columnas similares a:
  - `fecha`, `cliente`, `producto`, `cantidad`, `precio_unitario`, `costo_unitario`
- Ejecutá el CLI apuntando a tu archivo.

## Ideas de mejora (para subir el nivel)
- Migrar de SQLite a PostgreSQL
- Agregar incremental loads
- Tests (pytest) + quality gates
- Publicar report HTML
