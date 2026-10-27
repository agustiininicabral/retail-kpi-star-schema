# Retail KPI Star Schema

Proyecto de Data Analytics que transforma un Excel de ventas en un modelo estrella SQLite, ejecuta consultas SQL y genera KPIs, reportes y gráficos.

## Requisitos

- Python 3.12
- Dependencias de ejecución en `requirements.txt`
- Pytest y Ruff en `requirements-dev.txt`

## Instalación (Windows PowerShell)

Desde la raíz del repositorio:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

## Ejecutar la demo

```powershell
.\.venv\Scripts\python.exe -m retail_kpi run --excel examples/ventas_retail_demo.xlsx --sheet Ventas
```

El pipeline genera:

- `data/warehouse.db`: tablas `dim_date`, `dim_product`, `dim_customer` y `fact_sales`.
- `outputs/kpis.xlsx`: totales, ticket, ventas mensuales, top productos y top clientes.
- `outputs/kpis_summary.csv`: resumen de KPIs.
- `outputs/REPORT.md`: metadatos, conteo de filas rechazadas y lista de artefactos.
- `outputs/rejected_rows.csv`: filas descartadas con número de fila y motivo de rechazo.
- `outputs/charts/ventas_por_mes.png` y `outputs/charts/top_productos.png`.

La base y los archivos generados se excluyen de Git.

## Limpieza y validación

Los nombres de columnas se normalizan a minúsculas y se eliminan espacios exteriores. Se descartan filas completamente vacías, fechas inválidas, campos críticos nulos o vacíos, valores numéricos no interpretables, cantidades no positivas y precios/costos negativos. Si no queda ninguna fila válida, el proceso termina con un error explícito y no genera KPIs vacíos.

Las filas descartadas no se escriben en el modelo estrella. Se conserva una auditoría local en `outputs/rejected_rows.csv`, con el número de fila de Excel y códigos de motivo (fecha inválida, campos faltantes, valores numéricos no interpretables o rangos inválidos). Revisá el archivo antes de utilizar los resultados para decisiones comerciales; puede contener datos originales y no debe publicarse si el Excel contiene información sensible.

## Pruebas y controles de calidad

```powershell
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe -m ruff format --check retail_kpi tests
.\.venv\Scripts\python.exe -m ruff check retail_kpi tests
.\.venv\Scripts\python.exe -m pytest -q
```

Las pruebas crean un Excel temporal con datos válidos e inválidos y comprueban el modelo estrella, las salidas, los gráficos, la limpieza y el manejo de columnas faltantes. GitHub Actions ejecuta estas pruebas y la demo de extremo a extremo; los resultados se publican como artefactos temporales.

## Estructura

- `retail_kpi/pipeline.py`: limpieza, modelo estrella, SQL, Excel y gráficos.
- `retail_kpi/cli.py`: interfaz de línea de comandos.
- `sql/kpis.sql`: consultas de KPI.
- `examples/ventas_retail_demo.xlsx`: dataset de demostración.
- `tests/test_pipeline.py`: pruebas automatizadas del pipeline.

## Limitaciones

El proyecto calcula métricas a nivel de fila de ventas porque el Excel de entrada no incluye un identificador de pedido. Por eso, el indicador denominado ticket promedio representa el promedio de ventas por registro del archivo, no necesariamente el valor promedio de una orden de compra completa. Los datos de ejemplo no representan operaciones reales.
