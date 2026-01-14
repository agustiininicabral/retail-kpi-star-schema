-- KPI queries (SQLite)
-- Se asume modelo estrella:
-- dim_date(date_key, fecha, anio, mes, mes_nombre)
-- dim_product(product_key, producto)
-- dim_customer(customer_key, cliente)
-- fact_sales(date_key, product_key, customer_key, cantidad, precio_unitario, costo_unitario, venta, costo, margen)

-- 1) Ventas y margen total
SELECT
  'ventas_y_margen_total' AS metric,
  SUM(venta) AS ventas_total,
  SUM(margen) AS margen_total,
  ROUND( (SUM(margen) * 100.0) / NULLIF(SUM(venta), 0), 2) AS margen_pct
FROM fact_sales;

-- 2) Ticket promedio (venta promedio por venta)
SELECT
  'ticket_promedio' AS metric,
  ROUND(AVG(venta), 2) AS ticket_promedio
FROM fact_sales;

-- 3) Ventas por mes
SELECT
  d.anio,
  d.mes,
  d.mes_nombre,
  SUM(f.venta) AS ventas
FROM fact_sales f
JOIN dim_date d ON d.date_key = f.date_key
GROUP BY d.anio, d.mes, d.mes_nombre
ORDER BY d.anio, d.mes;

-- 4) Top productos por ventas
SELECT
  p.producto,
  SUM(f.venta) AS ventas
FROM fact_sales f
JOIN dim_product p ON p.product_key = f.product_key
GROUP BY p.producto
ORDER BY ventas DESC
LIMIT 10;

-- 5) Top clientes por ventas
SELECT
  c.cliente,
  SUM(f.venta) AS ventas
FROM fact_sales f
JOIN dim_customer c ON c.customer_key = f.customer_key
GROUP BY c.cliente
ORDER BY ventas DESC
LIMIT 10;
