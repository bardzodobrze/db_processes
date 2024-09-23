import datetime
import numpy as np
import pandas
import pandas as pd
import pyodbc
import time

import queries


store_ids = [
    "18600",
    "18601",
    "18602",
    "18603",
    "18604",
    "18605",
    "18606",
    "18607",
    "18608",
    "18609",
    "18610",
    "18611",
    "18612",
    "18614",
    "18615",
    "18616",
    "18617",
    "18618",
    "18619",
    "18620",
    "18621",
    "18622",
    "18623",
    "18624",
]

# store_ids = ["18600"]

start_time = time.time()

# We define an interval of a week (7 days) ending today
today = datetime.datetime(2024, 9, 15)
# today = datetime.datetime.today().replace(hour=0, minute=0, second=0)
start_date = today - datetime.timedelta(days=6)

total_sales_by_store = pd.DataFrame(columns=["loc_code", "start_date", "end_date", "master_sales"])

deliveries_by_store = pd.DataFrame(columns=[
    "Tienda", "Documento_Referencia", "Fecha_Hora_Movimiento", "Total_soles", "Nombre_Inventario", "Codigo_Inventario",
    "Codigo_Vendedor", "Cantidad_Producto", "Precio_Porcion_Producto"
])

for store_id in store_ids:
    print(store_id)
    # Define your connection parameters
    server = "PULSEBOS" + store_id
    database = "pos"
    conn, cursor = queries.set_connection_and_cursor(server, database)

    # 1. VENTAS DBO
    cursor.execute("{CALL dbo.spExtractWeeklyKeysV34 (?, ?, ?)}", (store_id, start_date, today))
    rows = cursor.fetchall()
    for row in rows:
        total_sales_by_store.loc[len(total_sales_by_store)] = [row[2], row[3], row[4], float(row[6])]

    # 2. REPARTOS: ENTRADAS Y SALIDAS
    deliveries = queries.get_deliveries_from_db(cursor, today)
    for delivery in deliveries:
        deliveries_by_store.loc[len(deliveries_by_store)] = list(delivery)

    cursor.close()
    conn.close()

# Casting columns as float

total_sales_by_store['master_sales'] = total_sales_by_store['master_sales'].astype(float)
deliveries_by_store['Total_soles'] = deliveries_by_store['Total_soles'].astype(float)
deliveries_by_store['Cantidad_Producto'] = deliveries_by_store['Cantidad_Producto'].astype(float)
deliveries_by_store['Precio_Porcion_Producto'] = deliveries_by_store['Precio_Porcion_Producto'].astype(float)

print(total_sales_by_store.to_string())
print(deliveries_by_store.to_string())
total_sales_by_store.to_excel("total_sales_by_store.xlsx", index=False)
deliveries_by_store.to_excel("deliveries_by_store.xlsx", index=False)

print(f"Runtime: {time.time() - start_time} seconds.")