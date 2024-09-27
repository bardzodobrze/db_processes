import datetime
import os
import pandas as pd
import pyodbc
import sys
import time
import xlwings as xw
import win32com.client as win32

import queries


def send_email_with_observed_notes(email_addresses, date_1, date_2, attachment_path):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(email_addresses)
    mail.Subject = f'Ventas DPI del {date_1.strftime("%d/%m/%Y")} al {date_2.strftime("%d/%m/%Y")}'
    mail.Attachments.Add(attachment_path)
    mail.HTMLBody = "Se adjunta la información mencionada en el asunto. Este es un correo autogenerado."

    mail.Send()
    print("Correo enviado.")

    # if os.path.isfile(attachment_path):
        # os.remove(attachment_path)


def main(date_1, date_2, store_ids):
    start_time = time.time()

    total_sales_by_store = pd.DataFrame(columns=[
        'RecordType', 'DatabaseVersion', 'Location_Code', 'BeginDate', 'EndDate', 'Order_Count',
        'Royalty_Sales', 'Pizza_Quantity', 'Food_Cost_Percent', 'Labor_Cost_Percent'
    ])

    skipped_stores = []

    for store_id in store_ids:
        try:
            print(store_id)
            # Define your connection parameters
            server = "PULSEBOS" + store_id
            database = "pos"
            conn, cursor = queries.set_connection_and_cursor(server, database)

            # 1. VENTAS DBO
            cursor.execute("{CALL dbo.spExtractWeeklyKeysV34 (?, ?, ?)}", (store_id, date_1, date_2))
            rows = cursor.fetchall()
            for row in rows:
                total_sales_by_store.loc[len(total_sales_by_store)] = [
                    row[0], row[1], row[2], row[3], row[4], row[11], row[22], row[73], row[85], row[97]
                ]
            cursor.close()
            conn.close()
        except pyodbc.OperationalError:
            skipped_stores.append(store_id)
            print(f"Store with id {store_id} skipped.")
            sys.exit()

    for column_name in ["Royalty_Sales", "Food_Cost_Percent", "Labor_Cost_Percent"]:
        total_sales_by_store[column_name] = total_sales_by_store[column_name].astype(float)

    print(total_sales_by_store.to_string())

    wb = xw.Book("ventas_dpi_template.xlsx")
    ws = wb.sheets["data"]
    ws["A1:J30"].value = ""
    ws["A1"].options(pd.DataFrame, index=False, expand='table').value = total_sales_by_store

    filename = f"ventas_dpi_{date_2.strftime("%Y-%m-%d")}.xlsx"
    wb.save(filename)
    wb.close()

    email_addresses = [
        "reportesbi@dominos.com.pe",
        "impuestos@dominos.com.pe",
    ]

    filepath = os.path.join(os.getcwd(), filename)
    send_email_with_observed_notes(email_addresses, date_1, date_2, filepath)

    print(f"Skipped stores: {skipped_stores}")
    print(f"Runtime: {time.time() - start_time} seconds.")



# DISEÑADO PARA CORRER TODOS LOS LUNES EN LA NOCHE (DEBE ESTAR DISPONIBLE LOS MARTES EN LA MAÑANA)
# This must be the Sunday of the target week
date_2 = datetime.datetime(2024, 9, 22)

date_1 = date_2 - datetime.timedelta(days=6)
main(date_1, date_2, queries.store_ids)
