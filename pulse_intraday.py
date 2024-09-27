import datetime
import os
import pandas as pd
import pyodbc
import sys
import time
import xlwings as xw
import win32com.client as win32

import queries


def send_email_with_observed_notes(email_addresses, date_2, attachment_path, skipped_stores):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = ";".join(email_addresses)
    mail.Subject = f'Data intraday Pulse, {date_2.strftime("%d/%m/%Y %H:%M:%S")}'
    mail.Attachments.Add(attachment_path)
    mail.HTMLBody = "<p>Se adjunta la información mencionada en el asunto. Este es un correo autogenerado.</p>"
    if skipped_stores:
        mail.HTMLBody += f"<p>Tiendas sin datos: {skipped_stores}</p>"

    mail.Send()
    print("Correo enviado.")

    # if os.path.isfile(attachment_path):
        # os.remove(attachment_path)


def main(store_ids, datetime_1, datetime_2):
    start_time = time.time()

    summary_df = pd.DataFrame(columns=[
        'Tienda', 'Ordenes', 'Venta', 'ADT'
    ])

    skipped_stores = []

    for store_id in store_ids:
        try:
            print(store_id)
            # Define your connection parameters
            server = "PULSEBOS" + store_id
            database = "pos"
            conn, cursor = queries.set_connection_and_cursor(server, database)

            summary_rows = queries.get_intraday_metrics(cursor, datetime_1, datetime_2)
            for row in summary_rows:
                summary_df.loc[len(summary_df)] = list(row)
            cursor.close()
            conn.close()
        except pyodbc.OperationalError:
            skipped_stores.append(store_id)
            print(f"Store with id {store_id} skipped.")

    print(summary_df.to_string())

    # NOW WE ADD INFO ABOUT THE STORES
    server = "PULSEBOS18624"
    database = "MIGRAPOS"
    conn, cursor = queries.set_connection_and_cursor(server, database)
    stores_rows = queries.get_store_info(cursor)
    stores_df = pd.DataFrame(columns=["zona", "supervisor", "id_tienda", "nombre_tienda"])
    for row in stores_rows:
        stores_df.loc[len(stores_df)] = list(row)

    master_df = summary_df.merge(stores_df, left_on="Tienda", right_on="id_tienda")
    master_df = master_df[["zona", "supervisor", "id_tienda", "nombre_tienda", "Ordenes", "Venta", "ADT"]]
    master_df.sort_values(["zona", "Venta"], ascending=[True, False], inplace=True)

    wb = xw.Book("pulse_intraday_BS.xlsx")
    ws = wb.sheets["Hoja1"]
    # ws["A1:J30"].value = ""
    ws["A1"].options(pd.DataFrame, index=False, expand='table').value = master_df

    filename = f"pulse_intraday_{datetime_2.strftime("%Y-%m-%d_%H%M")}.xlsx"
    wb.save(filename)
    wb.close()

    email_addresses = [
        "reportesbi@dominos.com.pe",
        #"impuestos@dominos.com.pe",
    ]

    filepath = os.path.join(os.getcwd(), filename)
    # send_email_with_observed_notes(email_addresses, datetime_2, filepath, skipped_stores)

    print(f"Skipped stores: {skipped_stores}")
    print(f"Runtime: {time.time() - start_time} seconds.")



# DISEÑADO PARA CORRER TODOS LOS LUNES EN LA NOCHE (DEBE ESTAR DISPONIBLE LOS MARTES EN LA MAÑANA)
# This must be the Sunday of the target week

datetime_2 = datetime.datetime(2024, 9, 25, 22, 0, 0)
datetime_1 = datetime_2.replace(hour=12)
main(queries.store_ids, datetime_1, datetime_2)

