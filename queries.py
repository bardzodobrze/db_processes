import pyodbc

def set_connection_and_cursor(server, database):
    # Create the connection string with Windows Authentication
    connection_string = (
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={server};'
        f'DATABASE={database};'
        f'Trusted_Connection=yes;'
    )

    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()

    return conn, cursor


def get_deliveries_from_db(cursor, date):
    cursor.execute(
        f"""
        SELECT
            D.Location_Code as Tienda,
            D.Invoice_Number as Documento_Referencia,
            D.Added as Fecha_Hora_Movimiento,
            D.Delivery_Amount as Total_soles,
            Y.Description as Nombre_Inventario,
            I.Inventory_Code as Codigo_Inventario, 
            I.VendorItemCode as Codigo_Vendedor,
            I.PortionQty as Cantidad_Producto, 
            I.PortionPrice as Precio_Porcion_Producto
        FROM Deliveries D
        INNER JOIN INVDeliveryAmounts I on  D.Delivery_ID = I.Delivery_ID
        INNER JOIN Inventory_Items Y on Y.Inventory_Code = I.Inventory_Code
        WHERE YEAR(d.delivery_date) = ? and MONTH(d.Delivery_Date) = ?
        ORDER BY Fecha_Hora_Movimiento
        """,
        (date.year, date.month)
    )
    return cursor.fetchall()


def get_orders_by_coupon_and_dates(cursor, coupon_name, date_1, date_2):
    cursor.execute(
        """
        SELECT 
            oc.[Location_Code],
            oc.[OrdCpnUpdateDate],
            oc.[Order_Number],
            o.[OrderFinalPrice],
            oc.[CouponCode],
            c2.[CouponDescText],
            SUM(oc.[OrdCpnQty]) AS Cantidad,
            mo.[Description],
            CASE 
                WHEN CHARINDEX('@', CouponPosDescText) > 0 THEN 
                    LTRIM(SUBSTRING(CouponPosDescText, CHARINDEX('@', CouponPosDescText) + 1, LEN(CouponPosDescText)))
                ELSE 
                    ''
            END AS Precio
        FROM 
            [POS].[dbo].[OrderCoupons] oc
        INNER JOIN 
            POS..[Coupons2] c2 ON c2.[CouponCode] = oc.[CouponCode]
        INNER JOIN 
            POS..Orders o ON oc.[Order_Number] = o.[Order_Number]
        INNER JOIN 
            POS..Order_Type_Codes mo ON o.Location_Code = mo.Location_Code
                                     AND o.Order_Type_Code = mo.Order_Type_Code
        WHERE 
            oc.[Order_Date] BETWEEN ? AND ?
            AND oc.[CouponCode] = ?
        GROUP BY 
            oc.[Location_Code],
            oc.[OrdCpnUpdateDate],
            o.[OrderFinalPrice],
            oc.[Order_Number],
            oc.[CouponCode],
            mo.[Description],
            c2.CouponPosDescText,
            c2.[CouponDescText]
        ORDER BY 
            oc.[Location_Code],
            oc.[OrdCpnUpdateDate],
            oc.[Order_Number];
        """,
        (date_1, date_2, coupon_name)
    )
    return cursor.fetchall()

def get_order_lines_by_store_and_date(cursor, date):
    cursor.execute(
        """
        SELECT
            [Added], [Location_Code], [Order_Number], [OrdLineFinalPrice]
        FROM [POS].[dbo].[Order_Lines]
        WHERE [Deleted] = 0
        AND [Added] >= ?
        """,
        date
    )
    return cursor.fetchall()



