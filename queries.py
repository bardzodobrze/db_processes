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

def get_orders_by_store_and_date(cursor, date):
    cursor.execute(
        """
        SELECT [Added], [Location_Code], [Order_Number], [OrderFinalPrice]
        FROM [POS].[dbo].[Orders]
        WHERE [Order_Date] = ?
        """,
        date
    )
    return cursor.fetchall()


def get_header_rows_between_dates(cursor, date_1, date_2):
    cursor.execute(
        """
        SELECT [Tienda], [Numero_Orden], [Orden_Fecha], [Orden_Hora], [SubTotal], [IGV], [RecCon], [OrderFinalPrice],
        [Customer], [serie_ce], [preimpreso_ce], [Doc_Anulado], [Estado_Orden], [Ruc], [RazonSocial], [MontoDelivery],
        [Met_Servicio], [Tipo_Pago], [name], [phone_number], [Vendedor], [Delivery]
        FROM [MIGRAPOS].[dbo].[Documentos_Cabecera_Hist]
        WHERE [Orden_Fecha] BETWEEN ? AND ?
        """,
        (date_1, date_2)
    )
    return cursor.fetchall()

def get_detail_rows_between_dates(cursor, date_1, date_2):
    cursor.execute(
        """
        SELECT [Numero_Orden], [Tienda], [FechaOrden], [HoraOrden], [SubTotal], [IGV], [RecargoCon],
        [PrecioFinal], [Cliente], [SerieComprobante], [PreimpresoComprobante], [DocumentoAnulado],
        [EstadoOrden], [RUC], [RazonSocial], [MetodoServicio], [TipoPago], [Item], [Descuento],
        [PrecioActual], [PrecioAntesDescuento], [PrecioDespuesDescuento], [Impuestos], [IGVDetalle],
        [RecargoCDetalle], [PrecioSinImpuestos], [PrecioTotal], [Cantidad], [ToppingDescriptions],
        [CodigoProducto], [Cupon], [DescripcionLinea], [ToppingDescriptionsPrepLinea], 
        [ToppingDescriptionsMakeLinea]
        FROM [MIGRAPOS].[dbo].[Detalle_Documentos_Hist]
        WHERE [FechaOrden] BETWEEN ? AND ?
        """,
        (date_1, date_2)
    )
    return cursor.fetchall()

def get_intraday_metrics(cursor, datetime_1, datetime_2):
    cursor.execute(
        """
        SELECT
            [Location_Code] AS [Tienda],
            COUNT(*) AS [Ordenes],
            SUM(OrderFinalPrice) AS [Venta],
            AVG(CAST(CASE WHEN [Delivery_Time] IS NOT NULL THEN DATEDIFF(MINUTE, [Order_Saved], [Delivery_Time]) END AS FLOAT)) AS [ADT]
        FROM
            [POS].[dbo].[Orders]
        WHERE 
            [Order_Saved] BETWEEN ? AND ?
        GROUP BY 
            [Location_Code];
        """,
        (datetime_1, datetime_2)
    )

    return cursor.fetchall()

def get_store_info(cursor):
    cursor.execute(
        """
        SELECT [Zona], [Responsable], [Id], [Descripcion]
        FROM [MIGRAPOS].[dbo].[Tienda_Detalle]
        """
    )
    return cursor.fetchall()



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
    "18625",
]


