/*
===================================================================================================
 CONSULTA SAP BUSINESS ONE HANA: EXTRACCIÓN DE PARÁMETROS DE PLANIFICACIÓN DE ARTÍCULOS
 Empresa: Lazarus & Lazarus
 Motor de Base de Datos: SAP HANA Database (HDB)
 Tabla Principal: "OITM" (Datos Maestros de Artículo)
 Tabla Secundaria: "OITB" (Grupos de Artículos)
===================================================================================================
 PROPÓSITO:
   Extraer los artículos activos con sus parámetros actuales de planificación para generar
   el archivo Excel que se alimentará en el bot de actualización (ejecutar_bot_articulos.bat).
 
 CAMPOS CLAVE:
   - "Codigo_Articulo": "ItemCode"
   - "Descripcion":     "ItemName"
   - "Minimo_Compra":   "MinOrdrQty" (Cantidad de pedido mínimo)
   - "Tiempo_Lead":     "LeadTime"   (Tiempo lead en días)
   - "Dias_Tolerancia": "TlrnceDays" (Días de tolerancia / retraso)
===================================================================================================
 INSTRUCCIONES:
 1. Abre el Generador de Consultas en SAP B1.
 2. Pega este código SQL y presiona "Ejecutar".
 3. Exporta a Excel, modifica los valores de Lead Time, Mínimo de Compra o Tolerancia,
    y cárgalo en el bot para que los actualice automáticamente en SAP B1.
===================================================================================================
*/

SELECT 
    T0."ItemCode"                                                   AS "Codigo_Articulo",
    T0."ItemName"                                                   AS "Descripcion",
    T1."ItmsGrpNam"                                                 AS "Grupo_Articulos",
    CASE 
        WHEN T0."PrcrmntMtd" = 'B' THEN 'Comprar'
        WHEN T0."PrcrmntMtd" = 'M' THEN 'Fabricar'
        ELSE 'Otro'
    END                                                             AS "Metodo_Aprovisionamiento",
    IFNULL(T0."MinOrdrQty", 0)                                      AS "Minimo_Compra",
    IFNULL(T0."LeadTime", 0)                                        AS "Tiempo_Lead",
    IFNULL(T0."TlrnceDays", 0)                                      AS "Dias_Tolerancia",
    T0."InvntItem"                                                  AS "Es_Inventariable",
    T0."SellItem"                                                   AS "Es_Venta",
    T0."PrchseItem"                                                 AS "Es_Compra"
FROM "OITM" T0
LEFT JOIN "OITB" T1 ON T0."ItmsGrpCod" = T1."ItmsGrpCod"
WHERE T0."validFor" = 'Y'                                           -- Solo artículos activos
  AND T0."PrchseItem" = 'Y'                                         -- Artículos de compra
ORDER BY T0."ItemCode" ASC;
