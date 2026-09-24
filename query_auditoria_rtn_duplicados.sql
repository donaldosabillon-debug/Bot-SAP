/*
===================================================================================================
 CONSULTA SAP BUSINESS ONE HANA: AUDITORÍA DE RTN DUPLICADOS Y VINCULACIÓN COMERCIAL
 Empresa: Lazarus & Lazarus
 Propósito:
   Identificar todos los clientes con RTN repetido y auditar con precisión sus documentos:
   - Facturas de Venta ("OINV")
   - Pedidos de Venta   ("ORDR")
   - Entregas           ("ODLN")
   - Cotizaciones / Ofertas de Venta ("OQUT")
 
 DIAGNÓSTICO INTELIGENTE:
   1. "CONSERVAR (Tiene ventas/pedidos)"   -> Clientes con facturación real.
   2. "BLOQUEADO POR COTI (Inactivar)"    -> Clientes sin ventas pero con cotizaciones (SAP NO deja borrarlos).
   3. "LISTO PARA BORRAR (0 historial)"   -> Clientes 100% limpios sin cotizaciones ni ventas (ELIMINABLES).
===================================================================================================
*/

WITH "DuplicadosRTN" AS (
    SELECT 
        "LicTradNum"
    FROM "OCRD"
    WHERE "CardType" = 'C'
      AND "LicTradNum" IS NOT NULL
      AND TRIM("LicTradNum") <> ''
      AND TRIM("LicTradNum") NOT IN ('0', '00000000000000', '99999999999999', 'CF', 'C/F')
    GROUP BY "LicTradNum"
    HAVING COUNT(*) > 1
)
SELECT 
    T0."CardCode"                                                   AS "Codigo_SN",
    T0."CardName"                                                   AS "Nombre_SN",
    T0."LicTradNum"                                                 AS "RTN",
    IFNULL(TRIM(T0."Phone1"), '')                                   AS "Telefono_1",
    T0."CreateDate"                                                 AS "Fecha_Creacion",
    T0."validFor"                                                   AS "Activo",
    IFNULL(T0."U_WBCUSTID", '')                                     AS "WBCUSTID",
    IFNULL(T0."U_SyncFlag", 'F')                                    AS "SyncFlag",
    -- Conteo de Cotizaciones / Ofertas de Venta
    (SELECT COUNT(*) FROM "OQUT" WHERE "CardCode" = T0."CardCode")  AS "Cant_Cotizaciones",
    -- Conteo de Facturas de Venta
    (SELECT COUNT(*) FROM "OINV" WHERE "CardCode" = T0."CardCode")  AS "Cant_Facturas",
    -- Conteo de Pedidos de Venta
    (SELECT COUNT(*) FROM "ORDR" WHERE "CardCode" = T0."CardCode")  AS "Cant_Pedidos",
    -- Conteo de Entregas de Mercancía
    (SELECT COUNT(*) FROM "ODLN" WHERE "CardCode" = T0."CardCode")  AS "Cant_Entregas",
    -- Diagnóstico exacto para la toma de decisiones
    CASE 
        -- Caso 1: Tiene ventas reales (Facturas, Pedidos o Entregas)
        WHEN (
            (SELECT COUNT(*) FROM "OINV" WHERE "CardCode" = T0."CardCode") > 0 
            OR (SELECT COUNT(*) FROM "ORDR" WHERE "CardCode" = T0."CardCode") > 0 
            OR (SELECT COUNT(*) FROM "ODLN" WHERE "CardCode" = T0."CardCode") > 0
        ) THEN 'CONSERVAR (Tiene ventas o pedidos)'

        -- Caso 2: No tiene ventas, pero está ligado a una Cotización (SAP bloquea la eliminación)
        WHEN (SELECT COUNT(*) FROM "OQUT" WHERE "CardCode" = T0."CardCode") > 0 
        THEN 'BLOQUEADO POR COTI (Inactivar en vez de borrar)'

        -- Caso 3: Totalmente limpio (0 cotizaciones, 0 ventas) -> SAP sí permite eliminarlo
        ELSE 'LISTO PARA BORRAR (0 historial comercial)'
    END                                                             AS "Diagnostico_SAP"
FROM "OCRD" T0
INNER JOIN "DuplicadosRTN" T1 ON T0."LicTradNum" = T1."LicTradNum"
WHERE T0."CardType" = 'C'
ORDER BY 
    T0."LicTradNum" ASC, 
    "Cant_Facturas" DESC, 
    "Cant_Cotizaciones" DESC, 
    T0."CreateDate" ASC;
