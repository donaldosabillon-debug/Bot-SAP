/*
===================================================================================================
 CONSULTA SAP BUSINESS ONE HANA: VERIFICACIÓN POST-ELIMINACIÓN DE CLIENTES
 Empresa: Lazarus & Lazarus
 Propósito:
   Verificar de inmediato cuáles de los clientes enviados a eliminar con el bot fueron
   efectivamente borrados de SAP y cuáles permanecieron en la base por estar vinculados
   a Cotizaciones ("OQUT") o ventas.
===================================================================================================
 INSTRUCCIONES:
 1. Abre el Generador de Consultas en SAP B1.
 2. Pega esta consulta.
 3. En la cláusula IN (...), puedes pegar la lista de códigos que mandaste a borrar con el bot.
===================================================================================================
*/

SELECT 
    T0."CardCode"                                                   AS "Codigo_SN",
    T0."CardName"                                                   AS "Nombre_SN",
    T0."LicTradNum"                                                 AS "RTN",
    T0."validFor"                                                   AS "Activo",
    -- Conteo de Cotizaciones / Ofertas
    (SELECT COUNT(*) FROM "OQUT" WHERE "CardCode" = T0."CardCode")  AS "Cant_Cotizaciones",
    -- Conteo de Facturas
    (SELECT COUNT(*) FROM "OINV" WHERE "CardCode" = T0."CardCode")  AS "Cant_Facturas",
    -- Conteo de Pedidos
    (SELECT COUNT(*) FROM "ORDR" WHERE "CardCode" = T0."CardCode")  AS "Cant_Pedidos",
    -- Diagnóstico de por qué no se borró
    CASE 
        WHEN (SELECT COUNT(*) FROM "OQUT" WHERE "CardCode" = T0."CardCode") > 0 
        THEN 'BLOQUEADO: Ligado a Cotización (Debe ponerse Inactivo)'
        
        WHEN (SELECT COUNT(*) FROM "OINV" WHERE "CardCode" = T0."CardCode") > 0 
          OR (SELECT COUNT(*) FROM "ORDR" WHERE "CardCode" = T0."CardCode") > 0 
        THEN 'BLOQUEADO: Ligado a Ventas/Facturas (Debe ponerse Inactivo)'
        
        ELSE 'PENDIENTE DE REVISIÓN EN SAP'
    END                                                             AS "Motivo_No_Eliminado"
FROM "OCRD" T0
WHERE T0."CardType" = 'C'
  -- Filtra los códigos que intentaste eliminar pero que aún existen en SAP
  AND (
        (SELECT COUNT(*) FROM "OQUT" WHERE "CardCode" = T0."CardCode") > 0
        OR (SELECT COUNT(*) FROM "OINV" WHERE "CardCode" = T0."CardCode") > 0
        OR (SELECT COUNT(*) FROM "ORDR" WHERE "CardCode" = T0."CardCode") > 0
      )
ORDER BY "Cant_Cotizaciones" DESC, "Cant_Facturas" DESC;
