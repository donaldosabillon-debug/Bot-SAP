/*
===================================================================================================
 CONSULTA NATIVA SAP BUSINESS ONE SOBRE SAP HANA (HDB) - LOCALIZACIÓN HONDURAS
 Empresa: Lazarus & Lazarus
 Propósito: Extracción de clientes para Zoho CRM.
 Filtros de Negocio:
   1. Tipo Socio: Solo Clientes ("CardType" = 'C').
   2. Series de Código: Solo códigos que comiencen con 'CLL' o 'CED'.
   3. Exclusiones Específicas:
      - CED-000087 (Medina De Lazarus Elena Maria)
      - CED-000088 (Lazarus Medina Ernesto)
      - CED-000131 (Lazarus Y Lazarus Eua)
      - CED-000139
   4. Pendientes: Sin ID de Zoho ("U_WBCUSTID" vacío/nulo) y "U_SyncFlag" <> 'T'.
   5. Estado: Clientes activos ("validFor" = 'Y').
   6. Prioridad Dirección: "DESTINO" ('S' / "ShipToDef") que rige el ISV en Honduras.
===================================================================================================
*/

SELECT 
    T0."CardCode"                                                   AS "Codigo_SN",
    T0."CardName"                                                   AS "Nombre_SN",
    IFNULL(TRIM(T0."Phone1"), '')                                   AS "Telefono_1",
    CASE 
        WHEN T0."Cellular" IS NOT NULL AND TRIM(T0."Cellular") <> '' 
        THEN TRIM(T0."Cellular")
        ELSE IFNULL(TRIM(T0."Phone1"), '')
    END                                                             AS "Telefono_Movil",
    IFNULL(TRIM(T0."E_Mail"), '')                                   AS "Correo_Electronico",
    IFNULL(TRIM(T0."LicTradNum"), '')                               AS "RTN",
    IFNULL(TRIM(T0."CntctPrsn"), '')                                AS "Persona_Contacto",
    COALESCE(
        NULLIF(TRIM(T1."Address"), ''), 
        NULLIF(TRIM(T2."Address"), ''), 
        NULLIF(TRIM(T3."Address"), ''), 
        IFNULL(TRIM(T0."Address"), '')
    )                                                               AS "Direccion",
    COALESCE(
        NULLIF(TRIM(T1."City"), ''), 
        NULLIF(TRIM(T2."City"), ''), 
        NULLIF(TRIM(T3."City"), ''), 
        IFNULL(TRIM(T0."City"), '')
    )                                                               AS "Ciudad"
FROM "OCRD" T0
-- 1. PRIORIDAD PRINCIPAL HONDURAS: Destino ('S') donde se configura el ISV
LEFT JOIN "CRD1" T1 ON T0."CardCode" = T1."CardCode" 
                   AND T1."Address" = T0."ShipToDef" 
                   AND T1."AdresType" = 'S'
-- 2. Fallback secundario: Facturación ('B')
LEFT JOIN "CRD1" T2 ON T0."CardCode" = T2."CardCode" 
                   AND T2."Address" = T0."BillToDef" 
                   AND T2."AdresType" = 'B'
-- 3. Fallback terciario: Cualquier dirección en CRD1 priorizando Destino ('S')
LEFT JOIN (
    SELECT 
        "CardCode",
        "Address",
        "City",
        ROW_NUMBER() OVER (
            PARTITION BY "CardCode" 
            ORDER BY 
                CASE WHEN "AdresType" = 'S' THEN 1 ELSE 2 END,
                CASE WHEN "City" IS NOT NULL AND TRIM("City") <> '' THEN 1 ELSE 2 END,
                "LineNum" ASC
        ) AS "rn"
    FROM "CRD1"
) T3 ON T0."CardCode" = T3."CardCode" AND T3."rn" = 1
WHERE T0."CardType" = 'C'                                           -- Solo Clientes
  AND (
        T0."CardCode" LIKE 'CLL%' 
        OR T0."CardCode" LIKE 'CED%'
      )                                                             -- Solo series CLL y CED
  AND T0."CardCode" NOT IN (
        'CED-000087', 
        'CED-000088', 
        'CED-000131', 
        'CED-000139'
      )                                                             -- Clientes excluidos explícitamente
  AND (
        T0."U_WBCUSTID" IS NULL 
        OR T0."U_WBCUSTID" = '' 
        OR TRIM(T0."U_WBCUSTID") = ''
      )                                                             -- Sin ID interno de Zoho
  AND (
        T0."U_SyncFlag" IS NULL 
        OR T0."U_SyncFlag" <> 'T'
      )                                                             -- Bandera distinta de 'T' (True)
  AND T0."validFor" = 'Y'                                           -- Solo clientes activos
ORDER BY T0."CardCode" ASC;
