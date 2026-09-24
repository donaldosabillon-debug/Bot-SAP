# Suite de Bots RPA para SAP Business One (HANA)
### Lazarus & Lazarus | Soluciones Automatizadas de Gestión

Esta carpeta contiene dos aplicaciones de automatización RPA ejecutables mediante archivos **`.bat`** de un solo clic:

---

## 📦 1. Bot de Artículos: Planificación (Lead Time, Mínimo y Tolerancia)
👉 **Archivo de Inicio:** `ejecutar_bot_articulos.bat`

Automatiza la actualización masiva de los parámetros de planificación en la ventana de **"Datos maestros de artículo"** (`Módulos > Inventario > Datos maestros de artículo` o `Ctrl + I`):
1. **`Cantidad de pedido mínimo`** (Mínimo de compra).
2. **`Tiempo lead`** (Lead Time en días).
3. **`Días de tolerancia`** (Tiempo de retraso / tolerancia en días).

### Pasos de Uso:
1. Abre tu sesión de **Escritorio Remoto** (`182.160.29.90`) con SAP B1 y deja abierta la ventana de **"Datos maestros de artículo"** (`Ctrl + I`).
2. En tu laptop, haz doble clic en **[`ejecutar_bot_articulos.bat`](file:///c:/Users/luis.sabillon/OneDrive%20-%20Lazarus%20&%20Lazarus/Escritorio/Proyectos/Bot%20Clientes/ejecutar_bot_articulos.bat)**.
3. Carga tu archivo Excel (dispones de [`ejemplo_articulos.xlsx`](file:///c:/Users/luis.sabillon/OneDrive%20-%20Lazarus%20&%20Lazarus/Escritorio/Proyectos/Bot%20Clientes/ejemplo_articulos.xlsx)).
4. **Calibración Rápida (Solo la primera vez)**:
   - Haz clic en `🚀 Asistente Rápido: Calibrar los 4 puntos en secuencia`.
   - Ubica el mouse en tu pantalla de SAP sobre cada elemento cuando la cuenta regresiva te lo pida:
     1. Pestaña *'Datos de Planificación'*
     2. Campo *'Cantidad Pedido Mínimo'*
     3. Campo *'Tiempo Lead'*
     4. Campo *'Días de Tolerancia'*
5. Presiona **`▶ INICIAR ACTUALIZACIÓN DE ARTÍCULOS`**.
6. El bot buscará cada artículo (`Ctrl+F`), cambiará a la pestaña de planificación, digitará los 3 valores y presionará `Actualizar` (`Alt+A`).

---

## 📦 2. Bot de Clientes: Sincronización Zoho y Eliminación de Duplicados
👉 **Archivo de Inicio:** `ejecutar_bot.bat`

Contiene dos modos operativos en la ventana de **"Datos maestros de socio de negocios"**:
- **Modo 1: Sincronización Zoho CRM**: Digita `WBCUSTID` y marca `SyncFlag = 'T'`.
- **Modo 2: Eliminación Masiva de Duplicados**: Busca cada cliente y ejecuta la orden de eliminar (`Alt + D` ➔ `Eliminar` ➔ Confirmar).

---

## 📊 Consultas SQL para SAP Business One HANA

1. 📄 **[`query_articulos_planificacion.sql`](file:///c:/Users/luis.sabillon/OneDrive%20-%20Lazarus%20&%20Lazarus/Escritorio/Proyectos/Bot%20Clientes/query_articulos_planificacion.sql)**:
   - Extrae artículos activos de compra de `OITM` con su Mínimo de Compra, Lead Time y Tolerancia para exportar a Excel.
2. 📄 **[`query_clientes_sin_sincronizar_zoho.sql`](file:///c:/Users/luis.sabillon/OneDrive%20-%20Lazarus%20&%20Lazarus/Escritorio/Proyectos/Bot%20Clientes/query_clientes_sin_sincronizar_zoho.sql)**:
   - Extrae clientes `CLL%` y `CED%` pendientes de sincronizar con prioridad fiscal Honduras (`Destino`).
3. 📄 **[`query_auditoria_rtn_duplicados.sql`](file:///c:/Users/luis.sabillon/OneDrive%20-%20Lazarus%20&%20Lazarus/Escritorio/Proyectos/Bot%20Clientes/query_auditoria_rtn_duplicados.sql)**:
   - Audita clientes con RTN repetido contando Cotizaciones (`OQUT`) y Ventas para clasificar candidatos a eliminar.
4. 📄 **[`query_verificar_eliminados.sql`](file:///c:/Users/luis.sabillon/OneDrive%20-%20Lazarus%20&%20Lazarus/Escritorio/Proyectos/Bot%20Clientes/query_verificar_eliminados.sql)**:
   - Consulta post-eliminación para verificar qué clientes fueron borrados y cuáles sobrevivieron por bloqueos.

---

## 🛑 Control de Emergencia

En cualquiera de los bots:
- Presiona **`ESC`** o mueve el mouse a cualquier esquina de la pantalla para detener la automatización inmediatamente.
