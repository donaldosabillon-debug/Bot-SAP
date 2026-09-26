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

## 👥 2. Bot de Clientes: Actualización Masiva de Socios de Negocios
👉 **Archivo de Inicio:** `ejecutar_bot.bat`

Automatiza la actualización selectiva y masiva de socios de negocios en **"Datos maestros socio de negocios"** con mapeo interactivo de columnas, modelado numérico sin decimales y auto-recuperación ante errores:
- **Cabecera**: Código de Cliente (`card_code`), RTN (`rtn`).
- **Pestaña General**: Teléfono 1 (`telefono`), Teléfono Móvil (`movil`), Correo Electrónico (`correo`), Estado Activo (`activo`).
- **Panel Lateral UDF**: `WBCUSTID` y `SyncFlag` (Sincronización con Zoho CRM).
- **Pestaña Direcciones**: ID de dirección (`id_direccion`), Calle/ Número (`calle_numero`), Ciudad (`ciudad`), Indicador de impuestos (`indicador_impuestos`).

### Características Clave:
1. **Modelado Numérico Sin Errores**: Normaliza automáticamente enteros en teléfonos, RTN, WBCUSTID y direcciones, eliminando decimales residuales de Excel (`.0`) y preservando ceros a la izquierda (vital para el formato RTN de Honduras).
2. **Ciclo de Búsqueda Seguro en RDP**: Utiliza la secuencia probada de productos: Clic en Lupa (`[2297, -48]`) ➔ Pegado de código ➔ Clic en botón inferior *Buscar/Actualizar* (`[2000, 951]`) ➔ Pegado atómico en campos ➔ Clic en *Actualizar*.
3. **Protocolo de Auto-Recuperación ante Errores**: Si SAP rechaza guardar un socio de negocio (por validación fiscal de RTN, dirección duplicada, etc.), el bot no se detiene ni se traba: ejecuta automáticamente **Crear nuevo ➔ Descartar modificaciones ➔ Buscar (Lupa)** y pasa al siguiente cliente.
4. **Prueba con 1 Cliente (`🧪 PROBAR CON 1 CLIENTE`)**: Permite verificar la sincronización de un único registro antes de lanzar la corrida masiva.
5. **Compatibilidad Dual-Monitor**: Diseñado para configuraciones de pantalla extendida (ej. Laptop + Monitor RDP) con exclusión automática de ventanas como WhatsApp.

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
