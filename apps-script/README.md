# Apps Script — Transporte Cinema Shuttle

Dos proyectos de Apps Script:

| Carpeta | Va en el archivo | Para qué |
|---|---|---|
| `machote-solicitud/` | **MACHOTE MAESTRO — Solicitud de Transporte** (y cada copia por proyecto) | Captura de solicitudes: hoja con cascada Vehículo → Proveedor → Placas/Chofer, y formulario (panel lateral + página web para clientes). |
| `consolidador/` | **Tablero Analítico Transportación** | Base de datos maestra: lee la pestaña `Solicitudes Transporte` de todos los archivos de la carpeta y la junta en una sola tabla para reportes. |

---

## 1. Machote (`machote-solicitud/`)

### Instalar
1. Abre el MACHOTE → **Extensiones → Apps Script**.
2. Reemplaza todo el contenido de `Código.gs` con `machote-solicitud/Codigo.gs`.
3. Abre el archivo HTML `Formulario` (o créalo: **+ → HTML**, nombre exacto `Formulario`) y reemplázalo con `machote-solicitud/Formulario.html`.
4. Guarda, recarga la hoja y en el menú **🚐 Machote Transporte** corre **Configuración inicial completa**.
   - La pestaña `Solicitud` se renombra sola a **`Solicitudes Transporte`** (el nombre que lee el consolidador).
5. Captura **Fecha de inicio** (G5). Sin esa fecha el formulario no deja enviar.

### Qué se corrigió del formulario
- Había una letra `E` suelta en el `<script>` del HTML → error de sintaxis → **todo el JavaScript dejaba de funcionar** (se quedaba en "Cargando opciones…" y los botones no hacían nada).
- Las fechas `2026-09-28` se leían en UTC y en México caían un día antes → ahora se calculan por día exacto.
- Si la fecha de inicio estaba vacía el servidor tronaba; ahora avisa.
- Al escribir el vehículo desde el formulario no aparecía la lista de proveedores (los scripts no disparan `onEdit`) → ahora se aplica.
- Candado (`LockService`) para que dos personas enviando al mismo tiempo no escriban en la misma fila.
- Nuevo: nombre de quien solicita, cantidad de unidades (crea N filas), tipo de servicio (Full / Foráneo / Transfer) y horario.

### Usarlo
- **Dentro de la hoja:** menú → **📝 Abrir formulario de captura** (panel lateral).
- **Para clientes (web):** Implementar → Nueva implementación → Aplicación web → Ejecutar como *Yo* → Acceso *Cualquier usuario*.
  - ⚠️ Cada vez que cambies el código: **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva versión**. Si no, la URL `/exec` sigue sirviendo la versión vieja (causa muy común de "no funciona").
  - Cada copia del machote (cada proyecto) tiene su propia URL.

---

## 2. Base maestra (`consolidador/`)

### Instalar
1. Abre **Tablero Analítico Transportación** → **Extensiones → Apps Script**.
2. Reemplaza `Código.gs` con `consolidador/Codigo.gs` (el `Dashboard.html` que ya tienes sigue funcionando).
3. Revisa `FOLDER_IDS` al inicio del código. Ya vienen:
   - `PROYECTOS TRANSPO 2026` (nuevos)
   - `PROYECTOS 2026` (proyectos anteriores)
4. Recarga la hoja → menú **📊 Base Maestra Transporte → ▶ Consolidar ahora**. La primera vez pide permisos.
5. (Opcional) Botón: en la hoja **Panel**, Insertar → Dibujo → rectángulo "CONSOLIDAR" → ⋮ → *Asignar secuencia de comandos* → `consolidarAhora`.

### Cómo decide qué leer
- De cada Google Sheet de las carpetas (y subcarpetas) abre **solo** la pestaña `Solicitudes Transporte`. Las demás (Catálogos, Crew List, Estimate, Cierre, Hoja 1…) se ignoran.
- Busca la fila de encabezados (la que tiene `VEHICULO` y `PROVEEDOR`) y reconoce las columnas **por nombre**, así que sirve para el machote nuevo y para los formatos anteriores (`DIA 13`, `COSTO FINAL`, `PAGO PROVEEDOR`, `CINEMA`, etc.).
- Toma una fila solo si tiene **vehículo + algún dato real** (concepto, proveedor, chofer, placas, días en el calendario o costo). Los renglones vacíos con fórmulas en $0 no entran.
- Se detiene al llegar a los bloques de resumen de abajo (`Proveedor | Total | PO`, `Suma total`, `COSTO TOTAL DEL PROYECTO`…).
- Proyecto / Productor / Mes se leen del encabezado de la hoja; si no hay, se usa el nombre del archivo.

### Incremental
- Solo vuelve a leer los archivos **nuevos o modificados** desde la última corrida. Los archivos que sacas de la carpeta se quitan de la base.
- Para agregar proyectos anteriores: mételos a la carpeta y presiona **Consolidar ahora**.
- Si son muchos, procesa por lotes de ~4.5 min y se reprograma solo hasta terminar (mira la hoja **Panel**).
- **Reprocesar TODO** vuelve a leer todo (úsalo después de editar la hoja *Alias*).

### Pestañas que genera
| Pestaña | Contenido |
|---|---|
| **Panel** | Estado de la última corrida e instrucciones. |
| **BD Solicitudes** | Una fila por unidad solicitada: proyecto, productor, mes, vehículo, concepto, proveedor (limpio y original), chofer, placas, días por tipo, tarifas, costos, cinema, pago proveedor, validaciones, incidencia, observaciones, archivo de origen. |
| **BD Uso Diario** | Una fila por unidad por día de uso (fecha, código, modalidad, estatus). Sirve para reportes por fecha / ocupación. |
| **Registro Archivos** | Qué archivo se leyó, cuántas filas dio, y si le falta la pestaña o dio error. |
| **Alias** | Unifica nombres escritos de varias formas (ej. `CHRISTIAN ANGEL` → `CHRISTIAN ALEXIS ANGEL CHAVEZ`). |
| **Resumen por Proyecto**, **KPI …**, **KPI Resumen** | Tablas resumen (las usa el Dashboard). |
| **Log** | Historial de corridas. |

La pestaña vieja **Datos Consolidados** ya no se usa; puedes borrarla.
