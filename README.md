# GEMCO · Cumplimiento de proveedores

Página web que muestra de forma visual e interactiva el cumplimiento de proveedores de GEMCO a partir del Excel consolidado *Cumplimiento proveedores – Consolidado*.

**Página:** https://juan-martin-g.github.io/GEMCO-Cumplimiento-Proveedores/

## Qué hace

- Lee el Excel **100 % en el navegador** con [SheetJS](https://sheetjs.com/). El archivo no se sube a ningún servidor, no se guarda y desaparece al cerrar o recargar la pestaña. Además, el panel declara `connect-src 'none'` en su política de seguridad (CSP), así que no puede enviar datos a ningún lado.
- El acceso pide un correo autorizado y un código de 6 dígitos enviado por email (ver [Control de acceso](#control-de-acceso)).
- Recalcula, desde las bases `BBDD Facturacion` y `BBDD BO abiertas`, lo mismo que calculan las hojas de marca del Excel y lo muestra en cuatro vistas:
  1. **Resumen**: una fila por marca con moneda, facturación, BO abierto, Fact+BO, forecast anual, cumplimiento anual (Excel), cumplimiento YTD (calculado por la app) y mini-barras Q1–Q4. Se puede ordenar y filtrar por "solo con forecast".
  2. **Ficha por marca**: KPIs, gráfico mensual (facturación + BO con la línea de forecast), cumplimiento trimestral, desglose por categoría, tablas de BO abiertas y de facturas (con búsqueda, orden y exportación CSV). Si la hoja tiene bloque de año fiscal USA, aparece un selector *Año calendario / Año fiscal USA*.
  3. **Conciliación**: compara lo recalculado con los valores que Excel dejó guardados en cada hoja de marca visible (filas 3 a 7, el cumplimiento de cada trimestre y los totales de las filas 15 y 23). Si algo no cuadra, aparece un aviso arriba de la página.
  4. **Cómo se calcula**: cada métrica y su celda o fórmula equivalente en el Excel. Además, cada KPI tiene un tooltip con su origen.
- Todo en español y con formato numérico chileno (miles con punto y decimales con coma). Tiene modo claro y oscuro y se puede usar en notebook o celular.

## Cómo actualizar

1. Actualiza el Excel como siempre y **guárdalo desde Excel** para que queden calculados los valores de las fórmulas.
2. Abre la página (o `index.html` en local), ingresa con tu correo autorizado y el código que llega por email, y arrastra o selecciona el `.xlsx`.
3. Revisa la pestaña **Conciliación**: todo debería aparecer con ✔.

No hay que subir nada al repositorio. Los datos nunca se commitean: `.gitignore` excluye `*.xlsx`, `*.xls`, `*.csv` y `data/`.

## ⚠️ Hoja duplicada: "BBDD Facturacion " (con espacio al final)

El libro tiene dos hojas con nombre casi idéntico:

| Hoja | Estado | Uso |
|---|---|---|
| `BBDD Facturacion` (sin espacio) | visible | **la correcta**: es la que usan todas las fórmulas (2025–2026) |
| `BBDD Facturacion ` (con espacio final) | oculta | base antigua de 2025 con otras columnas (*Info.detallada*, *Saldo vencido (ML)*, *Importe Inicial*…). **No se lee nunca.** |

La app busca la hoja por su **nombre exacto** (sin aplicar trim) y valida que tenga los encabezados esperados. Si falta alguno, detiene la carga con un mensaje de error. La cabecera de la página muestra el nombre exacto de la hoja que leyó y su número de filas, junto a la fecha de corte. `BBDD BO abiertas` pasa por la misma validación.

Si en algún momento cambias los nombres de las hojas o de las columnas, ajusta `CONFIG.hojas` y `CONFIG.encabezados` al inicio de `app.js`.

## Conversión de moneda

- Cada marca se reporta en su **"Moneda marca"**. La app suma siempre la columna **"Total moneda marca"** (nunca "Total documento (ME)" ni "Total Neto ME"), igual que las hojas de marca.
- La hoja `Parametros Moneda` define:
  - en **A:B**, la moneda de reporte de cada marca (hoy: G-Flex → EUR, Bien Air → USD y Normmed → EUR);
  - en **D:E**, el tipo de cambio multiplicador por par (por ejemplo, `USD>EUR` y `EUR>USD`). Los valores vigentes se ven en la ficha de cada marca convertida y en "Cómo se calcula".
  Las marcas que no aparecen en esa tabla quedan en su moneda original (TC = 1).
- El Excel ya trae cada documento convertido (`Total moneda marca = monto original × TC a moneda marca`). La app no convierte por su cuenta: muestra y verifica la conversión.
- En las marcas que tienen documentos en otra moneda (hoy G-Flex, Bien Air y Normmed), la ficha muestra el recuadro **Conversión de moneda** con el monto original por moneda, el TC usado y el monto convertido. Sus KPIs llevan un tooltip que lo indica.
- **Nunca se suman montos de marcas distintas.** El resumen general compara solo porcentajes y conteos.
- Validaciones: cada marca visible debe tener **una sola** "Moneda marca" en ambas BBDD, y todos los "TC a moneda marca" deben ser numéricos. Si no se cumple, aparece un error en Conciliación.

## Cálculos

| Métrica | Cálculo | Equivalente Excel |
|---|---|---|
| Facturación | Σ `Total moneda marca` de BBDD Facturacion con Proveedor (trim) = marca, y año y mes de la fecha de contabilización | fila 3 |
| Backorder | Σ `Total moneda marca` de BBDD BO abiertas con Marca (trim) = marca, y año y mes de la fecha OC | fila 4 |
| Fact + BO | Facturación + Backorder | fila 5 |
| Forecast | input manual de la fila 6 (mes, trimestre y total) | fila 6 |
| Cumplimiento trimestral | Fact+BO del trimestre / Forecast del trimestre | fila 7 (E, I, M, Q) |
| Cumplimiento anual | Fact+BO total / Forecast total | R7 |
| **Cumplimiento YTD** *(calculada por la app)* | Fact+BO de enero al mes de corte / Forecast de enero al mes de corte | no existe en el Excel |

- La **clave de marca** es el valor de A1 con trim, no el nombre de la hoja (por ejemplo, la hoja "Creomedical" dice "Creo Medical" en A1).
- **Bloque fiscal USA** (julio a junio, hoy en Pentax, G-Flex y Durr Dental): se detecta buscando "Item" en la fila 2 a la derecha de la columna S y usa los pares año/mes de sus filas 1–2.
- **Forecast vacío o 0** = "sin forecast": se muestra "—", nunca 0 %. Si las celdas mensuales de un trimestre están vacías y el trimestre tiene valor, el forecast mensual es trimestre/3 y se marca como *derivado* (línea punteada).
- **Categorías**: facturación por "Categoria" y BO por "Grupo". Las estándar son Equipos, Repuestos, Accesorios y Consumibles. Cualquier otro valor (0, vacío, "ERROR", "No aplica Meta", "Trazabilidad"…) va a **Sin categoría**, que se muestra y no se descarta. Los totales de las filas 15 y 23 del Excel suman solo las 4 estándar.
- **Fecha de corte**: es la fecha máxima de cada BBDD y se muestran ambas. El *mes de corte* del YTD es el de la fecha más reciente entre las dos.
- **Q4 parcial**: un trimestre se marca *parcial* cuando la fecha de corte cae dentro de él y antes de su último mes.

## Control de acceso

Usa el mismo esquema que el panel EERR:

1. `index.html` pide el correo y lo compara (sin distinguir mayúsculas) contra la lista `ACCESO.autorizados` de `login.js`. Si no está en la lista, muestra "Correo no autorizado".
2. Si está autorizado, genera un código de 6 dígitos que vence en 15 minutos y lo envía con EmailJS (plantilla con las variables `{{to_email}}` y `{{codigo}}`). Hay que esperar 45 segundos entre reenvíos.
3. Con el código correcto, guarda la sesión en `sessionStorage` (claves `gcp_*`, distintas de las de EERR) y abre `panel.html`. El panel vuelve al login si no hay sesión. La sesión se pierde con **Cerrar sesión** o al cerrar la pestaña.

Para agregar o quitar personas, edita `ACCESO.autorizados` en `login.js` y haz push.

**Limitaciones (es una barrera básica, no seguridad real):** todo se valida en el navegador. El código se genera y se verifica en el cliente, y la lista de correos y la Public Key de EmailJS quedan visibles en el código fuente. Cualquiera puede abrir `panel.html` copiando una sesión falsa en `sessionStorage`. Aun así, en este proyecto **no hay datos que exponer**: el sitio y el repo no contienen números, y el panel no muestra nada hasta que alguien carga su propio Excel. El login sirve para evitar accesos casuales e indexación (`noindex`).

## Supuestos y limitaciones

- Solo se leen las **hojas de marca visibles** (A2 = "Item" y A6 = "Forecast"). Las hojas ocultas se ignoran por completo, porque usan columnas antiguas y no están vigentes.
- La app lee los **valores cacheados** del Excel. Si el archivo se generó sin recalcular (por ejemplo, desde otra herramienta), la conciliación puede mostrar diferencias.
- El mes y el año se derivan de la columna de fecha, lo que equivale al `TEXT(fecha,"mmmm")` y al `YEAR()` del Excel. La conciliación verifica que coincidan con las columnas Año y Mes cacheadas.
- El YTD incluye el mes de corte **completo**, aunque ese mes todavía no termine. Por eso, a comienzos de mes puede subestimar el cumplimiento.
- El BO de cada período considera solo las OC fechadas dentro del período, igual que el Excel. En la tabla de BO de la ficha se pueden incluir las OC de otros períodos.
- Tolerancias de conciliación: 0,01 en montos y 0,0001 en porcentajes (configurables en `CONFIG.tolerancia`).
- Requiere conexión a internet para descargar SheetJS y Chart.js (cdnjs) y el SDK de EmailJS (jsDelivr), todos con verificación de integridad SRI. El Excel nunca sale del equipo.

## Estructura

```
index.html   login (correo autorizado + código por email)
login.js     configuración de EmailJS, lista de autorizados y flujo del login
panel.html   la aplicación y la vista "Cómo se calcula"
app.js       CONFIG (semáforo, tolerancias, hojas) + motor de cálculo + interfaz
styles.css   estilos, modo claro/oscuro y diseño responsive
```

Sin paso de compilación: funciona abriendo `index.html` en local o desde GitHub Pages.
