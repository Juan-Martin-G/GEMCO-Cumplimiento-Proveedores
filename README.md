# GEMCO · Cumplimiento de proveedores

Panel web que muestra el cumplimiento de proveedores de GEMCO (facturación + backorder contra forecast) a partir del Excel consolidado *Cumplimiento proveedores – Consolidado*. Usa la misma arquitectura que el panel EERR-GEMCO.

**Panel:** https://juan-martin-g.github.io/GEMCO-Cumplimiento-Proveedores/

## Arquitectura

- **Sitio estático** (HTML, JS y Chart.js) en `docs/`, servido por GitHub Pages desde `main` y la carpeta `/docs`. No hay servidor ni Python corriendo en vivo.
- **Datos precalculados:** `generar_datos.py` lee el Excel desde `data/raw/`, recalcula y valida cada fila de las BBDD, aplica la lógica de negocio, concilia contra las hojas de marca, compara contra la versión anterior y deja el resultado en `docs/datos_cumplimiento.json`. El navegador solo lee ese JSON y dibuja.
- **Login:** `docs/index.html` pide un correo autorizado y un código de 6 dígitos enviado con EmailJS. `docs/panel.html` revisa la sesión al cargar y, si no existe, redirige al login.
- **El Excel nunca se sube:** `.gitignore` excluye `data/raw/*.xlsx` y cualquier `.xlsx`, `.xls` o `.csv` del repo, incluidas las subcarpetas.

```
docs/
  index.html               login (correo autorizado + código EmailJS)
  login.js                 configuración de EmailJS, autorizados y flujo del login
  panel.html               panel (Resumen, Ficha, Conciliación, Cambios vs versión anterior, Cómo se calcula)
  app.js                   interfaz del panel; lee datos_cumplimiento.json
  styles.css               estilos, modo claro/oscuro, responsive
  datos_cumplimiento.json  datos precalculados (lo genera generar_datos.py)
data/raw/                  Excel vigente: AAAAMMDD <nombre>.xlsx (NO se sube)
data/raw/old/              versiones anteriores (NO se suben; la más reciente se usa para el reporte de cambios)
reportes/                  cambios_<anterior>_a_<actual>.md con el detalle por documento (NO se sube)
generar_datos.py           recálculo y validación de las BBDD + cálculos + conciliación + cambios → JSON
tests/                     pytest: valores de referencia (20261109), selección, validaciones y comparación
requirements.txt
```

## Actualización semanal

1. Mueve el Excel anterior de `data/raw/` a `data/raw/old/`.
2. Deja el nuevo en `data/raw/` con la fecha al inicio del nombre: `AAAAMMDD Cumplimiento proveedores - Consolidado.xlsx`.
3. En el Excel, extiende las fórmulas de `BBDD Facturacion` y `BBDD BO abiertas` hasta la última fila (también si insertaste filas al medio).
4. Categoriza los folios y productos nuevos en `Diccionario Categorias` (folio en B → categoría en C; producto en G → categoría en H).
5. Guarda en Excel, para que queden calculados los valores de las fórmulas.
6. Corre `python generar_datos.py` y revisa las validaciones y el reporte `reportes/cambios_<anterior>_a_<actual>.md`.
7. Corre los tests: `pytest tests/ -v`.
8. Publica:

```bash
git status            # debe ir docs/datos_cumplimiento.json y NINGÚN .xlsx ni reportes/
git add -A
git commit -m "Datos de cumplimiento al AAAA-MM-DD"
git push
```

> **Nota:** en `Parametros Moneda` los tipos de cambio se escriben **con coma decimal** (por ejemplo `0,00092`), como número. Si queda como texto (por ejemplo `"0.00092"` escrito con punto en un Excel chileno), `generar_datos.py` se detiene e indica la celda.

- **Si algo no cuadra,** `generar_datos.py` termina con error y **no escribe el JSON**, así que el panel sigue mostrando los datos anteriores. La consola indica la fila exacta de Excel, con su folio u OC.
- **Selección del archivo:** se usa el `.xlsx` de `data/raw/` (no recursivo) cuyo nombre empieza con la fecha `AAAAMMDD` más reciente, **según el nombre** y no según la fecha de modificación. Los archivos `~$…` se ignoran y `data/raw/old/` nunca se usa como fuente. El script termina con error si no hay archivos válidos, si la fecha del nombre no es válida o si hay dos archivos con la misma fecha. Si hay más de un `.xlsx` válido, avisa cuál usó y que los demás deberían ir a `old/`.
- **`--permitir-sin-categoria`:** sigue adelante aunque haya documentos sin categoría estándar, mostrando la advertencia en el panel.
- **Tests de referencia:** corresponden al archivo **20261109** (lo buscan en `data/raw/` o `data/raw/old/`). Cuando valides un archivo nuevo, actualiza `ARCHIVO_REFERENCIA` y las tablas `CALENDARIO_CON_FORECAST`, `CALENDARIO_SIN_FORECAST` y `FISCAL` en `tests/test_generar_datos.py`. No ajustes los números para que pasen: si no coinciden, el error está en el código o en el Excel.
- **Instalación (una vez):** `pip install -r requirements.txt`.
- **Para ver el panel en local:** `python -m http.server -d docs` y abre `http://localhost:8000`. Si abres el archivo directo (`file://`), el navegador no permite leer el JSON.

### `INCLUIR_DETALLE`

Al inicio de `generar_datos.py`:

- `INCLUIR_DETALLE = True` (por defecto): el JSON incluye las tablas de **BO abiertas** y de **facturas** (OC, productos, precios, folios, comentarios).
- `INCLUIR_DETALLE = False`: el JSON lleva **solo agregados** (montos por mes, trimestre y categoría, cumplimiento, conciliación) y el panel oculta esas tablas.

## ⚠️ Hoja duplicada: "BBDD Facturacion " (con espacio al final)

| Hoja | Estado | Uso |
|---|---|---|
| `BBDD Facturacion` (sin espacio) | visible | **la correcta**: la usan todas las fórmulas (2025–2026) |
| `BBDD Facturacion ` (con espacio final) | oculta | base antigua de 2025 con otras columnas. **No se lee nunca.** |

El script busca la hoja por **nombre exacto** (sin trim) y valida sus encabezados. Si falta alguno, se detiene con error. `BBDD BO abiertas` pasa por la misma validación. El panel muestra el nombre exacto de la hoja leída, su número de filas y su fecha de corte. Si cambian los nombres de hojas o columnas, ajusta las constantes al inicio de `generar_datos.py`.

## Cálculos (replican el Excel)

| Métrica | Cálculo | Equivalente Excel |
|---|---|---|
| Facturación | Σ `Total moneda marca` de BBDD Facturacion con Proveedor (trim) = marca, y año/mes de la fecha de contabilización | fila 3 |
| Backorder | Σ `Total moneda marca` de BBDD BO abiertas con Marca (trim) = marca, y año/mes de la fecha OC | fila 4 |
| Fact + BO | Facturación + Backorder | fila 5 |
| Forecast | input manual de la fila 6 (mes, trimestre y total) | fila 6 |
| Cumplimiento trimestral | Fact+BO del trimestre / Forecast del trimestre | fila 7 (E, I, M, Q) |
| Cumplimiento anual | Fact+BO total / Forecast total | R7 |
| **Cumplimiento YTD** *(calculada por la app)* | Fact+BO de enero al mes de corte / Forecast de enero al mes de corte (incluye el mes de corte completo) | no existe en el Excel |

- **Hojas de marca:** solo se leen las **visibles** (A2 = "Item" y A6 = "Forecast"). La clave de marca es A1 con trim, no el nombre de la hoja.
- **Bloque fiscal USA** (hoy en Pentax, G-Flex y Durr Dental): se detecta buscando "Item" en la fila 2 a la derecha de la columna S y usa los pares año/mes de sus filas 1–2.
- **Forecast vacío o 0** = sin forecast: se muestra "—", nunca 0 %. Si los 3 meses de un trimestre están vacíos y el trimestre tiene valor, cada mes recibe trimestre/3, marcado como *derivado*.
- **Categorías:** facturación por "Categoria" y BO por "Grupo". Las estándar son Equipos, Repuestos, Accesorios y Consumibles. Cualquier otro valor va a **Sin categoría**, que se muestra y no se descarta.
- **Montos:** se suma siempre "Total moneda marca". Cada marca se reporta en su moneda, y la vista general compara solo porcentajes y conteos.
- **Conversión de moneda:** las marcas con documentos convertidos (Moneda del documento ≠ Moneda marca) muestran el monto original, el TC de "Parametros Moneda" y el monto convertido.
- **Fecha de corte:** es la máxima fecha de cada BBDD. Q4 se marca *parcial* mientras el corte sea anterior a diciembre.

## Recálculo de las BBDD (sin depender de las columnas con fórmula)

Para cada fila con datos de `BBDD Facturacion` y `BBDD BO abiertas`, el script recalcula en Python todo lo que hoy calculan las fórmulas, a partir de los datos crudos:

| Dato | Cómo se recalcula |
|---|---|
| Marca | `Parametros Proveedor`: nombre del proveedor (A) → marca (B), con trim. |
| Año y Mes | Desde la fecha. |
| Categoría | `Diccionario Categorias`. Facturación: folio (B) → categoría (C). BO: producto (G) → categoría (H). Los números y textos se comparan normalizados, así que un folio guardado como texto encuentra al mismo folio numérico. Con códigos repetidos se usa la primera aparición (como XLOOKUP) y se avisa si las categorías difieren. |
| Moneda marca, TC y Total moneda marca | `Parametros Moneda`: A:B (marca → moneda de reporte) y D:E (`ORIGEN>DESTINO` → multiplicador). Si la moneda de la fila está vacía o es igual a la de reporte, TC = 1. En BO, `Total Neto ME = Precio × Cantidad pdte`. |

- **Comparación fila por fila:** lo recalculado se compara contra el valor cacheado de cada columna con fórmula. Si una fila tiene la fórmula vacía o un valor distinto, el script indica la fila exacta de Excel, con su folio u OC, y termina con error. Un caso real: al insertar facturas al medio de la base, quedaron filas sin fórmulas y una marca perdió facturación sin que ninguna hoja de marca lo mostrara.
- **Única excepción:** una fila con Cantidad pdte = 0 y monto 0 donde solo difiere la categoría queda como aviso, porque no mueve ningún número.
- **Filas ignoradas:** las filas sin ningún dato crudo, como las filas del final con fórmulas arrastradas, se ignoran.
- **Panel:** usa los valores recalculados.

## Validaciones (solo marcas con hoja visible)

- **Categorías:** toda la facturación del período del panel (año calendario y, donde exista, año fiscal Jul–Jun) y todo el BO con Cantidad pdte > 0 debe estar en Equipos, Repuestos, Accesorios o Consumibles. Las filas que no cumplen se listan con fila, marca, fecha, folio u OC, producto, descripción y monto. Es error por defecto, o advertencia con `--permitir-sin-categoria`.
- **Monedas:** cada documento debe estar en la moneda de reporte de su marca o tener un tipo de cambio definido. Si no, es error. Ejemplo real: Steelco tiene una línea en CLP (OC 1588/26), que hoy se convierte con `CLP>EUR`.
- **Parametros Moneda:** todos los tipos de cambio deben ser numéricos. Un TC guardado como texto es error, con la celda indicada.
- **Marcas sin hoja:** las marcas sin hoja de marca visible no se validan.

## Conciliación

`generar_datos.py` compara lo recalculado con los valores cacheados de cada hoja de marca visible y de cada bloque (calendario y fiscal). Compara:

- los totales de las filas 3 a 7;
- el cumplimiento de cada trimestre;
- los totales de las filas 15 y 23.

La tolerancia es 0,01 en montos y 0,0001 en %. La pestaña **Conciliación** del panel muestra la tabla completa. Cualquier diferencia, una estructura de hoja inesperada o fechas inválidas bloquean la publicación.

## Cambios vs versión anterior

- **Versión anterior:** es el archivo de fecha más reciente en `data/raw/old/` que sea anterior al actual. Si no hay, el paso se omite con un aviso.
- **Qué se compara:** ambas versiones se calculan con la misma lógica y se comparan por marca visible, año y mes: Facturación, BO, Fact+BO y cumplimientos trimestrales y anuales.

Cada diferencia se explica con los documentos que la causan: folios agregados, eliminados o modificados (con la OC que mencionan en Comentarios) y líneas de OC (Referencia + Producto) con cambio en Cantidad pdte o monto. La suma del detalle explica exactamente la diferencia del mes. Cada documento se clasifica así:

| Clase | Casos |
|---|---|
| ✔ **ESPERABLE** | BO que baja porque se facturó (hay una factura nueva que menciona la OC) · factura nueva posterior al corte anterior · OC nueva |
| ⚠ **REVISAR** | folio nuevo con fecha retroactiva · folio que desaparece, cambia de monto o de fecha · BO que sube en una OC antigua · BO que baja o desaparece sin una factura nueva que mencione la OC · cambio de categoría · cambio de forecast · cambio en `Parametros Moneda` |

- **Tipo de cambio:** un cambio de tipo de cambio, o la incorporación de una marca a `Parametros Moneda`, cambia montos sin que cambien los documentos. Se reporta como causa aparte.
- **Reporte detallado:** queda en `reportes/cambios_<anterior>_a_<actual>.md`, que no se sube.
- **En el JSON:** solo va un resumen por marca y mes, sin OC, productos ni precios, para la pestaña **Cambios vs versión anterior** del panel.

## Control de acceso

1. `index.html` compara el correo (sin distinguir mayúsculas) con `ACCESO.autorizados` en `docs/login.js`.
2. Si el correo está autorizado, genera un código de 6 dígitos que vence en 15 minutos y lo envía con EmailJS. El envío tiene un timeout de 15 segundos y muestra "No se pudo enviar el código: <motivo>" si falla. Hay que esperar 45 segundos entre reenvíos.
3. Con el código correcto, guarda `cumpl_auth` en `sessionStorage` y abre `panel.html`. El flag es **propio de este sitio**: EERR usa `eerr_auth`. Ambos sitios comparten el dominio `juan-martin-g.github.io`, así que con un mismo flag entrar a uno daría acceso al otro.
4. La sesión se pierde con **Cerrar sesión** o al cerrar la pestaña.

Para agregar o quitar personas, edita `ACCESO.autorizados` en `docs/login.js` y haz push. El navegador debe poder llegar a `api.emailjs.com`, y algunas redes o firewalls lo bloquean.

## Limitaciones (decisión informada)

- **El JSON es público.** `docs/datos_cumplimiento.json` se puede descargar directo por su URL (`…/GEMCO-Cumplimiento-Proveedores/datos_cumplimiento.json`) y además está en el repo público con todo su historial. Cualquiera que lo busque ve los números sin pasar por el login. Con `INCLUIR_DETALLE = False` se publica menos (sin OC, productos ni precios), pero los agregados siguen siendo públicos. Los JSON de versiones anteriores quedan en el historial de git.
- **El login es solo una barrera básica,** no seguridad real. Todo se valida en el navegador: el código se genera y se verifica en el cliente, y la lista de correos y la Public Key de EmailJS quedan visibles en el código fuente. Sirve para evitar accesos casuales e indexación (`noindex`), no para proteger los datos frente a alguien decidido.
- **Valores cacheados:** las BBDD se recalculan en Python, pero la conciliación y la comparación fila por fila usan los valores cacheados del Excel. Por eso el archivo debe estar guardado desde Excel con las fórmulas calculadas.
- **Reporte de cambios:** la clasificación ESPERABLE/REVISAR se basa en reglas. Para saber si un BO bajó porque se facturó, busca facturas nuevas que mencionen la OC en Comentarios u "Orden de compra" (formato `1450/26`). Un BO recibido sin factura, o una factura que no menciona la OC, queda como REVISAR.
- **YTD:** incluye el mes de corte completo, aunque ese mes todavía no termine.
- **BO por período:** el BO de cada período considera solo las OC fechadas dentro del período, igual que el Excel.
- **Conexión a internet:** el panel necesita internet para cargar Chart.js (cdnjs) y el SDK de EmailJS (jsDelivr), ambos con verificación SRI.
