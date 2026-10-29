# GEMCO · Cumplimiento de proveedores

Panel web que muestra el cumplimiento de proveedores de GEMCO (facturación + backorder contra forecast) a partir del Excel consolidado *Cumplimiento proveedores – Consolidado*. Usa la misma arquitectura que el panel EERR-GEMCO.

**Panel:** https://juan-martin-g.github.io/GEMCO-Cumplimiento-Proveedores/

## Arquitectura

- **Sitio estático** (HTML, JS y Chart.js) en `docs/`, servido por GitHub Pages desde `main` y la carpeta `/docs`. No hay servidor ni Python corriendo en vivo.
- **Datos precalculados:** `generar_datos.py` lee el Excel desde `data/raw/`, aplica toda la lógica de negocio, concilia contra el Excel y deja el resultado en `docs/datos_cumplimiento.json`. El navegador solo lee ese JSON y dibuja.
- **Login:** `docs/index.html` pide un correo autorizado y un código de 6 dígitos enviado con EmailJS. `docs/panel.html` revisa la sesión al cargar y, si no existe, redirige al login.
- **El Excel nunca se sube:** `.gitignore` excluye `data/raw/*.xlsx` y cualquier `.xlsx`, `.xls` o `.csv` del repo, incluidas las subcarpetas.

```
docs/
  index.html               login (correo autorizado + código EmailJS)
  login.js                 configuración de EmailJS, autorizados y flujo del login
  panel.html               panel (Resumen, Ficha por marca, Conciliación, Cómo se calcula)
  app.js                   interfaz del panel; lee datos_cumplimiento.json
  styles.css               estilos, modo claro/oscuro, responsive
  datos_cumplimiento.json  datos precalculados (lo genera generar_datos.py)
data/raw/                  Excel fuente (NO se sube)
generar_datos.py           lectura del Excel + cálculos + conciliación → JSON
tests/                     pytest: valores de referencia y reglas de negocio
requirements.txt
```

## Actualización mensual

```bash
# 1. Reemplaza el Excel en data/raw/ (guardado desde Excel, con fórmulas calculadas)
# 2. Genera los datos (usa el .xlsx más reciente de data/raw/)
python generar_datos.py
# 3. Corre los tests
pytest tests/ -v
# 4. Revisa que NO vaya ningún .xlsx y publica
git status
git add -A
git commit -m "Datos de cumplimiento al <fecha de corte>"
git push
```

- **Si la conciliación no cuadra,** `generar_datos.py` termina con error y **no escribe el JSON**, así que el panel sigue mostrando los datos anteriores. La consola indica qué celda no cuadra.
- **Tests de referencia:** los valores de referencia de los tests corresponden al archivo **20261026**. Con un Excel nuevo esos tests van a fallar porque los montos cambian. Actualiza las tablas `CALENDARIO_CON_FORECAST`, `CALENDARIO_SIN_FORECAST` y `FISCAL` en `tests/test_generar_datos.py` con los valores validados del nuevo archivo. Los tests de reglas, que usan libros sintéticos, no dependen del Excel.
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

## Conciliación

`generar_datos.py` compara lo recalculado con los valores cacheados de cada hoja de marca visible y de cada bloque (calendario y fiscal). Compara:

- los totales de las filas 3 a 7;
- el cumplimiento de cada trimestre;
- los totales de las filas 15 y 23.

La tolerancia es 0,01 en montos y 0,0001 en %. La pestaña **Conciliación** del panel muestra la tabla completa.

- **Bloquean la publicación** (no se escribe el JSON):
  - cualquier diferencia;
  - una estructura de hoja inesperada;
  - fechas inválidas;
  - "Total moneda marca" o "TC a moneda marca" no numéricos.
- **Se publican con aviso** (banner en el panel): una marca con más de una "Moneda marca" en las BBDD. Hoy pasa con Steelco por la fila 348 de BBDD BO abiertas, que tiene moneda vacía y total 0, así que no afecta ningún monto.

## Control de acceso

1. `index.html` compara el correo (sin distinguir mayúsculas) con `ACCESO.autorizados` en `docs/login.js`.
2. Si el correo está autorizado, genera un código de 6 dígitos que vence en 15 minutos y lo envía con EmailJS. El envío tiene un timeout de 15 segundos y muestra "No se pudo enviar el código: <motivo>" si falla. Hay que esperar 45 segundos entre reenvíos.
3. Con el código correcto, guarda `cumpl_auth` en `sessionStorage` y abre `panel.html`. El flag es **propio de este sitio**: EERR usa `eerr_auth`. Ambos sitios comparten el dominio `juan-martin-g.github.io`, así que con un mismo flag entrar a uno daría acceso al otro.
4. La sesión se pierde con **Cerrar sesión** o al cerrar la pestaña.

Para agregar o quitar personas, edita `ACCESO.autorizados` en `docs/login.js` y haz push. El navegador debe poder llegar a `api.emailjs.com`, y algunas redes o firewalls lo bloquean.

## Limitaciones (decisión informada)

- **El JSON es público.** `docs/datos_cumplimiento.json` se puede descargar directo por su URL (`…/GEMCO-Cumplimiento-Proveedores/datos_cumplimiento.json`) y además está en el repo público con todo su historial. Cualquiera que lo busque ve los números sin pasar por el login. Con `INCLUIR_DETALLE = False` se publica menos (sin OC, productos ni precios), pero los agregados siguen siendo públicos. Los JSON de versiones anteriores quedan en el historial de git.
- **El login es solo una barrera básica,** no seguridad real. Todo se valida en el navegador: el código se genera y se verifica en el cliente, y la lista de correos y la Public Key de EmailJS quedan visibles en el código fuente. Sirve para evitar accesos casuales e indexación (`noindex`), no para proteger los datos frente a alguien decidido.
- **Valores cacheados:** se leen los valores cacheados del Excel, así que el archivo debe estar guardado desde Excel con las fórmulas calculadas.
- **YTD:** incluye el mes de corte completo, aunque ese mes todavía no termine.
- **BO por período:** el BO de cada período considera solo las OC fechadas dentro del período, igual que el Excel.
- **Conexión a internet:** el panel necesita internet para cargar Chart.js (cdnjs) y el SDK de EmailJS (jsDelivr), ambos con verificación SRI.
