"""
Tests de generar_datos.py

- Regresión con los valores de referencia validados contra el Excel 20261109.
  Requieren ese archivo en data/raw/ o data/raw/old/; si no está, se omiten.
- Selección del archivo, validaciones, reglas de negocio y comparación entre
  versiones con libros sintéticos (no requieren el Excel real).

    pytest tests/ -v
"""

import json
import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pytest

import generar_datos as gd

RAIZ = Path(__file__).resolve().parent.parent
ARCHIVO_REFERENCIA = "20261109"


# ──────────────────────────────────────────────────────────────────────
# Excel real de referencia (20261109)
# ──────────────────────────────────────────────────────────────────────

def _excel_referencia():
    raw = RAIZ / "data" / "raw"
    carpetas = [raw] + [p for p in raw.iterdir() if p.is_dir()] if raw.is_dir() else []
    for c in carpetas:
        for p in c.glob(f"{ARCHIVO_REFERENCIA}*.xlsx"):
            if not p.name.startswith("~$"):
                return p
    return None


@pytest.fixture(scope="module")
def datos():
    ruta = _excel_referencia()
    if ruta is None:
        pytest.skip(f"No está el Excel {ARCHIVO_REFERENCIA} en data/raw/ ni en data/raw/old/.")
    return gd.procesar_libro(ruta, incluir_detalle=True)


def marca(datos, nombre):
    return next(m for m in datos["marcas"] if m["nombre"] == nombre)


def bloque(datos, nombre, tipo="calendario"):
    return next(b for b in marca(datos, nombre)["bloques"] if b["tipo"] == tipo)["calc"]


def pct(v, decimales=1):
    return None if v is None else round(v * 100, decimales)


# marca, moneda, fact, bo, forecast, cumpl anual, trimestres (None = sin forecast)
CALENDARIO_CON_FORECAST = [
    ("Steelco", "EUR", 1557639, 180248.93, 1250000, 139.0, [86.4, 88.9, 199.5, 181.3]),
    ("Melag", "EUR", 277105, 13851, 520200, 55.9, [146.3, 141.9, 30.9, 28.9]),
    ("Pentax", "USD", 773085, 33787.33, 1249999.2, 64.5, [97.2, 60.4, 87.1, 13.5]),
    ("Eurosets", "EUR", 686857, 531650, 800000, 152.3, [71.5, 97.4, 139.8, 300.6]),
    ("Creo Medical", "USD", 76441, 10688, 90000, 96.8, [41.6, None, None, None]),
    ("Bien Air", "USD", 58514.4, 0, 50400, 116.1, [104.3, 83.5, 249.6, 27.1]),
]

CALENDARIO_SIN_FORECAST = [
    ("DDC Dolphin", 53479, 0),
    ("Yipak", 147402, 0),
    ("Franklab", 3468, 11376),
    ("G-Flex", 76662.8, 5012.95),
    ("Durr Dental", 89614, 0),
    ("Normmed", 34522.1, 0),
]

# Fiscal Jul-25 – Jun-26: marca, fact, bo, cumpl anual, trimestres
FISCAL = [
    ("Pentax", 1211512, 3710, 97.2, [62.6, 168.7, 97.2, 60.4]),
    ("G-Flex", 149250.4, 0, None, [None, None, None, None]),
    ("Durr Dental", 41870, 0, None, [None, None, None, None]),
]


@pytest.mark.parametrize("nombre,moneda,fact,bo,fc,cumpl,trims", CALENDARIO_CON_FORECAST)
def test_calendario_con_forecast(datos, nombre, moneda, fact, bo, fc, cumpl, trims):
    c = bloque(datos, nombre)
    assert marca(datos, nombre)["moneda"] == moneda
    assert c["total"]["fact"] == pytest.approx(fact, abs=0.01)
    assert c["total"]["bo"] == pytest.approx(bo, abs=0.01)
    assert c["total"]["fc"] == pytest.approx(fc, abs=0.01)
    assert pct(c["total"]["cumpl"]) == cumpl
    assert [pct(t["cumpl"]) for t in c["trims"]] == trims


@pytest.mark.parametrize("nombre,fact,bo", CALENDARIO_SIN_FORECAST)
def test_calendario_sin_forecast(datos, nombre, fact, bo):
    c = bloque(datos, nombre)
    assert c["total"]["fact"] == pytest.approx(fact, abs=0.01)
    assert c["total"]["bo"] == pytest.approx(bo, abs=0.01)
    assert c["tieneForecast"] is False
    # Sin forecast = "—" (None), nunca 0 %
    assert c["total"]["cumpl"] is None and c["ytd"]["cumpl"] is None
    assert all(t["cumpl"] is None for t in c["trims"])


@pytest.mark.parametrize("nombre,fact,bo,cumpl,trims", FISCAL)
def test_bloque_fiscal(datos, nombre, fact, bo, cumpl, trims):
    m = marca(datos, nombre)
    b = next(x for x in m["bloques"] if x["tipo"] == "fiscal")
    assert b["titulo"] == "Año fiscal USA (Jul-25 – Jun-26)"
    c = b["calc"]
    assert c["total"]["fact"] == pytest.approx(fact, abs=0.01)
    assert c["total"]["bo"] == pytest.approx(bo, abs=0.01)
    assert pct(c["total"]["cumpl"]) == cumpl
    assert [pct(t["cumpl"]) for t in c["trims"]] == trims


def test_columna_etiquetas_fiscal(datos):
    col = {m["nombre"]: gd.col_letra(b["c0"]) for m in datos["marcas"] for b in m["bloques"] if b["tipo"] == "fiscal"}
    assert col == {"Pentax": "U", "G-Flex": "V", "Durr Dental": "U"}


def test_conciliacion_y_validaciones(datos):
    r = datos["resumen"]
    # 253 celdas por bloque (filas 3–7, 11–15 y 19–23) × 15 bloques (12 calendario + 3 fiscales)
    assert r["comparaciones"] == 253 * 15
    assert r["diferencias"] == 0
    assert r["erroresBloqueantes"] == 0


BLOQUES = [(n, "calendario") for n in [
    "Steelco", "Melag", "DDC Dolphin", "Yipak", "Franklab", "Pentax", "Creo Medical",
    "G-Flex", "Bien Air", "Durr Dental", "Eurosets", "Normmed"]] + [("Pentax", "fiscal"), ("G-Flex", "fiscal"), ("Durr Dental", "fiscal")]


@pytest.mark.parametrize("nombre,tipo", BLOQUES)
def test_vista_excel_cuadra_celda_a_celda(datos, nombre, tipo):
    b = next(x for x in marca(datos, nombre)["bloques"] if x["tipo"] == tipo)
    assert b["celdas"]["n"] == 253
    assert b["celdas"]["fallas"] == []
    vx = b["vista"]
    assert [c["etiqueta"] for c in vx["columnas"]][:5] == (["Ene", "Feb", "Mar", "Q1", "Abr"] if tipo == "calendario" else ["Jul", "Ago", "Sep", "Q1", "Oct"])
    assert [f["fila"] for f in vx["filas"]] == [3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 19, 20, 21, 22, 23]
    # Totales de la vista = totales de la ficha
    fila = {f["clave"]: f for f in vx["filas"]}
    assert fila["fact"]["valores"][-1] == pytest.approx(b["calc"]["total"]["fact"])
    assert fila["bo_total"]["valores"][-1] + 0 == pytest.approx(sum(b["calc"]["catB"][c] for c in gd.CATEGORIAS))
    if not b["calc"]["tieneForecast"]:
        assert all(v is None for v in fila["cumpl"]["valores"])        # sin forecast: "—", nunca 0 %


def test_lee_hoja_exacta_e_ignora_filas_vacias(datos):
    assert datos["meta"]["fact"]["hoja"] == "BBDD Facturacion"
    assert datos["meta"]["fact"]["filas"] == 722
    assert datos["meta"]["bo"]["filas"] == 1445          # sin las filas finales que solo tienen fórmulas
    assert datos["meta"]["fechaArchivo"] == "2026-11-09"
    assert any("BBDD␣Facturacion␣" in p["titulo"] for p in datos["problemas"])


def test_solo_hojas_de_marca_visibles(datos):
    assert [m["nombre"] for m in datos["marcas"]] == [
        "Steelco", "Melag", "DDC Dolphin", "Yipak", "Franklab", "Pentax", "Creo Medical",
        "G-Flex", "Bien Air", "Durr Dental", "Eurosets", "Normmed"]
    assert marca(datos, "Creo Medical")["hoja"] == "Creomedical"
    assert marca(datos, "Bien Air")["hoja"] == "Bien Air "


def test_steelco_clp_se_convierte(datos):
    m = marca(datos, "Steelco")
    assert m["convierte"] is True
    assert {"orig": "CLP", "dest": "EUR", "tc": 0.00092} in m["paresConversion"]
    linea = next(r for r in m["detalle"]["bo"] if r["referencia"] == "1588/26" and r["producto"] == "R010037")
    assert linea["montoOrig"] == 287600 and linea["total"] == pytest.approx(264.592)


def test_ytd_es_enero_a_mes_de_corte(datos):
    c = bloque(datos, "Steelco")
    y = c["ytd"]
    assert y["nMeses"] == 11 and y["hasta"] == {"anio": 2026, "mes": 10}
    assert y["cumpl"] == pytest.approx(y["factbo"] / y["fc"])
    assert [t["estado"] for t in c["trims"]] == ["cerrado", "cerrado", "cerrado", "parcial"]


def test_sin_detalle_no_incluye_oc_productos_ni_precios():
    ruta = _excel_referencia()
    if ruta is None:
        pytest.skip("No está el Excel de referencia.")
    d = gd.procesar_libro(ruta, incluir_detalle=False)
    assert all("detalle" not in m for m in d["marcas"])
    texto = json.dumps(d, ensure_ascii=False, default=str)
    for campo in ('"referencia"', '"producto"', '"precio"', '"folio"', '"descripcion"', "1588/26", "R010037"):
        assert campo not in texto


# ──────────────────────────────────────────────────────────────────────
# Selección del archivo
# ──────────────────────────────────────────────────────────────────────

def _tocar(carpeta: Path, *nombres):
    carpeta.mkdir(parents=True, exist_ok=True)
    for n in nombres:
        (carpeta / n).write_bytes(b"")


def test_seleccion_elige_el_mas_reciente_por_nombre(tmp_path):
    _tocar(tmp_path, "20261109 Consolidado.xlsx", "20261026 Consolidado.xlsx", "otro.xlsx")
    # La fecha de modificación no importa: se toca el más antiguo al final.
    (tmp_path / "20261026 Consolidado.xlsx").write_bytes(b"x")
    sel = gd.seleccionar_actual(tmp_path)
    assert sel["ruta"].name == "20261109 Consolidado.xlsx"
    assert sel["fecha"] == date(2026, 11, 9)
    assert len(sel["avisos"]) == 1 and "old" in sel["avisos"][0]


def test_seleccion_ignora_old_y_temporales(tmp_path):
    _tocar(tmp_path, "20261026 Consolidado.xlsx", "~$20261231 bloqueo.xlsx")
    _tocar(tmp_path / "old", "20261231 Futuro.xlsx")
    sel = gd.seleccionar_actual(tmp_path)
    assert sel["ruta"].name == "20261026 Consolidado.xlsx"
    assert sel["avisos"] == []


def test_seleccion_falla_con_fechas_duplicadas(tmp_path):
    _tocar(tmp_path, "20261109 A.xlsx", "20261109 B.xlsx")
    with pytest.raises(gd.ErrorCarga, match="misma fecha"):
        gd.seleccionar_actual(tmp_path)


def test_seleccion_falla_sin_archivos(tmp_path):
    _tocar(tmp_path, "Consolidado.xlsx")
    _tocar(tmp_path / "Old", "20261026 A.xlsx")
    with pytest.raises(gd.ErrorCarga, match="No hay archivos válidos"):
        gd.seleccionar_actual(tmp_path)


def test_seleccion_falla_con_fecha_invalida(tmp_path):
    _tocar(tmp_path, "20261341 Consolidado.xlsx")
    with pytest.raises(gd.ErrorCarga, match="no es válida"):
        gd.seleccionar_actual(tmp_path)


def test_seleccion_anterior_en_old(tmp_path):
    _tocar(tmp_path, "20261109 A.xlsx")
    _tocar(tmp_path / "Old", "20261026 A.xlsx", "20261012 A.xlsx", "20261201 Futuro.xlsx")
    ant = gd.seleccionar_anterior(date(2026, 11, 9), tmp_path)
    assert ant["ruta"].name == "20261026 A.xlsx"
    assert gd.seleccionar_anterior(date(2026, 1, 1), tmp_path) is None


# ──────────────────────────────────────────────────────────────────────
# Libro sintético
# ──────────────────────────────────────────────────────────────────────

MESES_HOJA = ["Enero", "Febrero", "Marzo", " Q1 ", "Abril", "Mayo", "Junio", " Q2 ", "Julio", "Agosto",
              "Septiembre", " Q3 ", "Octubre", "Noviembre", "Diciembre", " Q4 ", " Total "]
PROVEEDORES = {"MARCA X SA": "Marca X ", "OTRA SA": "Otra"}
FOLIOS = {1: ("Equipos", 2026), 2: (None, 2026), 3: ("Repuestos", 2026), 4: ("Repuestos", 2026), 5: ("Consumibles", 2026)}
PRODUCTOS = {"P1": "Repuestos", "P2": "Equipos"}


def _fact(folio, fecha, monto, nombre="MARCA X SA", moneda="EUR", comentarios=""):
    return {"nombre": nombre, "fecha": fecha, "folio": folio, "monto": monto, "moneda": moneda, "comentarios": comentarios}


def _bo(oc, producto, fecha, cantidad, precio, nombre="MARCA X SA", moneda="EUR"):
    return {"nombre": nombre, "fecha": fecha, "oc": oc, "producto": producto, "cantidad": cantidad, "precio": precio, "moneda": moneda}


def _moneda_tc(moneda, marca, monedas, pares):
    mm = monedas.get(marca.strip(), moneda)
    if not moneda or moneda == mm:
        return mm, 1
    tc = pares.get(f"{moneda}>{mm}", "FALTA TC")
    return mm, tc


def construir_libro(ruta, facturas, bo, monedas=None, pares=None, borrar=None, hoja_fact="BBDD Facturacion"):
    """Escribe un libro con los valores que Excel habría cacheado (borrar=(tipo, i, col) deja una celda vacía)."""
    monedas = {"Marca X": "EUR"} if monedas is None else monedas
    pares = {"USD>EUR": 0.9} if pares is None else pares
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = hoja_fact
    ws.append(gd.ENCABEZADOS_FACTURACION)
    for i, f in enumerate(facturas):
        marca = PROVEEDORES.get(f["nombre"], "")
        cat, anio_oc = FOLIOS.get(f["folio"], ("ERROR", "error"))
        mm, tc = _moneda_tc(f["moneda"], marca, monedas, pares)
        fila = {"nombre": f["nombre"], "marca": marca, "fecha": f["fecha"], "anio": f["fecha"].year,
                "mes": gd.MESES[f["fecha"].month - 1], "folio": f["folio"], "comentarios": f["comentarios"], "oc": None,
                "anio_oc": anio_oc, "cat": 0 if cat is None else cat, "moneda_orig": f["moneda"], "monto_orig": f["monto"],
                "moneda_marca": mm, "tc": tc, "total": f["monto"] * tc if isinstance(tc, (int, float)) else "#VALUE!"}
        if borrar and borrar[0] == "fact" and borrar[1] == i:
            fila[borrar[2]] = None
        ws.append([fila[k] for k in gd.COLUMNAS_FACT])
    wbo = wb.create_sheet("BBDD BO abiertas")
    wbo.append(gd.ENCABEZADOS_BO)
    for i, b in enumerate(bo):
        marca = PROVEEDORES.get(b["nombre"], "")
        mm, tc = _moneda_tc(b["moneda"], marca, monedas, pares)
        neto = b["cantidad"] * b["precio"]
        fila = {"fecha": b["fecha"], "anio": b["fecha"].year, "mes": gd.MESES[b["fecha"].month - 1], "nombre": b["nombre"],
                "referencia": b["oc"], "producto": b["producto"], "descripcion": "Prod", "cat": PRODUCTOS.get(b["producto"], "ERROR"),
                "marca": marca, "cantidad": b["cantidad"], "moneda_orig": b["moneda"], "precio": b["precio"], "monto_orig": neto,
                "moneda_marca": mm, "tc": tc, "total": neto * tc if isinstance(tc, (int, float)) else "#VALUE!"}
        if borrar and borrar[0] == "bo" and borrar[1] == i:
            fila[borrar[2]] = None
        wbo.append([fila[k] for k in gd.COLUMNAS_BO])
    # Filas finales con fórmulas pero sin datos: deben ignorarse
    wbo.append([None, 2026, "enero", None, None, None, None, "ERROR", "", None, None, None, 0, 0, 1, 0])

    pp = wb.create_sheet("Parametros Proveedor")
    pp.append(["Nombre de cliente/proveedor", "Proveedor"])
    for k, v in PROVEEDORES.items():
        pp.append([k, v])
    dic = wb.create_sheet("Diccionario Categorias")
    dic["B1"], dic["G1"] = "BBDD Facturacion", "BBDD BO"
    dic["B3"], dic["C3"], dic["D3"], dic["G3"], dic["H3"] = "Numero de Folio", "Categoria", "Año OC", "Producto", "Categoria"
    for r, (folio, (cat, anio)) in enumerate(FOLIOS.items(), start=4):
        dic.cell(r, 2, folio), dic.cell(r, 3, cat), dic.cell(r, 4, anio)
    for r, (prod, cat) in enumerate(PRODUCTOS.items(), start=4):
        dic.cell(r, 7, prod), dic.cell(r, 8, cat)
    pm = wb.create_sheet("Parametros Moneda")
    pm.append(["Marca", "Moneda reporte", None, "Par", "Tipo de cambio"])
    filas_pm = max(len(monedas), len(pares))
    for i in range(filas_pm):
        mk = list(monedas.items())[i] if i < len(monedas) else (None, None)
        pr = list(pares.items())[i] if i < len(pares) else (None, None)
        pm.append([mk[0], mk[1], None, pr[0], pr[1]])

    hm = wb.create_sheet("MarcaX")
    hm["A1"], hm["A2"], hm["A6"] = "Marca X", " Item ", "Forecast "
    for i, etiqueta in enumerate(MESES_HOJA, start=2):
        hm.cell(row=1, column=i, value=2026)
        hm.cell(row=2, column=i, value=etiqueta)
    hm["E6"], hm["R6"] = 300, 300            # forecast solo trimestral en Q1
    for r, etiqueta in zip([11, 12, 13, 14, 15, 19, 20, 21, 22, 23], ["EQUIPOS", "REPUESTOS", "ACCESORIOS", "CONSUMIBLES", "Total"] * 2):
        hm[f"A{r}"] = etiqueta
    oc = wb.create_sheet("Oculta")
    oc["A1"], oc["A2"], oc["A6"] = "Otra", " Item ", "Forecast "
    oc.sheet_state = "hidden"
    wb.save(ruta)
    return ruta


FACTURAS_BASE = [
    _fact(1, datetime(2026, 2, 10), 100),
    _fact(1, datetime(2026, 2, 10), 999, nombre="OTRA SA"),        # otra marca
    _fact(3, datetime(2026, 3, 5), 7),
]
BO_BASE = [_bo("0001/26", "P1", datetime(2026, 5, 5), 2, 25)]


def _bloqueantes(res, conciliacion=False):
    # La hoja de marca sintética no tiene valores cacheados: salvo que se pida, se ignora la conciliación celda a celda.
    return [p for p in res["problemas"] if p.get("bloqueante") and (conciliacion or "celdas de las hojas de marca" not in p["titulo"])]


def test_sintetico_reglas(tmp_path):
    res = gd.calcular(construir_libro(tmp_path / "20261109 s.xlsx", FACTURAS_BASE, BO_BASE))
    assert _bloqueantes(res) == []
    assert [m["nombre"] for m in res["marcas"]] == ["Marca X"]   # la hoja oculta no se lee
    c = res["marcas"][0]["bloques"][0]["calc"]
    assert c["total"]["fact"] == 107          # Proveedor "Marca X " se compara con trim
    assert c["total"]["bo"] == 50
    assert [m["fc"] for m in c["meses"][:3]] == [100, 100, 100]   # forecast derivado = trimestre / 3
    assert all(m["fcDerivado"] for m in c["meses"][:3])
    assert c["trims"][0]["cumpl"] == pytest.approx(107 / 300)
    assert [t["cumpl"] for t in c["trims"][1:]] == [None, None, None]
    assert c["ytd"]["nMeses"] == 5 and c["ytd"]["cumpl"] == pytest.approx(157 / 300)
    assert [t["estado"] for t in c["trims"]] == ["cerrado", "parcial", "futuro", "futuro"]


def test_sintetico_folio_texto_encuentra_folio_numerico(tmp_path):
    facturas = [_fact("3", datetime(2026, 3, 5), 7)]      # folio guardado como texto
    res = gd.calcular(construir_libro(tmp_path / "20261109 s.xlsx", facturas, []))
    reg = next(r for r in res["registros_f"])
    assert reg["cat"] == "Repuestos"


def test_sintetico_sin_categoria_es_error_salvo_flag(tmp_path):
    facturas = FACTURAS_BASE + [_fact(2, datetime(2026, 4, 1), 5)]   # folio 2 sin categoría en el diccionario
    ruta = construir_libro(tmp_path / "20261109 s.xlsx", facturas, BO_BASE)
    res = gd.calcular(ruta)
    assert any("sin categoría estándar" in p["titulo"] for p in _bloqueantes(res))
    res = gd.calcular(ruta, permitir_sin_categoria=True)
    assert _bloqueantes(res) == []
    assert any(p["nivel"] == "aviso" and "sin categoría estándar" in p["titulo"] for p in res["problemas"])


def test_sintetico_fila_sin_formula_falla(tmp_path):
    # La factura de la fila 4 de Excel (índice 2) quedó sin la fórmula de "Total moneda marca".
    res = gd.calcular(construir_libro(tmp_path / "20261109 s.xlsx", FACTURAS_BASE, BO_BASE, borrar=("fact", 2, "total")))
    errores = [p for p in _bloqueantes(res) if "no coincide con el recalculado" in p["titulo"]]
    assert errores and "fila 4" in errores[0]["detalle"] and "folio 3" in errores[0]["detalle"]
    assert "falta la fórmula" in errores[0]["detalle"]


def test_sintetico_tipo_de_cambio_como_texto_falla(tmp_path):
    res = gd.calcular(construir_libro(tmp_path / "20261109 s.xlsx", FACTURAS_BASE, BO_BASE, pares={"USD>EUR": "0.9"}))
    errores = [p for p in _bloqueantes(res) if "tipos de cambio no numéricos" in p["titulo"]]
    assert errores and "Parametros Moneda!E2" in errores[0]["detalle"]


def test_sintetico_moneda_sin_tipo_de_cambio_falla(tmp_path):
    facturas = FACTURAS_BASE + [_fact(4, datetime(2026, 4, 1), 50, moneda="CLP")]   # no hay par CLP>EUR
    res = gd.calcular(construir_libro(tmp_path / "20261109 s.xlsx", facturas, BO_BASE))
    assert any("sin tipo de cambio" in p["titulo"] for p in _bloqueantes(res))


def _escribir_hoja_marca(ruta, res, alterar=None):
    """Escribe en la hoja de marca los valores que Excel habría cacheado (alterar=(celda, valor) cambia uno)."""
    b = res["marcas"][0]["bloques"][0]
    vx = b["vista"]
    wb = openpyxl.load_workbook(ruta)
    ws = wb["MarcaX"]
    for f in vx["filas"]:
        for j, c in enumerate(vx["columnas"]):
            if f["clave"] == "cumpl" and c["tipo"] == "mes":
                continue
            v = f["celda"][j] if f["clave"] == "forecast" and c["tipo"] == "mes" else f["valores"][j]
            ws[f'{c["col"]}{f["fila"]}'] = 0 if v is None else v
        if "pct" in f:
            ws[f'{vx["colParticipacion"]}{f["fila"]}'] = f["pct"]
    for r, etiqueta in zip([11, 12, 13, 14, 15, 19, 20, 21, 22, 23], ["EQUIPOS", "REPUESTOS", "ACCESORIOS", "CONSUMIBLES", "Total"] * 2):
        ws[f"A{r}"] = etiqueta
    if alterar:
        ws[alterar[0]] = alterar[1]
    wb.save(ruta)


def test_sintetico_celda_de_hoja_de_marca_que_no_cuadra_falla(tmp_path):
    ruta = construir_libro(tmp_path / "20261109 s.xlsx", FACTURAS_BASE, BO_BASE)
    _escribir_hoja_marca(ruta, gd.calcular(ruta))
    res = gd.calcular(ruta)
    assert res["resumen"]["diferencias"] == 0 and res["resumen"]["comparaciones"] == 253
    # Se altera una celda de facturación por categoría (Repuestos · Mar = D20)
    _escribir_hoja_marca(ruta, res, alterar=("D20", 999))
    res = gd.calcular(ruta)
    assert res["resumen"]["diferencias"] == 1
    error = next(p for p in _bloqueantes(res, conciliacion=True) if "no cuadran" in p["titulo"])
    assert "MarcaX!D20 (Repuestos · Mar)" in error["detalle"]


def test_sintetico_hoja_con_espacio_falla(tmp_path):
    ruta = construir_libro(tmp_path / "20261109 s.xlsx", FACTURAS_BASE, BO_BASE, hoja_fact="BBDD Facturacion ")
    with pytest.raises(gd.ErrorCarga, match="nombre exacto"):
        gd.calcular(ruta)


# ──────────────────────────────────────────────────────────────────────
# Comparación entre versiones
# ──────────────────────────────────────────────────────────────────────

def _comparar(tmp_path, facturas_ant, bo_ant, facturas_act, bo_act):
    ant = gd.calcular(construir_libro(tmp_path / "20261026 a.xlsx", facturas_ant, bo_ant))
    act = gd.calcular(construir_libro(tmp_path / "20261109 b.xlsx", facturas_act, bo_act))
    return gd.comparar_versiones(ant, act)


def test_comparacion_bo_facturado_es_esperable(tmp_path):
    # La línea de BO de la OC 0001/26 se factura: el BO baja y aparece una factura nueva que menciona la OC.
    bo_act = [_bo("0001/26", "P1", datetime(2026, 5, 5), 0, 25)]
    fact_act = FACTURAS_BASE + [_fact(4, datetime(2026, 5, 20), 50, comentarios="INV 4 ; PO 0001/26")]
    cmp = _comparar(tmp_path, FACTURAS_BASE, BO_BASE, fact_act, bo_act)
    mayo = next(f for f in cmp["filas"] if f["marca"] == "Marca X" and (f["anio"], f["mes"]) == (2026, 4))
    assert mayo["estado"] == "esperable"
    assert mayo["bo"]["dif"] == -50 and mayo["fact"]["dif"] == 50
    assert mayo["bo"]["explicado"] and mayo["fact"]["explicado"]
    causas = {(i["tipo"], i["causa"]) for i in mayo["items"]}
    assert ("bo", "BO recibido/facturado (factura nueva que menciona la OC)") in causas
    assert ("fact", "Factura nueva posterior al corte anterior") in causas


def test_comparacion_factura_retroactiva_es_revisar(tmp_path):
    # Folio nuevo con fecha de febrero, anterior al corte de la versión vieja (05-05-2026).
    fact_act = FACTURAS_BASE + [_fact(5, datetime(2026, 2, 15), 80)]
    cmp = _comparar(tmp_path, FACTURAS_BASE, BO_BASE, fact_act, BO_BASE)
    feb = next(f for f in cmp["filas"] if (f["anio"], f["mes"]) == (2026, 1))
    assert feb["estado"] == "revisar"
    assert feb["fact"]["dif"] == 80 and feb["fact"]["explicado"]
    assert feb["items"][0]["causa"] == "Folio nuevo con fecha retroactiva (anterior al corte anterior)"
    assert "Folio 5" in feb["items"][0]["doc"]


def test_comparacion_bo_desaparece_sin_factura_es_revisar(tmp_path):
    bo_act = [_bo("0001/26", "P1", datetime(2026, 5, 5), 0, 25)]
    cmp = _comparar(tmp_path, FACTURAS_BASE, BO_BASE, FACTURAS_BASE, bo_act)
    mayo = next(f for f in cmp["filas"] if (f["anio"], f["mes"]) == (2026, 4))
    assert mayo["estado"] == "revisar"
    assert mayo["items"][0]["causa"] == "BO que desaparece sin factura nueva que mencione la OC"


def test_comparacion_linea_nueva_en_oc_reciente_es_esperable_y_en_oc_antigua_revisar(tmp_path):
    # Corte de BO de la versión vieja: 05-05-2026.
    bo_ant = BO_BASE + [_bo("0002/26", "P1", datetime(2026, 3, 1), 1, 10)]
    bo_act = bo_ant + [
        _bo("0001/26", "P2", datetime(2026, 5, 5), 1, 30),   # flete agregado a una OC de la misma semana
        _bo("0002/26", "P2", datetime(2026, 3, 1), 1, 20),   # línea nueva en una OC de 65 días antes del corte
    ]
    cmp = _comparar(tmp_path, FACTURAS_BASE, bo_ant, FACTURAS_BASE, bo_act)
    mayo = next(f for f in cmp["filas"] if (f["anio"], f["mes"]) == (2026, 4))
    assert mayo["estado"] == "esperable"
    assert mayo["items"][0]["causa"].startswith("Línea nueva o aumento en OC reciente")
    marzo = next(f for f in cmp["filas"] if (f["anio"], f["mes"]) == (2026, 2))
    assert marzo["estado"] == "revisar"
    assert marzo["items"][0]["causa"] == "BO que sube en una OC antigua"


def test_comparacion_probable_factura_sin_referencia_de_oc(tmp_path):
    # El BO de 50 desaparece y llega una única factura nueva de 49,8 que no menciona ninguna OC.
    bo_act = [_bo("0001/26", "P1", datetime(2026, 5, 5), 0, 25)]
    fact_act = FACTURAS_BASE + [_fact(4, datetime(2026, 5, 20), 49.8, comentarios="Fact.proveedores - 0005P")]
    cmp = _comparar(tmp_path, FACTURAS_BASE, BO_BASE, fact_act, bo_act)
    mayo = next(f for f in cmp["filas"] if (f["anio"], f["mes"]) == (2026, 4))
    item = next(i for i in mayo["items"] if i["tipo"] == "bo")
    assert item["causa"] == "Probable factura sin referencia de OC"
    assert item["clase"] == "esperable" and item["aviso"] is True
    assert "probable folio 4" in item["doc"]
    assert mayo["estado"] == "esperable" and mayo["aviso"] is True
    resumen = gd.resumen_cambios_json(cmp)
    assert "folio" not in json.dumps(resumen, ensure_ascii=False).lower()   # el folio no se publica


def test_comparacion_dos_facturas_candidatas_sigue_en_revisar(tmp_path):
    bo_act = [_bo("0001/26", "P1", datetime(2026, 5, 5), 0, 25)]
    fact_act = FACTURAS_BASE + [_fact(4, datetime(2026, 5, 20), 50, comentarios="sin ref"),
                                _fact(5, datetime(2026, 5, 21), 50.3, comentarios="sin ref")]
    cmp = _comparar(tmp_path, FACTURAS_BASE, BO_BASE, fact_act, bo_act)
    mayo = next(f for f in cmp["filas"] if (f["anio"], f["mes"]) == (2026, 4))
    item = next(i for i in mayo["items"] if i["tipo"] == "bo")
    assert item["clase"] == "revisar" and "2 facturas sin OC" in item["doc"]


def test_comparacion_cambio_de_tipo_de_cambio(tmp_path):
    facturas = FACTURAS_BASE + [_fact(4, datetime(2026, 4, 1), 100, moneda="USD")]
    ant = gd.calcular(construir_libro(tmp_path / "20261026 a.xlsx", facturas, BO_BASE, pares={"USD>EUR": 0.9}))
    act = gd.calcular(construir_libro(tmp_path / "20261109 b.xlsx", facturas, BO_BASE, pares={"USD>EUR": 0.95}))
    cmp = gd.comparar_versiones(ant, act)
    abril = next(f for f in cmp["filas"] if (f["anio"], f["mes"]) == (2026, 3))
    assert abril["fact"]["dif"] == pytest.approx(5)
    assert abril["items"][0]["causa"] == "Cambio de tipo de cambio (Parametros Moneda)"
    assert any("USD>EUR" in p for p in cmp["parametros"])


# ──────────────────────────────────────────────────────────────────────
# JSON publicado
# ──────────────────────────────────────────────────────────────────────

def test_json_publicado_es_valido_y_conciliado():
    ruta = RAIZ / gd.SALIDA_JSON
    if not ruta.exists():
        pytest.skip("Aún no se generó docs/datos_cumplimiento.json.")
    d = json.loads(ruta.read_text(encoding="utf-8"))
    assert d["resumen"]["diferencias"] == 0 and d["resumen"]["erroresBloqueantes"] == 0
    assert d["meta"]["fact"]["hoja"] == "BBDD Facturacion"
    assert d["meta"]["fechaArchivo"] == datetime.strptime(d["meta"]["archivo"][:8], "%Y%m%d").strftime("%Y-%m-%d")
    assert ("detalle" in d["marcas"][0]) == d["meta"]["incluirDetalle"]
    if d.get("cambios"):
        # El resumen de cambios no publica OC, productos ni precios
        texto = json.dumps(d["cambios"], ensure_ascii=False)
        assert not re.search(r"\d{1,6}/\d{2}", texto)      # números de OC (p. ej. 1588/26)
        for campo in ('"doc"', '"referencia"', '"producto"', '"precio"', '"folio"', "Folio "):
            assert campo not in texto
