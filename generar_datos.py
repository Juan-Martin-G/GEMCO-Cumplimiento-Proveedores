"""
generar_datos.py
================
Precalcula el cumplimiento de proveedores a partir del Excel consolidado y lo
guarda en docs/datos_cumplimiento.json, que consume el panel estático
(HTML/JS puro en GitHub Pages, sin Python en vivo).

USO (cada vez que se actualiza el Excel):

    python generar_datos.py
    python generar_datos.py --permitir-sin-categoria

El archivo fuente es el .xlsx de data/raw/ (no recursivo) cuyo nombre empieza con
la fecha AAAAMMDD más reciente. La versión anterior, para el reporte de cambios,
es la de fecha más reciente en data/raw/old/ que sea anterior a la actual.

Pasos:
  1. Recalcula en Python cada fila de "BBDD Facturacion" y "BBDD BO abiertas"
     desde los datos crudos (marca, año, mes, categoría, moneda marca, TC y total)
     y lo compara contra el valor cacheado de cada columna con fórmula.
  2. Valida categorías, monedas y tipos de cambio de las marcas con hoja visible.
  3. Calcula el cumplimiento y lo concilia contra las hojas de marca visibles.
  4. Compara contra la versión anterior y escribe reportes/cambios_<ant>_a_<act>.md.

Si algo no cuadra, termina con error y NO escribe el JSON.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from datetime import date, datetime, timedelta
from itertools import repeat
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

# ──────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ──────────────────────────────────────────────────────────────────────

# False → el JSON no incluye las tablas de BO abiertas ni de facturas
# (ni OC, ni productos, ni precios): solo agregados. El panel oculta esas tablas.
INCLUIR_DETALLE = True

CARPETA_EXCEL = Path("data/raw")
NOMBRE_CARPETA_ANTERIORES = "old"           # data/raw/old/ (sin distinguir mayúsculas)
CARPETA_REPORTES = Path("reportes")
SALIDA_JSON = Path("docs/datos_cumplimiento.json")
PATRON_ARCHIVO = re.compile(r"^(\d{8}).*\.xlsx$")

# Nombres EXACTOS (sin trim). Ojo: existe una hoja oculta "BBDD Facturacion "
# (con espacio final), base antigua de 2025, que NO debe leerse nunca.
HOJA_FACTURACION = "BBDD Facturacion"
HOJA_BO = "BBDD BO abiertas"
HOJA_MONEDA = "Parametros Moneda"
HOJA_PROVEEDOR = "Parametros Proveedor"
HOJA_DICCIONARIO = "Diccionario Categorias"

# Columna interna → encabezado en el Excel.
COLUMNAS_FACT = {
    "nombre": "Nombre de cliente/proveedor", "marca": "Proveedor", "fecha": "Fecha de contabilización",
    "anio": "Año", "mes": "Mes", "folio": "Número de folio", "comentarios": "Comentarios",
    "oc": "Orden de compra", "anio_oc": "Año de OC", "cat": "Categoria", "moneda_orig": "Moneda del documento",
    "monto_orig": "Total documento (ME)", "moneda_marca": "Moneda marca", "tc": "TC a moneda marca",
    "total": "Total moneda marca",
}
COLUMNAS_BO = {
    "fecha": "Fecha OC", "anio": "Año OC", "mes": "Mes", "nombre": "Nombre proveedor", "referencia": "Referencia",
    "producto": "Producto", "descripcion": "Descripcion", "cat": "Grupo", "marca": "Marca",
    "cantidad": "Cantidad pdte", "moneda_orig": "moneda", "precio": "Precio", "monto_orig": "Total Neto ME",
    "moneda_marca": "Moneda marca", "tc": "TC a moneda marca", "total": "Total moneda marca",
}
# Columnas que en el Excel son fórmulas: se recalculan y se comparan contra el cache.
FORMULAS = {
    "fact": ["marca", "anio", "mes", "anio_oc", "cat", "moneda_marca", "tc", "total"],
    "bo": ["anio", "mes", "cat", "marca", "monto_orig", "moneda_marca", "tc", "total"],
}
# Compatibilidad: listas de encabezados requeridos
ENCABEZADOS_FACTURACION = list(COLUMNAS_FACT.values())
ENCABEZADOS_BO = list(COLUMNAS_BO.values())

CATEGORIAS = ["Equipos", "Repuestos", "Accesorios", "Consumibles"]
SIN_CATEGORIA = "Sin categoría"

TOLERANCIA_MONTO = 0.01
TOLERANCIA_PCT = 0.0001

# Reporte de cambios
DIAS_OC_ANTIGUA = 30            # una OC es "antigua" si su fecha es > 30 días anterior al corte de BO de la versión vieja
COINCIDENCIA_PCT = 0.01         # factura sin OC que coincide con un BO desaparecido: diferencia < 1 % …
COINCIDENCIA_ABS = 1.0          # … o < 1 unidad de moneda

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
MESES_CORTOS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

# Desplazamientos relativos a la columna de etiquetas ("Item") de cada bloque.
LAYOUT_MESES = [1, 2, 3, 5, 6, 7, 9, 10, 11, 13, 14, 15]
LAYOUT_TRIMESTRES = [4, 8, 12, 16]
LAYOUT_TOTAL = 17
FILA = {"fact": 3, "bo": 4, "factbo": 5, "forecast": 6, "cumpl": 7, "bo_total": 15, "fact_total": 23}


class ErrorCarga(Exception):
    """Error que detiene la carga (archivo, hoja o encabezados faltantes, etc.)."""


# ──────────────────────────────────────────────────────────────────────
# UTILIDADES
# ──────────────────────────────────────────────────────────────────────

NO_NUM = object()  # celda con valor no numérico (texto, #N/A, etc.)


def txt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def clave(v) -> str:
    """Clave normalizada para comparar: trim + minúsculas; 123.0 y "123" → "123"."""
    return txt(v).lower()


def numero(v):
    """None = vacía · NO_NUM = no numérica · float en otro caso."""
    if v is None:
        return None
    if isinstance(v, bool):
        return NO_NUM
    if isinstance(v, (int, float)):
        return NO_NUM if isinstance(v, float) and math.isnan(v) else float(v)
    if isinstance(v, str) and v.strip() == "":
        return None
    return NO_NUM


def es_num(x) -> bool:
    return x is not None and x is not NO_NUM


def es_numerico(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, float) and math.isnan(v))


def a_fecha(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if es_numerico(v):
        try:
            return from_excel(v)
        except (ValueError, OverflowError):
            return None
    return None


def idx_mes(anio: int, mes: int) -> int:
    return anio * 12 + mes


def mes_corto(anio: int, mes: int) -> str:
    return f"{MESES_CORTOS[mes]}-{str(anio)[2:]}"


def col_letra(c0: int) -> str:
    return get_column_letter(c0 + 1)


def ref_hoja(h: str) -> str:
    return h if h.replace("_", "").isalnum() and h.isascii() else "'" + h.replace("'", "''") + "'"


def dir_celda(fila: int, c0: int) -> str:
    return f"{col_letra(c0)}{fila}"


def valor(ws, fila: int, c0: int):
    """fila 1-based, columna 0-based."""
    return ws.cell(row=fila, column=c0 + 1).value


def categoria(v) -> str:
    k = clave(v)
    for c in CATEGORIAS:
        if c.lower() == k:
            return c
    return SIN_CATEGORIA


def todas_categorias():
    return CATEGORIAS + [SIN_CATEGORIA]


def vacio_cat():
    return {c: 0.0 for c in todas_categorias()}


def hoja_oculta(ws) -> bool:
    return ws.sheet_state != "visible"


def fmt_cl(v: float, decimales: int = 2) -> str:
    """Formato es-CL: miles con punto, decimales con coma."""
    return f"{v:,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_monto(v) -> str:
    if not es_numerico(v):
        return "—"
    s = fmt_cl(v, 2)
    return s[:-3] if s.endswith(",00") else s


def _lista(arr, n=8):
    return " · ".join(arr[:n]) + (f" · … ({len(arr)} en total)" if len(arr) > n else "")


# ──────────────────────────────────────────────────────────────────────
# SELECCIÓN DEL ARCHIVO
# ──────────────────────────────────────────────────────────────────────

def _versiones(carpeta: Path) -> list[tuple[date, Path]]:
    """Archivos AAAAMMDD*.xlsx de la carpeta (no recursivo), sin "~$". Error si la fecha no es válida."""
    versiones = []
    for p in sorted(carpeta.iterdir()) if carpeta.is_dir() else []:
        if not p.is_file() or p.name.startswith("~$"):
            continue
        m = PATRON_ARCHIVO.match(p.name)
        if not m:
            continue
        try:
            fecha = datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            raise ErrorCarga(f'La fecha del nombre no es válida: "{p.name}" ({m.group(1)} no es AAAAMMDD).')
        versiones.append((fecha, p))
    return versiones


def _duplicadas(versiones, carpeta):
    vistas: dict = {}
    for f, p in versiones:
        vistas.setdefault(f, []).append(p.name)
    dup = {f: ns for f, ns in vistas.items() if len(ns) > 1}
    if dup:
        detalle = "; ".join(f'{f:%Y%m%d}: {", ".join(ns)}' for f, ns in dup.items())
        raise ErrorCarga(f"Hay dos o más archivos con la misma fecha en {carpeta}/: {detalle}.")


def seleccionar_actual(carpeta: Path = CARPETA_EXCEL) -> dict:
    """El .xlsx de data/raw/ con la fecha (del nombre) más reciente."""
    versiones = _versiones(carpeta)
    if not versiones:
        raise ErrorCarga(f"No hay archivos válidos en {carpeta}/: se espera un .xlsx cuyo nombre empiece con la fecha AAAAMMDD.")
    _duplicadas(versiones, carpeta)
    fecha, ruta = max(versiones)
    avisos = []
    if len(versiones) > 1:
        otros = ", ".join(f'"{p.name}"' for f, p in sorted(versiones) if p != ruta)
        avisos.append(f'Hay {len(versiones)} archivos .xlsx en {carpeta}/: se usó "{ruta.name}" (fecha más reciente). '
                      f"Los demás deberían ir a {carpeta}/{NOMBRE_CARPETA_ANTERIORES}/: {otros}.")
    return {"ruta": ruta, "fecha": fecha, "avisos": avisos}


def carpeta_anteriores(carpeta: Path = CARPETA_EXCEL) -> Path | None:
    if not carpeta.is_dir():
        return None
    for p in carpeta.iterdir():
        if p.is_dir() and p.name.lower() == NOMBRE_CARPETA_ANTERIORES:
            return p
    return None


def seleccionar_anterior(fecha_actual: date, carpeta: Path = CARPETA_EXCEL) -> dict | None:
    """La versión de data/raw/old/ con la fecha más reciente anterior a la actual (None si no hay)."""
    old = carpeta_anteriores(carpeta)
    if old is None:
        return None
    versiones = [(f, p) for f, p in _versiones(old) if f < fecha_actual]
    if not versiones:
        return None
    _duplicadas(versiones, old)
    fecha, ruta = max(versiones)
    return {"ruta": ruta, "fecha": fecha}


# ──────────────────────────────────────────────────────────────────────
# PARÁMETROS (Proveedor, Diccionario, Moneda)
# ──────────────────────────────────────────────────────────────────────

def _hoja(wb, nombre):
    if nombre not in wb.sheetnames:
        raise ErrorCarga(f'No se encontró la hoja "{nombre}" (se busca por nombre exacto).')
    return wb[nombre]


def _tabla_busqueda(ws, col_clave: int, cols_valor: dict, nombre: str, problemas: list, desde: int = 1) -> dict:
    """Como XLOOKUP sobre la columna completa: primera aparición de cada clave normalizada.
    Avisa si una clave repetida tiene valores distintos."""
    tabla, filas_por_clave = {}, {}
    errores_ref = []
    for r in range(desde, ws.max_row + 1):
        k = clave(ws.cell(r, col_clave).value)
        if not k:
            continue
        if k.startswith("#"):
            errores_ref.append(f"{get_column_letter(col_clave)}{r}")
            continue
        vals = {n: ws.cell(r, c).value for n, c in cols_valor.items()}
        if k not in tabla:
            tabla[k] = (r, vals)
            filas_por_clave[k] = [(r, vals)]
        else:
            filas_por_clave[k].append((r, vals))
    difieren = []
    for k, filas in filas_por_clave.items():
        if len(filas) > 1 and len({tuple(clave(v) for v in vals.values()) for _, vals in filas}) > 1:
            difieren.append(f'"{k}" en filas {", ".join(str(r) for r, _ in filas)} ('
                            + " / ".join(", ".join(txt(v) or "(vacío)" for v in vals.values()) for _, vals in filas)
                            + "); se usa la primera")
    if difieren:
        problemas.append({"nivel": "aviso", "titulo": f'{nombre}: {len(difieren)} códigos repetidos con categorías distintas', "detalle": _lista(difieren)})
    if errores_ref:
        problemas.append({"nivel": "aviso", "titulo": f'{nombre}: {len(errores_ref)} celdas con error (#REF!, etc.) en la columna de códigos', "detalle": _lista(errores_ref, 12)})
    return tabla


def leer_parametros(wb, problemas: list) -> dict:
    prov_ws = _hoja(wb, HOJA_PROVEEDOR)
    dic_ws = _hoja(wb, HOJA_DICCIONARIO)
    mon_ws = _hoja(wb, HOJA_MONEDA)

    proveedores = _tabla_busqueda(prov_ws, 1, {"marca": 2}, HOJA_PROVEEDOR, problemas)
    folios = _tabla_busqueda(dic_ws, 2, {"cat": 3, "anio_oc": 4}, f"{HOJA_DICCIONARIO} (folios, columnas B:D)", problemas)
    productos = _tabla_busqueda(dic_ws, 7, {"cat": 8}, f"{HOJA_DICCIONARIO} (productos, columnas G:H)", problemas)

    marcas, pares, tc_invalidos = {}, {}, []
    lista_marcas = []
    for r in range(2, mon_ws.max_row + 1):
        marca, mon = mon_ws.cell(r, 1).value, mon_ws.cell(r, 2).value
        if txt(marca):
            lista_marcas.append({"marca": str(marca), "clave": clave(marca), "moneda": txt(mon), "celda": f"A{r}"})
            marcas.setdefault(clave(marca), txt(mon))
        par, tc = txt(mon_ws.cell(r, 4).value), mon_ws.cell(r, 5).value
        if par:
            if not es_numerico(tc):
                tc_invalidos.append(f'{HOJA_MONEDA}!E{r} ({par}) = {tc!r}'
                                    + (" — está como texto: escríbelo como número, con coma decimal" if isinstance(tc, str) else " — vacío"))
            pares.setdefault(clave(par), {"par": par, "tc": float(tc) if es_numerico(tc) else tc, "celda": f"E{r}"})
    if tc_invalidos:
        problemas.append({"nivel": "error", "bloqueante": True,
                          "titulo": f'{HOJA_MONEDA}: {len(tc_invalidos)} tipos de cambio no numéricos', "detalle": _lista(tc_invalidos)})
    else:
        problemas.append({"nivel": "ok", "titulo": f'Todos los tipos de cambio de "{HOJA_MONEDA}" son numéricos ({len(pares)} pares)'})
    return {"proveedores": proveedores, "folios": folios, "productos": productos,
            "monedas": marcas, "lista_marcas": lista_marcas, "pares": pares}


def _buscar(tabla: dict, k, campo: str, defecto):
    """XLOOKUP: clave vacía o no encontrada → defecto; encontrada con celda vacía → ""."""
    kk = clave(k)
    if not kk or kk not in tabla:
        return defecto
    v = tabla[kk][1][campo]
    return "" if v is None else v


# ──────────────────────────────────────────────────────────────────────
# LECTURA Y RECÁLCULO DE BBDD
# ──────────────────────────────────────────────────────────────────────

def leer_tabla(wb_v, wb_f, nombre: str, columnas: dict, formulas: list[str]) -> dict:
    # Búsqueda por nombre EXACTO: "BBDD Facturacion" ≠ "BBDD Facturacion ".
    if nombre not in wb_v.sheetnames:
        hojas = ", ".join(f'"{n}"' for n in wb_v.sheetnames)
        raise ErrorCarga(f'No se encontró la hoja "{nombre}" (se busca por nombre exacto). Hojas del libro: {hojas}.')
    ws = wb_v[nombre]
    filas_v = ws.iter_rows(min_row=1, min_col=1, values_only=True)
    encabezado = next(filas_v, None)
    if not encabezado:
        raise ErrorCarga(f'La hoja "{nombre}" está vacía.')
    col: dict[str, int] = {}
    for i, h in enumerate(encabezado):
        k = txt(h)
        if k and k not in col:
            col[k] = i
    faltan = [h for h in columnas.values() if h not in col]
    if faltan:
        raise ErrorCarga(f'La hoja "{nombre}" no tiene los encabezados requeridos: '
                         + ", ".join(f'"{h}"' for h in faltan)
                         + ". Se detuvo la carga para no leer una base equivocada.")
    idx = {k: col[h] for k, h in columnas.items()}
    i_formulas = {idx[k] for k in formulas}
    crudas = [i for i in range(len(encabezado)) if i not in i_formulas]
    filas_f = wb_f[nombre].iter_rows(min_row=2, min_col=1, values_only=True) if wb_f is not None else repeat(None)
    datos = []
    for n, (fv, ff) in enumerate(zip(filas_v, filas_f), start=2):
        # Filas sin ningún dato crudo (p. ej. filas con fórmulas arrastradas al final) se ignoran.
        if all(txt(fv[i]) == "" for i in crudas if i < len(fv)):
            continue
        c = {k: (fv[i] if i < len(fv) else None) for k, i in idx.items()}
        f = {k: (ff[idx[k]] if ff is not None and idx[k] < len(ff) else None) for k in formulas} if ff is not None else None
        datos.append({"fila": n, "c": c, "f": f})
    return {"nombre": nombre, "filas": datos, "oculta": hoja_oculta(ws)}


def recalcular_fila(tipo: str, c: dict, p: dict) -> dict:
    """Replica las fórmulas del Excel desde los datos crudos."""
    fecha = a_fecha(c["fecha"])
    marca = _buscar(p["proveedores"], c["nombre"], "marca", "")
    r = {"fecha": fecha, "marca": marca,
         "anio": fecha.year if fecha else None, "mes": MESES[fecha.month - 1] if fecha else None}
    if tipo == "fact":
        r["cat"] = _buscar(p["folios"], c["folio"], "cat", "ERROR")
        r["anio_oc"] = _buscar(p["folios"], c["folio"], "anio_oc", "error")
        monto = c["monto_orig"] if c["monto_orig"] is not None else 0.0
    else:
        r["cat"] = _buscar(p["productos"], c["producto"], "cat", "ERROR")
        precio = c["precio"] if c["precio"] is not None else 0.0
        cant = c["cantidad"] if c["cantidad"] is not None else 0.0
        monto = precio * cant if es_numerico(precio) and es_numerico(cant) else "#VALUE!"
        r["monto_orig"] = monto
    moneda_doc = c["moneda_orig"]
    km = clave(marca)
    r["moneda_marca"] = p["monedas"][km] if km and km in p["monedas"] else ("" if moneda_doc is None else moneda_doc)
    if clave(moneda_doc) == "" or clave(moneda_doc) == clave(r["moneda_marca"]):
        r["tc"] = 1.0
    else:
        par = p["pares"].get(clave(f"{txt(moneda_doc)}>{txt(r['moneda_marca'])}"))
        r["tc"] = "FALTA TC" if par is None else par["tc"]
    r["total"] = monto * r["tc"] if es_numerico(monto) and es_numerico(r["tc"]) else "#VALUE!"
    return r


def _iguales(calc, cache) -> bool:
    calc_vacio, cache_vacio = calc is None or calc == "", cache is None or cache == ""
    if calc_vacio and cache_vacio:
        return True
    if calc_vacio and cache == 0:      # XLOOKUP que encuentra una celda vacía devuelve 0
        return True
    if es_numerico(calc) and es_numerico(cache):
        return abs(calc - cache) <= 1e-6 * max(1.0, abs(calc))
    return clave(calc) == clave(cache)


def _desc_doc(tipo, c) -> str:
    if tipo == "fact":
        return f'folio {txt(c["folio"]) or "(vacío)"}'
    return f'OC {txt(c["referencia"]) or "(vacía)"} · producto {txt(c["producto"]) or "(vacío)"}'


def normalizar(tabla: dict, tipo: str, p: dict, problemas: list) -> dict:
    """Recalcula cada fila, la compara contra el cache y arma los registros (con los valores recalculados)."""
    columnas = COLUMNAS_FACT if tipo == "fact" else COLUMNAS_BO
    registros, corte = [], None
    errores, avisos_cat, fijas, sin_fecha = [], [], [], []
    for f in tabla["filas"]:
        c, n = f["c"], f["fila"]
        calc = recalcular_fila(tipo, c, p)
        doc = _desc_doc(tipo, c)
        if calc["fecha"] is None:
            sin_fecha.append(f'fila {n} ({doc})')
            continue
        difs = []
        for k in FORMULAS[tipo]:
            formula = f["f"][k] if f["f"] is not None else None
            tiene_formula = isinstance(formula, str) and formula.startswith("=")
            if not _iguales(calc[k], c[k]):
                vacia = c[k] is None or c[k] == ""
                motivo = "celda vacía (falta la fórmula)" if vacia else ("valor fijo sin fórmula" if f["f"] is not None and not tiene_formula else "valor distinto")
                difs.append((k, motivo))
            elif f["f"] is not None and not tiene_formula and c[k] is not None:
                fijas.append(f'fila {n} col "{columnas[k]}"')
        if difs:
            monto_cero = es_numerico(calc["total"]) and calc["total"] == 0 and (tipo == "fact" or not c["cantidad"])
            detalle = "; ".join(f'"{columnas[k]}" Excel={c[k]!r} recalculado={calc[k]!r} ({m})' for k, m in difs)
            texto = f'{tabla["nombre"]} fila {n} · {doc} · marca "{txt(calc["marca"])}": {detalle}'
            if [k for k, _ in difs] == ["cat"] and monto_cero:
                avisos_cat.append(texto)
            else:
                errores.append(texto)
        if not corte or calc["fecha"] > corte:
            corte = calc["fecha"]
        fecha = calc["fecha"]
        monto_orig = numero(calc["monto_orig"] if tipo == "bo" else c["monto_orig"])
        reg = {
            "tipo": tipo, "fila": n, "fecha": fecha, "anio": fecha.year, "mes": fecha.month - 1,
            "marca": txt(calc["marca"]), "marca_clave": clave(calc["marca"]),
            "moneda_orig": txt(c["moneda_orig"]),
            "monto_orig": monto_orig if es_num(monto_orig) else 0.0,
            "moneda_marca": txt(calc["moneda_marca"]),
            "tc": calc["tc"] if es_numerico(calc["tc"]) else None,
            "tc_texto": None if es_numerico(calc["tc"]) else txt(calc["tc"]),
            "total": calc["total"] if es_numerico(calc["total"]) else 0.0,
            "cat_raw": txt(calc["cat"]),
            "cat": categoria(calc["cat"]),
        }
        if tipo == "fact":
            reg["folio"] = txt(c["folio"])
            reg["comentarios"] = txt(c["comentarios"])
            reg["oc"] = txt(c["oc"])
        else:
            cant, precio = numero(c["cantidad"]), numero(c["precio"])
            reg["referencia"] = txt(c["referencia"])
            reg["producto"] = txt(c["producto"])
            reg["descripcion"] = txt(c["descripcion"])
            reg["cantidad"] = cant if es_num(cant) else 0.0
            reg["precio"] = precio if es_num(precio) else 0.0
        reg["convertido"] = bool(reg["moneda_orig"]) and clave(reg["moneda_orig"]) != clave(reg["moneda_marca"])
        registros.append(reg)

    nombre = tabla["nombre"]
    if sin_fecha:
        problemas.append({"nivel": "error", "bloqueante": True, "titulo": f'{nombre}: {len(sin_fecha)} filas con datos sin fecha válida', "detalle": _lista(sin_fecha)})
    if errores:
        problemas.append({"nivel": "error", "bloqueante": True,
                          "titulo": f'{nombre}: {len(errores)} filas donde el valor de una fórmula no coincide con el recalculado (fórmula vacía o desactualizada)',
                          "detalle": _lista(errores, 30), "filas": errores})
    else:
        problemas.append({"nivel": "ok", "titulo": f'{nombre}: las {len(tabla["filas"])} filas con datos coinciden con el recálculo en Python (marca, año, mes, categoría, moneda, TC y total)'
                          + (f", salvo {len(avisos_cat)} con solo la categoría distinta y monto 0" if avisos_cat else "")})
    if avisos_cat:
        problemas.append({"nivel": "aviso", "titulo": f'{nombre}: {len(avisos_cat)} filas con cantidad y monto 0 donde solo difiere la categoría (valor de Excel desactualizado)',
                          "detalle": _lista(avisos_cat, 30), "filas": avisos_cat})
    if fijas:
        problemas.append({"nivel": "aviso", "titulo": f'{nombre}: {len(fijas)} celdas con valor fijo en vez de fórmula (hoy coinciden con el recálculo)', "detalle": _lista(fijas, 12)})
    return {"registros": registros, "corte": corte, "errores_fila": errores}


def indexar(registros: list[dict]) -> dict:
    """(marca, año, mes) → {total, cat, filas}"""
    idx: dict = {}
    for r in registros:
        k = (r["marca_clave"], r["anio"], r["mes"])
        e = idx.get(k)
        if e is None:
            e = idx[k] = {"total": 0.0, "cat": vacio_cat(), "filas": []}
        e["total"] += r["total"]
        e["cat"][r["cat"]] += r["total"]
        e["filas"].append(r)
    return idx


# ──────────────────────────────────────────────────────────────────────
# HOJAS DE MARCA
# ──────────────────────────────────────────────────────────────────────

def leer_bloque(ws, hoja: str, c0: int, tipo: str) -> dict:
    errores = []
    meses = []
    for off in LAYOUT_MESES:
        c = c0 + off
        etiqueta = txt(valor(ws, 2, c))
        mes = MESES.index(etiqueta.lower()) if etiqueta.lower() in MESES else -1
        anio = numero(valor(ws, 1, c))
        if mes < 0:
            errores.append(f'{hoja}!{dir_celda(2, c)}: se esperaba un mes y dice "{etiqueta}".')
        if not (es_num(anio) and anio > 0):
            errores.append(f"{hoja}!{dir_celda(1, c)}: falta el año del mes.")
            anio = None
        meses.append({"col": c, "mes": mes, "anio": int(anio) if anio else None})
    for q, off in enumerate(LAYOUT_TRIMESTRES):
        e = clave(valor(ws, 2, c0 + off))
        if e != f"q{q + 1}":
            errores.append(f'{hoja}!{dir_celda(2, c0 + off)}: se esperaba "Q{q + 1}" y dice "{e}".')
    if clave(valor(ws, 2, c0 + LAYOUT_TOTAL)) != "total":
        errores.append(f'{hoja}!{dir_celda(2, c0 + LAYOUT_TOTAL)}: se esperaba "Total".')
    if not errores:
        for i in range(1, 12):
            if idx_mes(meses[i]["anio"], meses[i]["mes"]) != idx_mes(meses[i - 1]["anio"], meses[i - 1]["mes"]) + 1:
                errores.append(f"{hoja}: los meses del bloque no son consecutivos ({dir_celda(2, meses[i]['col'])}).")
                break

    # Forecast (fila 6): celdas de mes, trimestre y total tal como están en la hoja.
    def fc_celda(off):
        n = numero(valor(ws, FILA["forecast"], c0 + off))
        return n if es_num(n) else None

    fc_mes = [fc_celda(off) for off in LAYOUT_MESES]
    fc_mes_celda = list(fc_mes)          # valor de la celda, antes de derivar trimestre/3
    fc_trim = [fc_celda(off) for off in LAYOUT_TRIMESTRES]
    fc_total = fc_celda(LAYOUT_TOTAL)
    fc_derivado = [False] * 12
    for t, q in enumerate(fc_trim):
        ids = [3 * t, 3 * t + 1, 3 * t + 2]
        if q and q > 0 and all(not fc_mes[i] for i in ids):
            for i in ids:
                fc_mes[i] = q / 3
                fc_derivado[i] = True

    ok = not errores
    ini, fin = meses[0], meses[11]
    if tipo == "calendario":
        titulo = f"Año calendario {ini['anio']}" if ok else "Año calendario"
    else:
        titulo = (f"Año fiscal USA ({mes_corto(ini['anio'], ini['mes'])} – {mes_corto(fin['anio'], fin['mes'])})"
                  if ok else "Año fiscal USA")
    return {"tipo": tipo, "hoja": hoja, "c0": c0, "titulo": titulo, "meses": meses, "errores": errores,
            "fc_mes": fc_mes, "fc_mes_celda": fc_mes_celda, "fc_trim": fc_trim, "fc_total": fc_total,
            "fc_derivado": fc_derivado, "ws": ws}


def calcular_bloque(b: dict, marca_clave: str, idx_f: dict, idx_b: dict, corte_idx: int) -> dict:
    vacio = {"total": 0.0, "cat": vacio_cat(), "filas": []}
    meses = []
    for i, m in enumerate(b["meses"]):
        F = idx_f.get((marca_clave, m["anio"], m["mes"]), vacio)
        B = idx_b.get((marca_clave, m["anio"], m["mes"]), vacio)
        meses.append({
            "anio": m["anio"], "mes": m["mes"], "etiqueta": mes_corto(m["anio"], m["mes"]),
            "fact": F["total"], "bo": B["total"], "factbo": F["total"] + B["total"],
            "fc": b["fc_mes"][i], "fcDerivado": b["fc_derivado"][i],
            "_catF": F["cat"], "_catB": B["cat"], "_filasF": F["filas"], "_filasB": B["filas"],
        })

    def suma(arr, k):
        return sum((x[k] or 0) for x in arr)

    trims = []
    for q in range(4):
        ms = meses[3 * q:3 * q + 3]
        fc = b["fc_trim"][q]
        factbo = suma(ms, "factbo")
        primero, ultimo = idx_mes(ms[0]["anio"], ms[0]["mes"]), idx_mes(ms[2]["anio"], ms[2]["mes"])
        estado = "cerrado" if corte_idx >= ultimo else "parcial" if corte_idx >= primero else "futuro"
        trims.append({"nombre": f"Q{q + 1}", "fact": suma(ms, "fact"), "bo": suma(ms, "bo"), "factbo": factbo,
                      "fc": fc, "cumpl": factbo / fc if fc and fc > 0 else None, "estado": estado})

    total = {"fact": suma(trims, "fact"), "bo": suma(trims, "bo"), "factbo": suma(trims, "factbo"), "fc": b["fc_total"]}
    total["cumpl"] = total["factbo"] / total["fc"] if total["fc"] and total["fc"] > 0 else None

    ms_ytd = [m for m in meses if idx_mes(m["anio"], m["mes"]) <= corte_idx]
    ytd = {"fact": suma(ms_ytd, "fact"), "bo": suma(ms_ytd, "bo"), "factbo": suma(ms_ytd, "factbo"),
           "fc": suma(ms_ytd, "fc"), "nMeses": len(ms_ytd),
           "desde": {"anio": meses[0]["anio"], "mes": meses[0]["mes"]},
           "hasta": {"anio": ms_ytd[-1]["anio"], "mes": ms_ytd[-1]["mes"]} if ms_ytd else None,
           "completo": len(ms_ytd) == 12}
    ytd["cumpl"] = ytd["factbo"] / ytd["fc"] if ytd["fc"] > 0 else None

    cat_f, cat_b, raw_f, raw_b = vacio_cat(), vacio_cat(), {}, {}
    for m in meses:
        for c in todas_categorias():
            cat_f[c] += m["_catF"][c]
            cat_b[c] += m["_catB"][c]
        for r in m["_filasF"]:
            if r["cat"] == SIN_CATEGORIA:
                k = r["cat_raw"] or "(vacío)"
                raw_f[k] = raw_f.get(k, 0.0) + r["total"]
        for r in m["_filasB"]:
            if r["cat"] == SIN_CATEGORIA:
                k = r["cat_raw"] or "(vacío)"
                raw_b[k] = raw_b.get(k, 0.0) + r["total"]

    return {
        "meses": meses, "trims": trims, "total": total, "ytd": ytd,
        "catF": cat_f, "catB": cat_b, "rawF": raw_f, "rawB": raw_b,
        "filasF": [r for m in meses for r in m["_filasF"]],
        "filasB": [r for m in meses for r in m["_filasB"]],
        "fcSumaMeses": suma(meses, "fc"),
        "tieneForecast": bool(total["fc"] and total["fc"] > 0),
        "hayDerivado": any(b["fc_derivado"]),
    }


def conciliar_bloque(b: dict, calc: dict) -> list[dict]:
    ws, c0, hoja = b["ws"], b["c0"], b["hoja"]
    items = []

    def add(concepto, fila, off, app, tipo, nota=""):
        ex = numero(valor(ws, fila, c0 + off))
        no_num = ex is NO_NUM
        ex_v = 0.0 if ex is None else (None if no_num else ex)  # vacío ≙ 0 (IFERROR del Excel)
        ap_v = 0.0 if app is None else app
        dif = None if ex_v is None else ap_v - ex_v
        tol = TOLERANCIA_PCT if tipo == "pct" else TOLERANCIA_MONTO
        ok = dif is not None and abs(dif) <= tol + 1e-9
        items.append({"concepto": concepto, "celda": f"{ref_hoja(hoja)}!{dir_celda(fila, c0 + off)}",
                      "excel": None if (ex is None or no_num) else ex, "excelVacia": ex is None,
                      "excelNoNumerico": no_num, "app": app, "dif": dif, "ok": ok, "tipo": tipo, "nota": nota})

    def sum4(cat):
        return sum(cat[c] for c in CATEGORIAS)

    T = LAYOUT_TOTAL
    add("Facturación total (fila 3)", FILA["fact"], T, calc["total"]["fact"], "monto")
    add("Backorder total (fila 4)", FILA["bo"], T, calc["total"]["bo"], "monto")
    add("Fact + BO total (fila 5)", FILA["factbo"], T, calc["total"]["factbo"], "monto")
    add("Forecast total (fila 6)", FILA["forecast"], T, calc["fcSumaMeses"], "monto",
        "Σ forecast mensual usado (incluye meses derivados = trimestre/3)" if calc["hayDerivado"] else "Σ forecast mensual usado")
    add("Cumplimiento total (fila 7)", FILA["cumpl"], T, calc["total"]["cumpl"], "pct",
        'Sin forecast: el panel muestra "—"; el Excel muestra 0 por IFERROR' if calc["total"]["cumpl"] is None else "")
    for q, t in enumerate(calc["trims"]):
        add(f'Cumplimiento {t["nombre"]} (fila 7)', FILA["cumpl"], LAYOUT_TRIMESTRES[q], t["cumpl"], "pct",
            "Sin forecast en el trimestre (Excel: 0 por IFERROR)" if t["cumpl"] is None else "")
    sin_b, sin_f = calc["catB"][SIN_CATEGORIA], calc["catF"][SIN_CATEGORIA]
    add("Total BO por categoría (fila 15)", FILA["bo_total"], T, sum4(calc["catB"]), "monto",
        f'Excluye "{SIN_CATEGORIA}" ({sin_b}), igual que la hoja' if sin_b else "")
    add("Total facturación por categoría (fila 23)", FILA["fact_total"], T, sum4(calc["catF"]), "monto",
        f'Excluye "{SIN_CATEGORIA}" ({sin_f}), igual que la hoja' if sin_f else "")
    return items


# ──────────────────────────────────────────────────────────────────────
# VISTA EXCEL (réplica de la hoja de marca) Y CONCILIACIÓN CELDA A CELDA
# ──────────────────────────────────────────────────────────────────────

FILAS_CATEGORIA = {"bo": [11, 12, 13, 14], "bo_total": 15, "fact": [19, 20, 21, 22], "fact_total": 23}
COL_PARTICIPACION = 18          # columna S (relativa a la columna de etiquetas)


def _columnas_vista(b: dict) -> list[dict]:
    cols = []
    for off in range(1, LAYOUT_TOTAL + 1):
        if off == LAYOUT_TOTAL:
            cols.append({"off": off, "tipo": "total", "etiqueta": "Total"})
        elif off in LAYOUT_TRIMESTRES:
            cols.append({"off": off, "tipo": "trim", "q": LAYOUT_TRIMESTRES.index(off), "etiqueta": f"Q{LAYOUT_TRIMESTRES.index(off) + 1}"})
        else:
            i = LAYOUT_MESES.index(off)
            m = b["meses"][i]
            cols.append({"off": off, "tipo": "mes", "i": i, "anio": m["anio"], "mes": m["mes"], "etiqueta": MESES_CORTOS[m["mes"]]})
    return cols


def _expandir(mensual: list, cols: list[dict]) -> list:
    """12 valores mensuales → 17 columnas (meses, Q1–Q4 = suma de sus meses, Total = suma de trimestres)."""
    trims = [sum((mensual[3 * q + k] or 0.0) for k in range(3)) for q in range(4)]
    salida = []
    for c in cols:
        salida.append(mensual[c["i"]] if c["tipo"] == "mes" else trims[c["q"]] if c["tipo"] == "trim" else sum(trims))
    return salida


def vista_excel(b: dict, calc: dict) -> tuple[dict, list[str]]:
    """Valores de las filas 3–7, 11–15 y 19–23 (columnas B–R y S) tal como los muestra la hoja de marca."""
    cols = _columnas_vista(b)
    meses = calc["meses"]
    filas = []

    def fila(clave, etiqueta, seccion, nfila, tipo, valores, **extra):
        filas.append({"clave": clave, "etiqueta": etiqueta, "seccion": seccion, "fila": nfila, "tipo": tipo, "valores": valores, **extra})

    fila("fact", "Facturación", "principal", FILA["fact"], "monto", _expandir([m["fact"] for m in meses], cols))
    fila("bo", "Backorder", "principal", FILA["bo"], "monto", _expandir([m["bo"] for m in meses], cols))
    fila("factbo", "Fact + BO", "principal", FILA["factbo"], "monto", _expandir([m["factbo"] for m in meses], cols))
    fila("forecast", "Forecast", "principal", FILA["forecast"], "monto", _expandir(list(b["fc_mes"]), cols),
         derivado=[b["fc_derivado"][c["i"]] if c["tipo"] == "mes" else False for c in cols],
         celda=[b["fc_mes_celda"][c["i"]] if c["tipo"] == "mes" else None for c in cols])
    cumpl = [calc["trims"][c["q"]]["cumpl"] if c["tipo"] == "trim" else calc["total"]["cumpl"] if c["tipo"] == "total" else None for c in cols]
    fila("cumpl", "Cumplimiento", "principal", FILA["cumpl"], "pct", cumpl)

    errores = []
    for seccion, base_total, clave_cat, filas_cat, fila_total in (
            ("bo", calc["total"]["bo"], "_catB", FILAS_CATEGORIA["bo"], FILAS_CATEGORIA["bo_total"]),
            ("fact", calc["total"]["fact"], "_catF", FILAS_CATEGORIA["fact"], FILAS_CATEGORIA["fact_total"])):
        suma = [0.0] * len(cols)
        suma_pct = 0.0
        for cat, nfila in zip(CATEGORIAS, filas_cat):
            etiqueta_hoja = clave(valor(b["ws"], nfila, b["c0"]))
            if etiqueta_hoja != cat.lower():
                errores.append(f'{b["hoja"]}!{dir_celda(nfila, b["c0"])}: se esperaba "{cat.upper()}" y dice "{etiqueta_hoja}".')
            vals = _expandir([m[clave_cat][cat] for m in meses], cols)
            pct = vals[-1] / base_total if base_total else 0.0          # IFERROR(R/R4,0) o IFERROR(R/R3,0)
            fila(f"{seccion}_{cat.lower()}", cat, seccion, nfila, "monto", vals, pct=pct)
            suma = [a + v for a, v in zip(suma, vals)]
            suma_pct += pct
        if clave(valor(b["ws"], fila_total, b["c0"])) != "total":
            errores.append(f'{b["hoja"]}!{dir_celda(fila_total, b["c0"])}: se esperaba "Total".')
        fila(f"{seccion}_total", "Total", seccion, fila_total, "monto", suma, pct=suma_pct)

    vista = {"columnas": [{k: v for k, v in c.items() if k != "off"} | {"col": col_letra(b["c0"] + c["off"])} for c in cols],
             "colParticipacion": col_letra(b["c0"] + COL_PARTICIPACION), "filas": filas}
    return vista, errores


def conciliar_celdas(b: dict, vista: dict) -> list[dict]:
    """Compara cada celda de la vista contra el valor cacheado de la misma celda en la hoja de marca."""
    ws, c0, hoja = b["ws"], b["c0"], b["hoja"]
    items = []

    def comparar(nfila, off, app, tipo, concepto):
        ex = numero(valor(ws, nfila, c0 + off))
        no_num = ex is NO_NUM
        ex_v = 0.0 if ex is None else (None if no_num else ex)          # vacío ≙ 0 (IFERROR del Excel)
        ap_v = 0.0 if app is None else app
        dif = None if ex_v is None else ap_v - ex_v
        tol = TOLERANCIA_PCT if tipo == "pct" else TOLERANCIA_MONTO
        items.append({"celda": f"{ref_hoja(hoja)}!{dir_celda(nfila, c0 + off)}", "concepto": concepto,
                      "excel": None if (ex is None or no_num) else ex, "excelNoNumerico": no_num,
                      "app": app, "dif": dif, "tipo": tipo, "ok": dif is not None and abs(dif) <= tol + 1e-9})

    cols = _columnas_vista(b)
    for f in vista["filas"]:
        for j, c in enumerate(cols):
            if f["clave"] == "cumpl" and c["tipo"] == "mes":
                continue                                    # la fila 7 solo tiene trimestres y total
            app = f["valores"][j]
            if f["clave"] == "forecast" and c["tipo"] == "mes":
                app = f["celda"][j]                         # el mes derivado (trimestre/3) es una celda vacía en el Excel
            comparar(f["fila"], c["off"], app, f["tipo"], f'{f["etiqueta"]} · {c["etiqueta"]}')
        if "pct" in f:
            comparar(f["fila"], COL_PARTICIPACION, f["pct"], "pct", f'{f["etiqueta"]} · % participación')
    return items


def resumen_conversion(calc: dict) -> list[dict]:
    """Monto original, TC y monto convertido por moneda original (recuadro "Conversión de moneda")."""
    salida = []
    for base, filas in ((HOJA_FACTURACION, calc["filasF"]),
                        (HOJA_BO, [r for r in calc["filasB"] if r["cantidad"] > 0 or r["total"]])):
        g: dict = {}
        for r in filas:
            k = r["moneda_orig"] or "(vacía)"
            e = g.get(k)
            if e is None:
                e = g[k] = {"base": base, "moneda": k, "n": 0, "orig": 0.0, "conv": 0.0, "tcs": [], "convertido": r["convertido"]}
            e["n"] += 1
            e["orig"] += r["monto_orig"]
            e["conv"] += r["total"]
            if r["tc"] not in e["tcs"]:
                e["tcs"].append(r["tc"])
        salida.extend(g.values())
    return salida


# ──────────────────────────────────────────────────────────────────────
# VALIDACIONES DE MARCAS VISIBLES
# ──────────────────────────────────────────────────────────────────────

def _desc_reg(r) -> str:
    if r["tipo"] == "fact":
        doc = f'folio {r["folio"]}' + (f' (OC {", ".join(ocs_mencionadas(r))})' if ocs_mencionadas(r) else "")
        extra = ""
    else:
        doc = f'OC {r["referencia"] or "(vacía)"}'
        extra = f' · producto {r["producto"] or "(vacío)"} · "{r["descripcion"]}"'
    hoja = HOJA_FACTURACION if r["tipo"] == "fact" else HOJA_BO
    return (f'{hoja} fila {r["fila"]} · {r["marca"]} · {r["fecha"]:%Y-%m-%d} · {doc}{extra} · '
            f'{fmt_monto(r["monto_orig"])} {r["moneda_orig"] or "(sin moneda)"} → {fmt_monto(r["total"])} {r["moneda_marca"]}')


def validar_marcas(marcas: list[dict], registros: list[dict], params: dict, problemas: list, permitir_sin_categoria: bool):
    por_marca: dict = {}
    for r in registros:
        por_marca.setdefault(r["marca_clave"], []).append(r)

    sin_cat, monedas_mal = [], []
    for m in marcas:
        regs = por_marca.get(m["clave"], [])
        periodos = {(x["anio"], x["mes"]) for b in m["bloques"] for x in b["meses"]}
        for r in regs:
            if r["cat"] != SIN_CATEGORIA:
                continue
            if (r["tipo"] == "fact" and (r["anio"], r["mes"]) in periodos) or (r["tipo"] == "bo" and r["cantidad"] > 0):
                sin_cat.append(f'{_desc_reg(r)} · categoría "{r["cat_raw"] or "(vacía)"}"')
        # Moneda de reporte: la de Parametros Moneda o, si la marca no está, la de la mayoría de sus documentos.
        en_param = m["clave"] in params["monedas"]
        if en_param:
            reporte = params["monedas"][m["clave"]]
        else:
            cuenta: dict = {}
            for r in regs:
                if r["moneda_marca"]:
                    cuenta[r["moneda_marca"]] = cuenta.get(r["moneda_marca"], 0) + 1
            reporte = max(cuenta, key=cuenta.get) if cuenta else ""
        m["moneda_reporte"] = reporte
        for r in regs:
            if r["tc"] is None:
                monedas_mal.append(f'{_desc_reg(r)} · falta el TC "{r["moneda_orig"]}>{r["moneda_marca"]}" en {HOJA_MONEDA}')
            elif not en_param and r["moneda_marca"] and r["moneda_marca"] != reporte:
                monedas_mal.append(f'{_desc_reg(r)} · {m["nombre"]} se reporta en {reporte} y no está en {HOJA_MONEDA}: '
                                   f'agrega la marca (A:B) y el par "{r["moneda_orig"]}>{reporte}" (D:E)')
    if sin_cat:
        problemas.append({"nivel": "aviso" if permitir_sin_categoria else "error", "bloqueante": not permitir_sin_categoria,
                          "titulo": f"{len(sin_cat)} documentos de marcas visibles sin categoría estándar (facturación del período y BO con cantidad pendiente)"
                          + (" — se continúa por --permitir-sin-categoria" if permitir_sin_categoria else " — usa --permitir-sin-categoria para continuar"),
                          "detalle": _lista(sin_cat, 40), "filas": sin_cat})
    else:
        problemas.append({"nivel": "ok", "titulo": "Categorías: toda la facturación del período y el BO pendiente de las marcas visibles es Equipos, Repuestos, Accesorios o Consumibles"})
    if monedas_mal:
        problemas.append({"nivel": "error", "bloqueante": True,
                          "titulo": f"{len(monedas_mal)} documentos de marcas visibles fuera de la moneda de reporte y sin tipo de cambio",
                          "detalle": _lista(monedas_mal, 30), "filas": monedas_mal})
    else:
        problemas.append({"nivel": "ok", "titulo": "Monedas: todos los documentos de las marcas visibles están en la moneda de reporte de su marca o tienen tipo de cambio"})


# ──────────────────────────────────────────────────────────────────────
# PROCESO DE UN LIBRO
# ──────────────────────────────────────────────────────────────────────

def fecha_desde_nombre(ruta: Path):
    m = PATRON_ARCHIVO.match(ruta.name)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d").date()
    except ValueError:
        return None


def calcular(ruta: Path | str, permitir_sin_categoria: bool = False) -> dict:
    """Lee el Excel, recalcula, valida y concilia. Devuelve el resultado interno (con registros)."""
    ruta = Path(ruta)
    wb = openpyxl.load_workbook(ruta, data_only=True)
    wb_f = openpyxl.load_workbook(ruta, data_only=False)
    problemas = []

    t_f = leer_tabla(wb, wb_f, HOJA_FACTURACION, COLUMNAS_FACT, FORMULAS["fact"])
    t_b = leer_tabla(wb, wb_f, HOJA_BO, COLUMNAS_BO, FORMULAS["bo"])
    problemas.append({"nivel": "ok", "titulo": f'Se leyó la hoja "{t_f["nombre"]}" (nombre exacto) · {len(t_f["filas"])} filas con datos · encabezados validados'})
    problemas.append({"nivel": "ok", "titulo": f'Se leyó la hoja "{t_b["nombre"]}" (nombre exacto) · {len(t_b["filas"])} filas con datos · encabezados validados'})
    for n in wb.sheetnames:
        if n != t_f["nombre"] and n.strip() == t_f["nombre"].strip():
            oculta = " (oculta)" if hoja_oculta(wb[n]) else ""
            problemas.append({"nivel": "info", "titulo": f'Se ignoró la hoja "{n.replace(" ", "␣")}"{oculta}: nombre casi igual a "{t_f["nombre"]}", es la base antigua y no se lee.'})
    if t_f["oculta"]:
        problemas.append({"nivel": "aviso", "titulo": f'La hoja "{t_f["nombre"]}" está oculta; se esperaba visible.'})

    params = leer_parametros(wb, problemas)
    n_f = normalizar(t_f, "fact", params, problemas)
    n_b = normalizar(t_b, "bo", params, problemas)
    idx_f, idx_b = indexar(n_f["registros"]), indexar(n_b["registros"])
    cortes = [c for c in (n_f["corte"], n_b["corte"]) if c]
    corte = max(cortes) if cortes else None
    corte_idx = idx_mes(corte.year, corte.month - 1) if corte else 10 ** 9

    ocultas, marcas = [], []
    todos = n_f["registros"] + n_b["registros"]
    for nombre in wb.sheetnames:
        ws = wb[nombre]
        if hoja_oculta(ws):
            ocultas.append(nombre)
            continue
        if clave(valor(ws, 2, 0)) != "item" or clave(valor(ws, 6, 0)) != "forecast":
            continue
        nombre_marca = txt(valor(ws, 1, 0))
        mc = clave(nombre_marca)
        bloques = [leer_bloque(ws, nombre, 0, "calendario")]
        for c in range(19, ws.max_column):
            if clave(valor(ws, 2, c)) == "item":
                bloques.append(leer_bloque(ws, nombre, c, "fiscal"))
                break
        for b in bloques:
            if b["errores"]:
                problemas.append({"nivel": "error", "bloqueante": True, "titulo": f'Hoja "{nombre}", bloque {b["tipo"]}: estructura inesperada', "detalle": " ".join(b["errores"])})
                b["calc"], b["conc"], b["conversion"], b["vista"], b["celdas"] = None, [], [], None, []
                continue
            b["calc"] = calcular_bloque(b, mc, idx_f, idx_b, corte_idx)
            b["conc"] = conciliar_bloque(b, b["calc"])
            b["conversion"] = resumen_conversion(b["calc"])
            b["vista"], errores_vista = vista_excel(b, b["calc"])
            if errores_vista:
                problemas.append({"nivel": "error", "bloqueante": True, "titulo": f'Hoja "{nombre}", bloque {b["tipo"]}: etiquetas de categoría inesperadas', "detalle": " ".join(errores_vista)})
            b["celdas"] = conciliar_celdas(b, b["vista"])
        regs = [r for r in todos if r["marca_clave"] == mc]
        convertidos = [r for r in regs if r["convertido"] and r["tc"] is not None]
        pares_conv = []
        for r in convertidos:
            par = {"orig": r["moneda_orig"], "dest": r["moneda_marca"], "tc": r["tc"]}
            if par not in pares_conv:
                pares_conv.append(par)
        marcas.append({"hoja": nombre, "nombre": nombre_marca, "clave": mc, "bloques": bloques,
                       "tieneFiscal": len(bloques) > 1, "convierte": bool(convertidos),
                       "paresConversion": pares_conv, "sinFilas": not regs, "registros": regs})
    if not marcas:
        raise ErrorCarga('No se encontraron hojas de marca visibles (A2 = "Item" y A6 = "Forecast").')

    validar_marcas(marcas, todos, params, problemas, permitir_sin_categoria)
    for m in marcas:
        m["moneda"] = m["moneda_reporte"] or "—"
        if m["sinFilas"]:
            problemas.append({"nivel": "info", "titulo": f'{m["nombre"]}: no tiene filas en ninguna BBDD'})
    if ocultas:
        problemas.append({"nivel": "info", "titulo": "Hojas ocultas ignoradas (no se leen): " + ", ".join(f'"{n}"' for n in ocultas)})

    # Conciliación celda a celda: todas las celdas de la vista Excel (incluye las de b["conc"]).
    items = [i for m in marcas for b in m["bloques"] for i in b["celdas"]]
    fallas = [i for i in items if not i["ok"]] + [i for m in marcas for b in m["bloques"] for i in b["conc"] if not i["ok"]]
    if fallas:
        def _v(x, tipo):
            return "no numérica" if x is None else (fmt_cl(x * 100, 2) + "%" if tipo == "pct" else fmt_monto(x))
        lineas = [f'{i["celda"]} ({i["concepto"]}): Excel={_v(i["excel"] if not i.get("excelNoNumerico") else None, i["tipo"])} '
                  f'recalculado={_v(i["app"] if i["app"] is not None else 0.0, i["tipo"])}' for i in fallas]
        problemas.append({"nivel": "error", "bloqueante": True,
                          "titulo": f"{len(fallas)} celdas de las hojas de marca no cuadran con el recálculo",
                          "detalle": _lista(lineas, 30), "filas": lineas})
    resumen = {
        "comparaciones": len(items),
        "ok": sum(1 for i in items if i["ok"]),
        "diferencias": len(fallas),
        "errores": sum(1 for p in problemas if p["nivel"] == "error"),
        "erroresBloqueantes": sum(1 for p in problemas if p.get("bloqueante")),
        "avisos": sum(1 for p in problemas if p["nivel"] == "aviso"),
    }
    return {
        "ruta": ruta, "fecha_archivo": fecha_desde_nombre(ruta),
        "t_f": t_f, "t_b": t_b, "corte_f": n_f["corte"], "corte_b": n_b["corte"], "corte": corte,
        "params": params, "registros_f": n_f["registros"], "registros_b": n_b["registros"],
        "idx_f": idx_f, "idx_b": idx_b, "marcas": marcas, "problemas": problemas, "ocultas": ocultas,
        "resumen": resumen,
    }


# ──────────────────────────────────────────────────────────────────────
# COMPARACIÓN CONTRA LA VERSIÓN ANTERIOR
# ──────────────────────────────────────────────────────────────────────

_PATRON_OC = re.compile(r"(?<!\d)(\d{1,6})\s*/\s*(\d{2})(?!\d)")


def _norm_oc(texto: str) -> set[str]:
    """OC tipo "1450/26" o "02466/25" → {"1450/26", "2466/25"} (sin ceros a la izquierda)."""
    return {f"{int(a)}/{b}" for a, b in _PATRON_OC.findall(texto or "")}


def ocs_mencionadas(r) -> list[str]:
    return sorted(_norm_oc(f'{r.get("comentarios", "")} {r.get("oc", "")}'))


def _agrupar_docs(registros, tipo):
    """Documento → {(año, mes): {total, orig, cant}} y datos del documento."""
    docs: dict = {}
    for r in registros:
        if tipo == "fact":
            k = (r["marca_clave"], clave(r["folio"]))
        else:
            k = (r["marca_clave"], clave(r["referencia"]), clave(r["producto"]))
        d = docs.setdefault(k, {"meses": {}, "regs": [], "cat": set()})
        e = d["meses"].setdefault((r["anio"], r["mes"]), {"total": 0.0, "orig": 0.0, "cant": 0.0, "tc": set(), "moneda": set()})
        e["total"] += r["total"]
        e["orig"] += r["monto_orig"]
        e["cant"] += r.get("cantidad", 0.0)
        e["tc"].add(r["tc"])
        e["moneda"].add((r["moneda_orig"], r["moneda_marca"]))
        d["regs"].append(r)
        d["cat"].add(r["cat"])
    return docs


def _doc_texto(tipo, k, d) -> str:
    r = d["regs"][0]
    if tipo == "fact":
        ocs = sorted({oc for x in d["regs"] for oc in ocs_mencionadas(x)})
        return f'Folio {r["folio"]}' + (f' (OC {", ".join(ocs)})' if ocs else " (sin OC en Comentarios)") + f' · {r["fecha"]:%Y-%m-%d}'
    return f'OC {r["referencia"] or "(vacía)"} · producto {r["producto"] or "(vacío)"} · {r["fecha"]:%Y-%m-%d}'


def comparar_versiones(ant: dict, act: dict) -> dict:
    """Compara dos resultados calculados con la misma lógica, por marca visible, año y mes."""
    corte_f_ant = ant["corte_f"]
    corte_b_ant = ant["corte_b"]
    docs = {
        "fact": (_agrupar_docs(ant["registros_f"], "fact"), _agrupar_docs(act["registros_f"], "fact")),
        "bo": (_agrupar_docs(ant["registros_b"], "bo"), _agrupar_docs(act["registros_b"], "bo")),
    }
    ocs_bo_ant = {(k[0], k[1]) for k in docs["bo"][0]}
    # OC mencionadas por facturas nuevas (folios agregados en la versión actual)
    folios_nuevos = [d for k, d in docs["fact"][1].items() if k not in docs["fact"][0]]
    ocs_facturadas: dict = {}
    for d in folios_nuevos:
        for r in d["regs"]:
            for oc in ocs_mencionadas(r):
                ocs_facturadas.setdefault((r["marca_clave"], oc), set()).add(r["folio"])
    # Facturas nuevas que no mencionan ninguna OC (candidatas a explicar un BO desaparecido)
    facturas_sin_oc: dict = {}
    for d in folios_nuevos:
        if any(ocs_mencionadas(r) for r in d["regs"]):
            continue
        r0 = d["regs"][0]
        facturas_sin_oc.setdefault(r0["marca_clave"], []).append(
            {"folio": r0["folio"], "total": sum(r["total"] for r in d["regs"]), "comentarios": r0["comentarios"]})

    marcas_act = {m["clave"]: m for m in act["marcas"]}
    marcas_ant = {m["clave"]: m for m in ant["marcas"]}
    filas, cumplimientos, forecast, categorias = [], [], [], []

    for mc, m in marcas_act.items():
        periodos = []
        for b in m["bloques"]:
            for x in b["meses"]:
                if (x["anio"], x["mes"]) not in periodos:
                    periodos.append((x["anio"], x["mes"]))
        for (anio, mes) in sorted(periodos):
            fila = {"marca": m["nombre"], "clave": mc, "moneda": m["moneda"], "anio": anio, "mes": mes,
                    "etiqueta": mes_corto(anio, mes), "items": []}
            for tipo, idx_ant, idx_act in (("fact", ant["idx_f"], act["idx_f"]), ("bo", ant["idx_b"], act["idx_b"])):
                v_ant = idx_ant.get((mc, anio, mes), {"total": 0.0})["total"]
                v_act = idx_act.get((mc, anio, mes), {"total": 0.0})["total"]
                fila[tipo] = {"ant": v_ant, "act": v_act, "dif": v_act - v_ant}
                d_ant, d_act = docs[tipo]
                explicado = 0.0
                for k in set(d_ant) | set(d_act):
                    if k[0] != mc:
                        continue
                    ea = d_ant.get(k, {"meses": {}})["meses"].get((anio, mes))
                    en = d_act.get(k, {"meses": {}})["meses"].get((anio, mes))
                    ta, tn = (ea["total"] if ea else 0.0), (en["total"] if en else 0.0)
                    if abs(tn - ta) < 1e-9:
                        continue
                    explicado += tn - ta
                    fila["items"].append(_clasificar(tipo, k, d_ant.get(k), d_act.get(k), ea, en, ta, tn,
                                                     corte_f_ant, corte_b_ant, ocs_bo_ant, ocs_facturadas, facturas_sin_oc))
                fila[tipo]["explicado"] = abs(explicado - fila[tipo]["dif"]) <= 0.005
            fila["factbo"] = {k: fila["fact"][k] + fila["bo"][k] for k in ("ant", "act", "dif")}
            if not fila["items"]:
                continue
            fila["estado"] = "revisar" if any(i["clase"] == "revisar" for i in fila["items"]) else "esperable"
            fila["aviso"] = any(i.get("aviso") for i in fila["items"])
            filas.append(fila)

        # Cumplimientos y forecast por bloque
        m_ant = marcas_ant.get(mc)
        for b in m["bloques"]:
            b_ant = next((x for x in (m_ant["bloques"] if m_ant else []) if x["tipo"] == b["tipo"]), None)
            if not b["calc"]:
                continue
            ca = b_ant["calc"] if b_ant else None
            for q in range(4):
                va = ca["trims"][q]["cumpl"] if ca else None
                vn = b["calc"]["trims"][q]["cumpl"]
                if va != vn and not (va is not None and vn is not None and abs(va - vn) < 1e-4):
                    cumplimientos.append({"marca": m["nombre"], "bloque": b["titulo"], "nombre": f"Q{q + 1}", "ant": va, "act": vn})
            va, vn = (ca["total"]["cumpl"] if ca else None), b["calc"]["total"]["cumpl"]
            if va != vn and not (va is not None and vn is not None and abs(va - vn) < 1e-4):
                cumplimientos.append({"marca": m["nombre"], "bloque": b["titulo"], "nombre": "Anual", "ant": va, "act": vn})
            if b_ant:
                for i, x in enumerate(b["meses"]):
                    fa, fn = b_ant["fc_mes"][i] if i < len(b_ant["fc_mes"]) else None, b["fc_mes"][i]
                    if (fa or 0) != (fn or 0) and abs((fa or 0) - (fn or 0)) > 1e-6:
                        forecast.append({"marca": m["nombre"], "bloque": b["titulo"], "periodo": mes_corto(x["anio"], x["mes"]), "ant": fa, "act": fn})
                if (b_ant["fc_total"] or 0) != (b["fc_total"] or 0):
                    forecast.append({"marca": m["nombre"], "bloque": b["titulo"], "periodo": "Total", "ant": b_ant["fc_total"], "act": b["fc_total"]})

    # Cambios de categoría (documentos presentes en ambas versiones)
    for tipo in ("fact", "bo"):
        d_ant, d_act = docs[tipo]
        for k in set(d_ant) & set(d_act):
            if k[0] in marcas_act and d_ant[k]["cat"] != d_act[k]["cat"]:
                categorias.append({"marca": marcas_act[k[0]]["nombre"], "tipo": tipo, "doc": _doc_texto(tipo, k, d_act[k]),
                                   "ant": ", ".join(sorted(d_ant[k]["cat"])), "act": ", ".join(sorted(d_act[k]["cat"]))})

    # Cambios en Parametros Moneda
    parametros = []
    pa, pn = ant["params"], act["params"]
    nombres = {p["clave"]: p["marca"].strip() for p in pa["lista_marcas"] + pn["lista_marcas"]}

    def tc_txt(v):
        return "(no estaba)" if v is None else (fmt_cl(v, 6).rstrip("0").rstrip(",") if es_numerico(v) else repr(v))

    for k in sorted(set(pa["monedas"]) | set(pn["monedas"])):
        if pa["monedas"].get(k) != pn["monedas"].get(k):
            parametros.append(f'Moneda de reporte de {nombres.get(k, k)}: {pa["monedas"].get(k) or "(no estaba)"} → {pn["monedas"].get(k) or "(se quitó)"}')
    for k in sorted(set(pa["pares"]) | set(pn["pares"])):
        par = (pn["pares"].get(k) or pa["pares"].get(k))["par"]
        ta, tn = (pa["pares"].get(k) or {}).get("tc"), (pn["pares"].get(k) or {}).get("tc")
        if ta != tn:
            parametros.append(f'Tipo de cambio {par}: {tc_txt(ta)} → {tc_txt(tn) if tn is not None else "(se quitó)"}')

    return {
        "anterior": {"archivo": ant["ruta"].name, "fecha": ant["fecha_archivo"], "corteFact": ant["corte_f"], "corteBO": ant["corte_b"]},
        "actual": {"archivo": act["ruta"].name, "fecha": act["fecha_archivo"], "corteFact": act["corte_f"], "corteBO": act["corte_b"]},
        "filas": filas, "cumplimientos": cumplimientos, "forecast": forecast,
        "categorias": categorias, "parametros": parametros,
    }


def _clasificar(tipo, k, doc_ant, doc_act, ea, en, ta, tn, corte_f_ant, corte_b_ant, ocs_bo_ant, ocs_facturadas,
                facturas_sin_oc=None) -> dict:
    """Clasifica la contribución de un documento a la diferencia de un mes."""
    dif = tn - ta
    doc = doc_act or doc_ant
    texto = _doc_texto(tipo, k, doc)
    mismo_orig = ea and en and abs(ea["orig"] - en["orig"]) < 1e-6 and ea["moneda"] == en["moneda"] and ea["cant"] == en["cant"]
    if ea and en and mismo_orig and ea["tc"] != en["tc"]:
        return {"tipo": tipo, "clase": "revisar", "causa": "Cambio de tipo de cambio (Parametros Moneda)", "doc": texto,
                "ant": ta, "act": tn, "dif": dif, "tcAnt": sorted(x for x in ea["tc"] if x is not None), "tcAct": sorted(x for x in en["tc"] if x is not None)}
    if tipo == "fact":
        if doc_ant is None:
            fecha = min(r["fecha"] for r in doc_act["regs"])
            if corte_f_ant is None or fecha > corte_f_ant:
                return {"tipo": tipo, "clase": "esperable", "causa": "Factura nueva posterior al corte anterior", "doc": texto, "ant": ta, "act": tn, "dif": dif}
            return {"tipo": tipo, "clase": "revisar", "causa": "Folio nuevo con fecha retroactiva (anterior al corte anterior)", "doc": texto, "ant": ta, "act": tn, "dif": dif}
        if doc_act is None:
            return {"tipo": tipo, "clase": "revisar", "causa": "Folio que desaparece", "doc": texto, "ant": ta, "act": tn, "dif": dif}
        if ea is None or en is None:
            return {"tipo": tipo, "clase": "revisar", "causa": "Folio que cambia de fecha", "doc": texto, "ant": ta, "act": tn, "dif": dif}
        return {"tipo": tipo, "clase": "revisar", "causa": "Folio que cambia de monto", "doc": texto, "ant": ta, "act": tn, "dif": dif}
    # BO
    oc = (k[0], next(iter(_norm_oc(doc["regs"][0]["referencia"])), clave(doc["regs"][0]["referencia"])))
    cant = {"cantAnt": ea["cant"] if ea else 0.0, "cantAct": en["cant"] if en else 0.0}
    if dif < 0:
        facturas = sorted(ocs_facturadas.get(oc, set()))
        if facturas:
            return {"tipo": tipo, "clase": "esperable", "causa": "BO recibido/facturado (factura nueva que menciona la OC)",
                    "doc": texto + f' · facturas {", ".join(facturas)}', "ant": ta, "act": tn, "dif": dif, **cant}
        # ¿Una única factura nueva de la marca, sin OC en Comentarios, por el mismo monto?
        monto = -dif
        candidatas = [f for f in (facturas_sin_oc or {}).get(k[0], [])
                      if abs(f["total"] - monto) < max(COINCIDENCIA_PCT * monto, COINCIDENCIA_ABS)]
        if len(candidatas) == 1:
            f = candidatas[0]
            return {"tipo": tipo, "clase": "esperable", "aviso": True, "causa": "Probable factura sin referencia de OC",
                    "doc": texto + f' · probable folio {f["folio"]} ({fmt_monto(f["total"])}, "{f["comentarios"]}")',
                    "ant": ta, "act": tn, "dif": dif, **cant}
        causa = "BO que desaparece sin factura nueva que mencione la OC" if not en or en["total"] == 0 else "BO que baja sin factura nueva que mencione la OC"
        if len(candidatas) > 1:
            texto += f' · {len(candidatas)} facturas sin OC con monto similar: {", ".join(f["folio"] for f in candidatas)}'
        return {"tipo": tipo, "clase": "revisar", "causa": causa, "doc": texto, "ant": ta, "act": tn, "dif": dif, **cant}
    if doc_ant is None and (k[0], k[1]) not in ocs_bo_ant:
        return {"tipo": tipo, "clase": "esperable", "causa": "OC nueva", "doc": texto, "ant": ta, "act": tn, "dif": dif, **cant}
    fecha_oc = min(r["fecha"] for r in doc["regs"])
    if corte_b_ant is not None and (corte_b_ant.date() - fecha_oc.date()) > timedelta(days=DIAS_OC_ANTIGUA):
        return {"tipo": tipo, "clase": "revisar", "causa": "BO que sube en una OC antigua", "doc": texto, "ant": ta, "act": tn, "dif": dif, **cant}
    return {"tipo": tipo, "clase": "esperable", "causa": f"Línea nueva o aumento en OC reciente (≤ {DIAS_OC_ANTIGUA} días antes del corte anterior)",
            "doc": texto, "ant": ta, "act": tn, "dif": dif, **cant}


def _estado_md(f) -> str:
    if f["estado"] == "revisar":
        return "⚠ REVISAR"
    return "✔ esperable (ⓘ con aviso)" if f.get("aviso") else "✔ esperable"


def reporte_markdown(cmp: dict) -> str:
    a, n = cmp["anterior"], cmp["actual"]
    L = [f'# Cambios {a["fecha"]:%Y-%m-%d} → {n["fecha"]:%Y-%m-%d}', "",
         f'- Versión anterior: `{a["archivo"]}` · corte facturación {a["corteFact"]:%Y-%m-%d} · corte BO {a["corteBO"]:%Y-%m-%d}',
         f'- Versión actual: `{n["archivo"]}` · corte facturación {n["corteFact"]:%Y-%m-%d} · corte BO {n["corteBO"]:%Y-%m-%d}',
         "- Montos en la moneda de reporte de cada marca. ESPERABLE = movimiento normal; REVISAR = requiere explicación; "
         "ⓘ aviso = esperable, pero conviene confirmarlo.", ""]
    revisar = [f for f in cmp["filas"] if f["estado"] == "revisar"]
    L += [f'**{len(cmp["filas"])} meses con cambios** ({len(revisar)} a revisar) · '
          f'{len(cmp["cumplimientos"])} cumplimientos que cambian · {len(cmp["forecast"])} cambios de forecast · '
          f'{len(cmp["categorias"])} cambios de categoría · {len(cmp["parametros"])} cambios en {HOJA_MONEDA}', ""]
    L += ["## Resumen por marca y mes", "", "| Marca | Mes | Δ Facturación | Δ BO | Δ Fact+BO | Estado |", "|---|---|---:|---:|---:|---|"]
    for f in cmp["filas"]:
        L.append(f'| {f["marca"]} ({f["moneda"]}) | {f["etiqueta"]} | {fmt_monto(f["fact"]["dif"])} | {fmt_monto(f["bo"]["dif"])} | '
                 f'{fmt_monto(f["factbo"]["dif"])} | {_estado_md(f)} |')
    L += ["", "## Detalle por marca y mes", ""]
    for f in cmp["filas"]:
        L.append(f'### {f["marca"]} · {f["etiqueta"]} — {_estado_md(f)}')
        L.append("")
        for tipo, nombre in (("fact", "Facturación"), ("bo", "Backorder")):
            v = f[tipo]
            if abs(v["dif"]) < 1e-9:
                continue
            items = [i for i in f["items"] if i["tipo"] == tipo]
            suma = sum(i["dif"] for i in items)
            L.append(f'**{nombre}:** {fmt_monto(v["ant"])} → {fmt_monto(v["act"])} {f["moneda"]} (Δ {fmt_monto(v["dif"])}) · '
                     f'suma del detalle {fmt_monto(suma)} {"✔ explica exactamente la diferencia" if v["explicado"] else "✕ NO explica la diferencia"}')
            L.append("")
            L.append("| Clase | Causa | Documento | Antes | Ahora | Δ |")
            L.append("|---|---|---|---:|---:|---:|")
            for i in sorted(items, key=lambda x: (x["clase"] != "revisar", x["causa"], -abs(x["dif"]))):
                extra = ""
                if "cantAnt" in i:
                    extra = f' · cant. {fmt_monto(i["cantAnt"])} → {fmt_monto(i["cantAct"])}'
                if "tcAnt" in i:
                    extra = f' · TC {i["tcAnt"]} → {i["tcAct"]}'
                clase = "⚠ revisar" if i["clase"] == "revisar" else ("esperable (ⓘ aviso)" if i.get("aviso") else "esperable")
                L.append(f'| {clase} | {i["causa"]} | {i["doc"]}{extra} | '
                         f'{fmt_monto(i["ant"])} | {fmt_monto(i["act"])} | {fmt_monto(i["dif"])} |')
            L.append("")
    L += ["## Cumplimientos que cambian", ""]
    if cmp["cumplimientos"]:
        L += ["| Marca | Bloque | Período | Antes | Ahora |", "|---|---|---|---:|---:|"]
        for c in cmp["cumplimientos"]:
            pa = "—" if c["ant"] is None else fmt_cl(c["ant"] * 100, 1) + "%"
            pn = "—" if c["act"] is None else fmt_cl(c["act"] * 100, 1) + "%"
            L.append(f'| {c["marca"]} | {c["bloque"]} | {c["nombre"]} | {pa} | {pn} |')
    else:
        L.append("Sin cambios.")
    L += ["", "## Cambios de forecast (⚠ revisar)", ""]
    L += [f'- {c["marca"]} · {c["bloque"]} · {c["periodo"]}: {fmt_monto(c["ant"])} → {fmt_monto(c["act"])}' for c in cmp["forecast"]] or ["Sin cambios."]
    L += ["", f"## Cambios en {HOJA_MONEDA} (⚠ revisar)", "",
          "Un cambio de tipo de cambio o la incorporación de una marca cambia montos sin que cambien los documentos.", ""]
    L += [f"- {p}" for p in cmp["parametros"]] or ["Sin cambios."]
    L += ["", "## Cambios de categoría (⚠ revisar)", ""]
    L += [f'- {c["marca"]} · {"Facturación" if c["tipo"] == "fact" else "BO"} · {c["doc"]}: {c["ant"]} → {c["act"]}' for c in cmp["categorias"]] or ["Sin cambios."]
    return "\n".join(L) + "\n"


def resumen_cambios_json(cmp: dict) -> dict:
    """Resumen para el panel: por marca y mes, sin OC, productos ni precios."""
    def fecha(d):
        return d.strftime("%Y-%m-%d") if d else None

    filas = []
    for f in cmp["filas"]:
        causas: dict = {}
        for i in f["items"]:
            k = (i["clase"], i["causa"], i["tipo"])
            e = causas.setdefault(k, {"clase": i["clase"], "aviso": bool(i.get("aviso")), "causa": i["causa"], "tipo": i["tipo"], "n": 0, "monto": 0.0})
            e["n"] += 1
            e["monto"] += i["dif"]
        filas.append({"marca": f["marca"], "moneda": f["moneda"], "anio": f["anio"], "mes": f["mes"], "etiqueta": f["etiqueta"],
                      "fact": f["fact"], "bo": f["bo"], "factbo": f["factbo"], "estado": f["estado"], "aviso": f["aviso"],
                      "explicado": f["fact"]["explicado"] and f["bo"]["explicado"],
                      "causas": sorted(causas.values(), key=lambda c: (c["clase"] != "revisar", -abs(c["monto"])))})
    return {
        "anterior": {k: (fecha(v) if isinstance(v, (date, datetime)) else v) for k, v in cmp["anterior"].items()},
        "actual": {k: (fecha(v) if isinstance(v, (date, datetime)) else v) for k, v in cmp["actual"].items()},
        "filas": filas, "cumplimientos": cmp["cumplimientos"], "forecast": cmp["forecast"],
        "parametros": cmp["parametros"], "nCategorias": len(cmp["categorias"]),
    }


# ──────────────────────────────────────────────────────────────────────
# SERIALIZACIÓN
# ──────────────────────────────────────────────────────────────────────

def _fila_detalle_fact(r):
    return {"fecha": r["fecha"].strftime("%Y-%m-%d"), "anio": r["anio"], "mes": r["mes"], "folio": r["folio"],
            "comentarios": r["comentarios"], "oc": r["oc"], "cat": r["cat"], "catRaw": r["cat_raw"],
            "monedaOrig": r["moneda_orig"], "montoOrig": r["monto_orig"], "tc": r["tc"], "total": r["total"],
            "convertido": r["convertido"]}


def _fila_detalle_bo(r):
    return {"fecha": r["fecha"].strftime("%Y-%m-%d"), "anio": r["anio"], "mes": r["mes"], "referencia": r["referencia"],
            "producto": r["producto"], "descripcion": r["descripcion"], "cat": r["cat"], "cantidad": r["cantidad"],
            "monedaOrig": r["moneda_orig"], "precio": r["precio"], "montoOrig": r["monto_orig"], "tc": r["tc"],
            "total": r["total"], "convertido": r["convertido"]}


def _serializar_marca(m: dict, incluir_detalle: bool) -> dict:
    salida = {k: m[k] for k in ("hoja", "nombre", "clave", "moneda", "tieneFiscal", "convierte", "paresConversion", "sinFilas")}
    salida["bloques"] = []
    for b in m["bloques"]:
        sb = {"tipo": b["tipo"], "titulo": b["titulo"], "c0": b["c0"], "errores": b["errores"],
              "meses": [{"anio": x["anio"], "mes": x["mes"]} for x in b["meses"]],
              "conc": b["conc"], "conversion": b["conversion"], "calc": None, "vista": b.get("vista"),
              "celdas": {"n": len(b.get("celdas") or []), "ok": sum(1 for i in b.get("celdas") or [] if i["ok"]),
                         "fallas": [i for i in b.get("celdas") or [] if not i["ok"]]}}
        c = b["calc"]
        if c:
            sb["calc"] = {
                "meses": [{k: v for k, v in x.items() if not k.startswith("_")} for x in c["meses"]],
                **{k: c[k] for k in ("trims", "total", "ytd", "catF", "catB", "rawF", "rawB", "fcSumaMeses", "tieneForecast", "hayDerivado")},
            }
        salida["bloques"].append(sb)
    if incluir_detalle:
        periodos = {(x["anio"], x["mes"]) for b in m["bloques"] for x in b["meses"]}
        salida["detalle"] = {
            "facturas": [_fila_detalle_fact(r) for r in m["registros"] if r["tipo"] == "fact" and (r["anio"], r["mes"]) in periodos],
            "bo": [_fila_detalle_bo(r) for r in m["registros"] if r["tipo"] == "bo" and r["cantidad"] > 0],
        }
    return salida


def serializar(res: dict, incluir_detalle: bool = INCLUIR_DETALLE, cambios: dict | None = None, avisos_seleccion=None) -> dict:
    def fecha_txt(d):
        return d.strftime("%Y-%m-%d") if d else None

    problemas = [{"nivel": "aviso", "titulo": a} for a in (avisos_seleccion or [])] + [
        {k: v for k, v in p.items() if k != "filas"} for p in res["problemas"]]
    if not incluir_detalle:
        # Sin detalle: no publicar OC, productos ni montos por documento en los textos de validación.
        for p in problemas:
            if p.get("detalle") and p["nivel"] in ("aviso", "error") and "fila" in p["detalle"]:
                p["detalle"] = "(detalle omitido: INCLUIR_DETALLE = False; ver la salida de generar_datos.py)"
    return {
        "meta": {
            "archivo": res["ruta"].name,
            "fechaArchivo": fecha_txt(res["fecha_archivo"]),
            "generado": datetime.now().isoformat(timespec="seconds"),
            "incluirDetalle": incluir_detalle,
            "hojas": {"facturacion": HOJA_FACTURACION, "bo": HOJA_BO, "moneda": HOJA_MONEDA},
            "fact": {"hoja": res["t_f"]["nombre"], "filas": len(res["t_f"]["filas"]), "corte": fecha_txt(res["corte_f"])},
            "bo": {"hoja": res["t_b"]["nombre"], "filas": len(res["t_b"]["filas"]), "corte": fecha_txt(res["corte_b"])},
            "corte": fecha_txt(res["corte"]),
            "tolerancia": {"monto": TOLERANCIA_MONTO, "pct": TOLERANCIA_PCT},
            "categorias": CATEGORIAS, "sinCategoria": SIN_CATEGORIA,
        },
        "resumen": res["resumen"],
        "problemas": problemas,
        "ocultas": res["ocultas"],
        "parametros": {"pares": {v["par"]: {"tc": v["tc"], "celda": v["celda"]} for v in res["params"]["pares"].values()},
                       "marcas": [{"marca": p["marca"], "moneda": p["moneda"]} for p in res["params"]["lista_marcas"]]},
        "marcas": [_serializar_marca(m, incluir_detalle) for m in res["marcas"]],
        "cambios": cambios,
    }


def procesar_libro(ruta: Path | str, incluir_detalle: bool = INCLUIR_DETALLE, permitir_sin_categoria: bool = False) -> dict:
    """Atajo: calcula y serializa un libro (sin comparación)."""
    return serializar(calcular(ruta, permitir_sin_categoria), incluir_detalle)


def _redondear(v):
    if isinstance(v, float):
        return round(v, 6)
    if isinstance(v, dict):
        return {k: _redondear(x) for k, x in v.items()}
    if isinstance(v, list):
        return [_redondear(x) for x in v]
    return v


# ──────────────────────────────────────────────────────────────────────
# SALIDA POR CONSOLA
# ──────────────────────────────────────────────────────────────────────

def imprimir_conciliacion(datos: dict) -> None:
    for m in datos["marcas"]:
        for b in m["bloques"]:
            print(f'\n## {m["nombre"]} ({m["moneda"]}) · {b["titulo"]}')
            for e in b["errores"]:
                print(f"  ✕ {e}")
            for it in b["conc"]:
                fmt = (lambda v: "—" if v is None else fmt_cl(v * 100, 1) + "%") if it["tipo"] == "pct" else (lambda v: "—" if v is None else fmt_cl(v))
                ex = "no numérica" if it["excelNoNumerico"] else fmt(it["excel"] if it["excel"] is not None else 0.0)
                print(f'  {"✔" if it["ok"] else "✕"} {it["concepto"]:<42} {it["celda"]:<18} excel={ex:>15} app={fmt(it["app"]):>15}'
                      + ("" if it["ok"] else f'   dif={it["dif"]}'))


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Genera docs/datos_cumplimiento.json desde el Excel de data/raw/.")
    ap.add_argument("--permitir-sin-categoria", action="store_true",
                    help="continuar aunque haya documentos de marcas visibles sin categoría estándar (queda como advertencia en el panel)")
    args = ap.parse_args(argv[1:])

    try:
        sel = seleccionar_actual()
        print(f'Archivo: {sel["ruta"]} (datos al {sel["fecha"]:%Y-%m-%d})')
        for a in sel["avisos"]:
            print(f"  [aviso] {a}")
        res = calcular(sel["ruta"], args.permitir_sin_categoria)
    except ErrorCarga as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    r = res["resumen"]
    print(f'Facturación: hoja "{res["t_f"]["nombre"]}" · {len(res["t_f"]["filas"])} filas · corte {res["corte_f"]:%Y-%m-%d}')
    print(f'Backorder:   hoja "{res["t_b"]["nombre"]}" · {len(res["t_b"]["filas"])} filas · corte {res["corte_b"]:%Y-%m-%d}')
    print("\nValidaciones:")
    for p in res["problemas"]:
        print(f'  [{p["nivel"]}] {p["titulo"]}')
        for linea in p.get("filas") or ([p["detalle"]] if p.get("detalle") else []):
            print(f"           {linea}")
    datos = serializar(res, INCLUIR_DETALLE, None, sel["avisos"])
    imprimir_conciliacion(datos)
    print(f'\nConciliación: {r["ok"]}/{r["comparaciones"]} cuadran · {r["diferencias"]} diferencias · '
          f'{r["errores"]} errores ({r["erroresBloqueantes"]} bloqueantes) · {r["avisos"]} avisos')

    if r["diferencias"] or r["erroresBloqueantes"]:
        print(f"\nERROR: hay validaciones que no cuadran. NO se escribió {SALIDA_JSON}.", file=sys.stderr)
        return 1

    # Comparación contra la versión anterior
    cambios = None
    try:
        ant_sel = seleccionar_anterior(sel["fecha"])
    except ErrorCarga as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1
    if ant_sel is None:
        print(f"\n[aviso] No hay una versión anterior en {CARPETA_EXCEL}/{NOMBRE_CARPETA_ANTERIORES}/: se omite el reporte de cambios.")
    else:
        print(f'\nComparando contra {ant_sel["ruta"].name} …')
        try:
            ant = calcular(ant_sel["ruta"], permitir_sin_categoria=True)
        except ErrorCarga as e:
            print(f"[aviso] No se pudo leer la versión anterior ({e}); se omite el reporte de cambios.")
            ant = None
        if ant is not None:
            cmp = comparar_versiones(ant, res)
            CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
            ruta_rep = CARPETA_REPORTES / f'cambios_{ant_sel["fecha"]:%Y%m%d}_a_{sel["fecha"]:%Y%m%d}.md'
            ruta_rep.write_text(reporte_markdown(cmp), encoding="utf-8")
            rev = sum(1 for f in cmp["filas"] if f["estado"] == "revisar")
            no_expl = [f for f in cmp["filas"] if not (f["fact"]["explicado"] and f["bo"]["explicado"])]
            print(f'Reporte de cambios: {ruta_rep} · {len(cmp["filas"])} meses con cambios ({rev} a revisar)')
            if no_expl:
                print(f"[aviso] {len(no_expl)} meses cuya diferencia no se explica exactamente con el detalle.")
            cambios = resumen_cambios_json(cmp)

    datos = serializar(res, INCLUIR_DETALLE, cambios, sel["avisos"])
    SALIDA_JSON.parent.mkdir(parents=True, exist_ok=True)
    with open(SALIDA_JSON, "w", encoding="utf-8") as f:
        json.dump(_redondear(datos), f, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    print(f"\nOK: {SALIDA_JSON} ({SALIDA_JSON.stat().st_size / 1024:,.0f} KB, detalle {'incluido' if INCLUIR_DETALLE else 'omitido'}).")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
