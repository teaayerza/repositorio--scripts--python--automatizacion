# -*- coding: utf-8 -*-
"""
OBTENER DATOS DE FACTURAS DESDE CAE
===================================
Flujo general del script:

  1. Lee el listado de facturas (Listado_CAE_14_digitos.xlsx) y los maestros (empresas, proveedores, fechas).
  2. Para cada factura busca su PDF en la carpeta de facturas pendientes.
  3. Intenta leer el QR fiscal de ARCA/AFIP (fuente principal y oficial de los datos).
  4. Completa lo que falte con el texto que ya tiene el PDF (OCR) y con el nombre del archivo.
  5. Calcula la fecha de vencimiento.
  6. Guarda todo en resultado_facturas_cae.xlsx y en un log.

REGLA DE IMPORTE (dato sensible)
--------------------------------
El IMPORTE sale de UNA sola fuente confiable por prioridad:
  a) QR fiscal de ARCA/AFIP (importe oficial).
  b) Si no hay QR: TOTAL del texto del PDF, SOLO si cierra aritmeticamente
     (el TOTAL tiene que ser igual a la suma de otros montos del mismo documento: subtotal + IVA [+ impuestos]).
  c) Si no se puede verificar: IMPORTE = NO_DETECTADO y queda una observacion para cargarlo a mano.
Ya NO se estima el importe (ni "mayor monto", ni "ultimo numero de la linea", ni la palabra IMPORTE).
Prefiere dejar el dato vacio antes que cargar uno incorrecto.

REGLA DE PUNTO DE VENTA / NUMERO DE COMPROBANTE (dato fiscal)
-------------------------------------------------------------
Se informa TAL CUAL ESTA IMPRESO en la factura (si dice 0012-00007353 se informa 0012 y 00007353).
  a) El QR de ARCA entrega el VALOR (ptoVta=12, nroCmp=7353) pero NO la cantidad de ceros impresos.
  b) El texto del PDF entrega la FORMA impresa. Solo se acepta si su valor numerico coincide con el del QR.
  c) QR + texto coinciden  -> se informa el texto literal (fuente exacta).
  d) QR sin texto utilizable -> se completa con ceros (5 / 8) y queda una observacion para verificar.
  e) Sin QR -> texto literal con validaciones; nunca se rellena ni se trunca en silencio.
"""
from __future__ import annotations

# =============================================================================
# 1. IMPORTACIONES
# =============================================================================
import base64
import json
import logging
import re
import sys
import time
import unicodedata
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import parse_qs, urlparse

import cv2
import zxingcpp                      # lector principal de QR (zxing-cpp)
import numpy as np
import pymupdf as fitz
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.datetime import from_excel

# =============================================================================
# 2. CONFIGURACION: nombres, carpetas y parametros
# =============================================================================

# --- Lectura del QR ---
QR_DPI = 200                       # resolucion con la que se renderiza SOLO el recuadro del QR
QR_PAGINA_COMPLETA = True          # ultimo recurso: pagina completa a QR_DPI (primeras 3 paginas y la ultima)

# --- Archivos y carpetas del proyecto (se buscan a partir de "98. AUTOMATIZACION") ---
LISTADO_NOMBRE = "Listado_CAE_14_digitos.xlsx"
SALIDA_NOMBRE = "resultado_facturas_cae.xlsx"
LOG_NOMBRE = "procesamiento_cae.xlsx"
CARPETA_FACTURAS = "01. FACTURAS PENDIENTES CONTABILIZAR (AI)"
CARPETA_MAESTROS = Path("02. Archivos Maestros") / "Maestro Unificado"
CARPETA_RESULTADO = "07. Resultado Final"
CARPETA_LOGS = "05. Logs"
CARPETA_INTERMEDIA = "04. Resultados"

# --- Respaldo cuando el QR no se puede leer (se usa el texto que ya tiene el PDF) ---
MIN_CARACTERES_TEXTO = 120         # menos caracteres que esto = el PDF no tiene capa de texto utilizable
PERMITIR_CUIT_FUERA_MAESTRO = False # el CUIT emisor debe existir en el maestro de proveedores

# --- Verificacion aritmetica del TOTAL (texto del PDF) ---
LINEAS_VENTANA_TOTAL = 6           # cuantas lineas despues de la palabra TOTAL se revisan buscando montos
MAX_TERMINOS_SUMA = 4              # el TOTAL puede ser suma de 2 a 4 montos (subtotal + IVA 21 + IVA 10.5 + percepcion)
TOLERANCIA_CENTAVOS = 1            # diferencia aceptada por redondeo (1 centavo)
ALICUOTAS_IVA = (0.21, 0.105, 0.27, 0.05, 0.025)   # si el IVA de un monto figura en el documento, ese monto es un NETO, no el total

# --- Punto de venta / numero de comprobante: se conservan TAL CUAL estan impresos ---
MAX_DIG_PTO = 5                    # ARCA admite hasta 5 digitos de punto de venta
MAX_DIG_NRO = 8                    # ARCA admite hasta 8 digitos de numero de comprobante
ANCHO_PTO_DEFECTO = 5              # SOLO si hay QR pero el texto no deja ver como se imprimio (se deja observacion)
ANCHO_NRO_DEFECTO = 8

# --- Columnas del Excel final (en este orden) ---
COLUMNAS = [
    "EMPRESA_RECEPTORA", "PROVEEDOR", "TIPO_COMPROBANTE", "PUNTO_VENTA", "NUMERO_COMPROBANTE",
    "FECHA_REFERENCIA_VENCIMIENTO", "FECHA_VENCIMIENTO_CALCULADA", "ORIGEN_VENCIMIENTO", "IMPORTE_TOTAL", "MONEDA",
    "CUIT_EMISOR_MAESTRO_DE_PROVEEDORES", "CUIT_EMISOR_FACTURA", "CAE", "NOMBRE_ARCHIVO", "RUTA_COMPLETA", "FECHA_MODIFICACION_ARCHIVO",
    "ESTADO_QR", "FECHA_EMISION", "FUENTE_DATOS", "OBSERVACIONES", "VALIDACION_PTO_NRO", "FECHA_HORA_PROCESAMIENTO",
]

# --- Codigos de comprobante de ARCA/AFIP ---
TIPOS = {
    1: "FACTURA A", 2: "ND A", 3: "NC A", 4: "RECIBO A", 5: "NV CONTADO A",
    6: "FACTURA B", 7: "ND B", 8: "NC B", 9: "RECIBO B", 10: "NV CONTADO B",
    11: "FACTURA C", 12: "ND C", 13: "NC C", 15: "RECIBO C", 16: "NV CONTADO C",
    19: "FACTURA EXPORTACION", 20: "ND EXTERIOR", 21: "NC EXTERIOR",
    39: "OTROS COMPROBANTES A (RG 1415)", 40: "OTROS COMPROBANTES B (RG 1415)", 41: "OTROS COMPROBANTES C (RG 1415)",
    51: "FACTURA M", 52: "ND M", 53: "NC M", 54: "RECIBO M", 55: "NV CONTADO M",
    60: "CTA VENTA Y LIQ PRODUCTO A", 61: "CTA VENTA Y LIQ PRODUCTO B", 63: "LIQUIDACION A", 64: "LIQUIDACION B",
    201: "FCE A", 202: "ND FCE A", 203: "NC FCE A", 206: "FCE B", 207: "ND FCE B", 208: "NC FCE B",
    211: "FCE C", 212: "ND FCE C", 213: "NC FCE C",
}
NC = {3, 8, 13, 21, 53, 203, 208, 213}   # codigos de notas de credito: su importe se guarda en negativo

# --- Nombres alternativos de columnas aceptados en los maestros ---
ALIASES = {
    "cuit": {"CUIT", "CUIT_NUMERO", "CUIT_EMISOR", "CUIT_PROVEEDOR", "NRO_DOCUMENTO", "DOCUMENTO"},
    "cuit_proveedor": {"CUIT_PROVEEDOR", "CUIT_EMISOR", "CUIT", "CUIT_NUMERO", "NRO_DOCUMENTO", "DOCUMENTO"},
    "cuit_receptor": {"CUIT_RECEPTOR", "CUIT_EMPRESA", "CUIT", "CUIT_NUMERO", "NRO_DOCUMENTO", "DOCUMENTO"},
    "proveedor": {"PROVEEDOR", "RAZON_SOCIAL", "RAZONSOCIAL", "NOMBRE", "DENOMINACION"},
    "empresa": {"EMPRESA", "SOCIEDAD", "RAZON_SOCIAL", "RAZONSOCIAL", "NOMBRE"},
    "archivo": {"ARCHIVO", "NOMBRE_DE_ARCHIVO", "NOMBRE_ARCHIVO"},
    "fecha_recepcion": {"FECHA_RECEPCION", "FECHA_INGRESO", "FECHA"},
    "fecha_modificacion": {"FECHA_MODIFICACION_ARCHIVO", "FECHA_MODIFICACION", "MODIFICADO"},
}

# --- Validaciones de CUIT y control de completitud ---
PREF_CUIT = {"20", "23", "24", "25", "26", "27", "30", "33", "34", "50", "51", "55"}   # prefijos validos de CUIT
CRITICOS = ("TIPO_COMPROBANTE", "PUNTO_VENTA", "NUMERO_COMPROBANTE", "IMPORTE_TOTAL", "MONEDA", "CUIT_EMISOR_MAESTRO_DE_PROVEEDORES", "PROVEEDOR", "EMPRESA_RECEPTORA")
FALTA = {"NO_DETECTADO", "NO_DETECTADA", "NO_IDENTIFICADA", "NO_ENCONTRADO_EN_MAESTRO", ""}   # valores que cuentan como "dato faltante"

MOTIVOS_IMPORTE = {
    "SIN_TEXTO": "el PDF no tiene montos legibles",
    "SIN_PALABRA_TOTAL": "no se encontro la palabra TOTAL seguida de montos",
    "TOTAL_NO_CIERRA": "ningun monto junto a TOTAL cierra como suma de otros montos (subtotal + IVA/impuestos)",
    "TOTAL_PARECE_SUBTOTAL": "el mayor monto que cierra parece un subtotal (su IVA figura en el documento pero no el total)",
}

# =============================================================================
# 3. UTILIDADES GENERALES
# =============================================================================
_ACENTOS = str.maketrans("áéíóúüÁÉÍÓÚÜ", "aeiouuAEIOUU")

def sin_acentos(s):
    return s.translate(_ACENTOS)

def campos_faltantes(reg):
    return [k for k in CRITICOS if reg.get(k) in FALTA]

def norm(v):
    t = "" if v is None else str(v).strip().upper()
    t = "".join(c for c in unicodedata.normalize("NFD", t) if unicodedata.category(c) != "Mn")
    return re.sub(r"[^A-Z0-9]+", "_", t).strip("_")

def digs(v):
    return re.sub(r"\D", "", "" if v is None else str(v))

def a_entero(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, float):
        return int(v) if v.is_integer() else None
    m = re.fullmatch(r"\s*0*(\d+)(?:\.0+)?\s*", str(v))
    return int(m.group(1)) if m else None

def rellenar(v, ancho):
    return str(v).zfill(ancho)

def recortar_ceros(d, maximo):
    d = digs(d)
    if not d:
        return None, False
    if len(d) <= maximo:
        return d, False
    sobrante = len(d) - maximo
    if set(d[:sobrante]) == {"0"}:
        return d[sobrante:], True
    return None, False

def fecha(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            return from_excel(v).date()
        except Exception:
            return None
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(str(v).strip(), f).date()
        except Exception:
            pass
    return None

def numero(s):
    if s is None:
        return None
    x = re.sub(r"[^0-9,.-]", "", str(s))
    neg = x.startswith("-")
    x = x.lstrip("+-")
    if not x:
        return None
    if "," in x and "." in x:
        x = x.replace(".", "").replace(",", ".") if x.rfind(",") > x.rfind(".") else x.replace(",", "")
    elif "," in x:
        a, b = x.rsplit(",", 1)
        x = a.replace(".", "") + "." + b if len(b) <= 2 else x.replace(",", "")
    elif "." in x:
        p = x.split(".")
        x = "".join(p) if len(p) > 2 or (len(p) == 2 and len(p[1]) == 3) else x
    try:
        n = float(x)
        return -n if neg else n
    except Exception:
        return None

def cuit_valido(c):
    c = digs(c)
    if len(c) != 11:
        return False
    pesos = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    r = 11 - sum(int(c[i]) * pesos[i] for i in range(10)) % 11
    dv = 0 if r == 11 else 9 if r == 10 else r
    return dv == int(c[-1])

# =============================================================================
# 4. TEXTO NATIVO DEL PDF
# =============================================================================
def filas_visuales(palabras):
    ps = [p for p in palabras if str(p[4]).strip()]
    if not ps:
        return []
    ps.sort(key=lambda p: (p[1] + p[3]) / 2)
    alt = sorted(p[3] - p[1] for p in ps)[len(ps) // 2] or 8
    tol = alt * 0.6
    filas = []
    cur = [ps[0]]
    cy = (ps[0][1] + ps[0][3]) / 2
    for p in ps[1:]:
        y = (p[1] + p[3]) / 2
        if abs(y - cy) <= tol:
            cur.append(p)
            cy = sum((q[1] + q[3]) / 2 for q in cur) / len(cur)
        else:
            filas.append(cur)
            cur = [p]
            cy = y
    filas.append(cur)
    return [" ".join(str(q[4]) for q in sorted(f, key=lambda q: q[0])) for f in filas]

def _lineas(t):
    return [re.sub(r"\s+", " ", x).strip() for x in t.splitlines() if x.strip()]

def vistas_texto_pdf(pdf):
    crudo = []
    filas = []
    with fitz.open(pdf) as d:
        for p in d:
            crudo.append(p.get_text("text") or "")
            filas += filas_visuales([(w[0], w[1], w[2], w[3], w[4]) for w in p.get_text("words")])
    return [
        ("TEXTO", _lineas("\n".join(crudo))),
        ("TEXTO_FILAS", [re.sub(r"\s+", " ", x).strip() for x in filas if x.strip()]),
    ]

def datos_nombre(nombre):
    s = Path(nombre).stem
    m = re.search(r"(?<![A-Za-z0-9])(?:FCA|FC|FACTURA|NCA|NC|NDA|ND)[\s_]*(\d{1,5})\s*[-_ ]\s*(\d{1,8})(?!\d)", s, re.I)
    if not m:
        return None, None
    p, n = int(m.group(1)), int(m.group(2))
    return (p, n) if p >= 1 and n >= 1 else (None, None)

def _letra(t):
    return t[-1] if t and len(t) >= 2 and t[-2] == " " and t[-1] in "ABCM" else ""

def tipo_por_recuadro(pdf):
    try:
        with fitz.open(pdf) as d:
            for pag in list(d)[:2]:
                pal = [(w[0], w[1], w[2], w[3], sin_acentos(str(w[4])).strip()) for w in pag.get_text("words") if str(w[4]).strip()]
                for c in pal:
                    if not re.fullmatch(r"COD(?:IGO)?\.?", c[4], re.I) or c[1] > pag.rect.height * 0.35:
                        continue
                    cx, cy = (c[0] + c[2]) / 2, (c[1] + c[3]) / 2
                    cod = None
                    for w in pal:
                        wx, wy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
                        derecha = abs(wy - cy) <= 6 and 0 <= w[0] - c[2] <= 40
                        debajo = 4 <= wy - cy <= 22 and abs(wx - cx) <= 40
                        mm = re.fullmatch(r"\D{0,5}?(\d{1,4})", w[4])
                        if mm and (derecha or debajo) and int(mm.group(1)) in TIPOS:
                            cod = int(mm.group(1))
                            break
                    letra = ""
                    for w in pal:
                        wx, wy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
                        if (re.fullmatch(r"[|\[\](){}_.\-P]{0,2}([ABCM])[|\])}_.]?", w[4])
                                and w[1] - 2 <= c[1] and wy < cy and cy - wy <= 50 and abs(wx - cx) <= 45):
                            letra = re.fullmatch(r"[|\[\](){}_.\-P]{0,2}([ABCM])[|\])}_.]?", w[4]).group(1)
                            break
                    if cod and (not letra or _letra(TIPOS[cod]) == letra):
                        return cod, ""
                    if letra:
                        zona = sorted((w for w in pal if c[1] - 60 <= w[1] <= c[3] + 20), key=lambda w: (round(w[1] / 6), w[0]))
                        txt = " ".join(w[4] for w in zona)
                        mejor = None
                        for pat, base in ((r"NOTA\s+DE\s+CREDITO", "NC"), (r"NOTA\s+DE\s+DEBITO", "ND"), (r" FACTURA ", "FACTURA"),
                                          (r" RECIBO ", "RECIBO"), (r" LIQUIDACION ", "LIQUIDACION")):
                            mt = re.search(pat, txt, re.I)
                            if mt and (mejor is None or mt.start() < mejor[0]):
                                mejor = (mt.start(), base)
                        if mejor:
                            return None, f"{mejor[1]} {letra}"
    except Exception:
        pass
    return None, ""

# =============================================================================
# 5. IMPORTE
# =============================================================================
PATRON_MONTO = r"(?<![\d.,])(\d+(?:[.,]\d{3})*[.,]\d{2})(?![\d.,])"

def a_centavos(v):
    return int(round(v * 100))

def montos_en_lineas(lineas):
    res = []
    for l in lineas:
        for mm in re.finditer(PATRON_MONTO, l):
            v = numero(mm.group(1))
            if v is not None and 0 < v < 10 ** 12:
                res.append(v)
    return res

def candidatos_junto_a_total(lineas):
    cands = []
    for i, l in enumerate(lineas):
        if not re.search(r" TOTAL ", l, re.I):
            continue
        cands += montos_en_lineas(lineas[i:i + 1 + LINEAS_VENTANA_TOTAL])
    return cands

def _cierra_suma(objetivo, valores, k_min, k_max=MAX_TERMINOS_SUMA, tol=TOLERANCIA_CENTAVOS):
    n = len(valores)
    def buscar(inicio, restante, usados):
        if usados >= k_min and abs(restante) <= tol:
            return True
        if usados == k_max or restante <= 0:
            return False
        for i in range(inicio, n):
            v = valores[i]
            if v > restante + tol:
                continue
            if sum(valores[i:i + (k_max - usados)]) + tol < restante:
                return False
            if buscar(i + 1, restante - v, usados + 1):
                return True
        return False
    return buscar(0, objetivo, 0)

def _parece_neto(c, cont, tol=2):
    return any(abs(round(c * r) - k) <= tol for r in ALICUOTAS_IVA for k in cont)

def determinar_importe_total(listas_montos, listas_candidatos, k_min=2):
    cont = {}
    for lista in listas_montos:
        for k, n in Counter(a_centavos(v) for v in lista).items():
            cont[k] = max(cont.get(k, 0), n)
    if not cont:
        return None, "SIN_TEXTO"
    candidatos = sorted({a_centavos(v) for lista in listas_candidatos for v in lista if a_centavos(v) in cont}, reverse=True)
    if not candidatos:
        return None, "SIN_PALABRA_TOTAL"
    for c in candidatos:
        resto = []
        for k, n in cont.items():
            resto += [k] * (n - 1 if k == c else n)
        resto.sort(reverse=True)
        if _cierra_suma(c, resto, k_min):
            if _parece_neto(c, cont):
                return None, "TOTAL_PARECE_SUBTOTAL"
            return c / 100, "SUMA_VERIFICADA"
    return None, "TOTAL_NO_CIERRA"

# =============================================================================
# 6. EXTRACCION DE CAMPOS DESDE EL TEXTO DEL PDF
# =============================================================================
def campos_vacios():
    return {
        "proveedor": "", "cuit": "", "cuit_cabecera": "", "cuits": [], "cuits_emisor": [], "cuits_emp": [],
        "tipo": "", "tipo_codigo": None, "pto": "", "nro": "",
        "pto_nro_cands": [], "textos": [],
        "cae": "", "cae_ok": False,
        "importe": None, "importe_motivo": "",
        "montos": [], "candidatos_total": [],
        "moneda": "", "fecha_emision": None,
        "fuente": {}, "caracteres": 0,
    }

PATRONES_PTO_NRO = (
    (100, "ETIQUETAS",
     r"(?:PUNTO\s+DE\s+VENTA|PTO\.?\s*(?:DE\s*)?VTA\.?)\s*[:#-]?\s*(?<!\d)(\d{1,9})(?!\d)[^\d]{0,40}?"
     r"(?:NRO|N[°ºO]|NUMERO)\.?\s*[:#-]?\s*(\d{1,10})(?!\d)", False),
    (90, "ETIQUETA_PREVIA",
     r"(?:NRO|N[°ºO]|NUMERO|COMP(?:ROBANTE)?)\.?\s*(?:NRO|N[°ºO]|COMP\.?)?\s*[:#-]?\s*(?<!\d)(\d{1,6})\s*[-–/]\s*(\d{1,9})(?![\d-])", True),
    (50, "CANONICO",
     r"(?<![\d-])(\d{4,6})\s*-\s*(\d{8,9})(?![\d-])", True),
)

_RE_EXCLUIR = re.compile(
    r"(?<![A-Z0-9])(?:CTG|CPE|C\.?C\.?P\.?P\.?|CARTA\s+DE\s+PORTE|REMITO|CBU|TURNO|CAE|C\.A\.E\.?)(?![A-Z0-9])", re.I)

def _ctx_excluido(plano, ini, fin):
    ls = plano.rfind(" | ", 0, ini)
    prefijo = plano[ls + 3:ini] if ls >= 0 else plano[:ini]
    if len(prefijo.strip()) <= 6 and ls >= 0:
        ps = plano.rfind(" | ", 0, ls)
        prefijo = plano[ps + 3:ls][-45:] + " " + prefijo
    le = plano.find(" | ", fin)
    sufijo = plano[fin:le if le >= 0 else len(plano)][:25]
    return bool(_RE_EXCLUIR.search(prefijo) or _RE_EXCLUIR.search(sufijo))

def candidatos_pto_nro(lineas, plano):
    cands, vistos = [], set()
    def agregar(p_txt, n_txt, score, origen, pos):
        p, rp = recortar_ceros(p_txt, MAX_DIG_PTO)
        n, rn = recortar_ceros(n_txt, MAX_DIG_NRO)
        if p is None or n is None or int(p) < 1 or int(n) < 1 or (p, n, score) in vistos:
            return
        vistos.add((p, n, score))
        cands.append({"pto": p, "nro": n, "score": score, "origen": origen, "pos": pos,
                      "recortado": rp or rn, "crudo": f"{p_txt}-{n_txt}"})

    for score, origen, pat, revisar in PATRONES_PTO_NRO:
        for m in re.finditer(pat, plano, re.I):
            if revisar and _ctx_excluido(plano, m.start(), m.end()):
                continue
            agregar(m.group(1), m.group(2), score, origen, m.start())
    for i, l in enumerate(lineas[:-1]):
        if re.search(r"PUNTO\s+DE\s+VENTA|PTO\.?\s*(?:DE\s*)?VTA", l, re.I) and re.search(r"COMP|NRO|N[°ºO]|NUMERO", l, re.I):
            mm = re.match(r"^\D{0,12}(\d{1,9})\s+(\d{1,10})(?!\d)", lineas[i + 1])
            if mm:
                agregar(mm.group(1), mm.group(2), 80, "TABLA", i)
    return cands

def elegir_candidato(cands):
    return min(cands, key=lambda c: (-c["score"], c.get("seq", 0), c["pos"])) if cands else None

def buscar_anclado(textos, p, n):
    for t in textos:
        t = sin_acentos(t)
        for m in re.finditer(rf"(?<![\d-])(0*{p})\s*[-–/]\s*(0*{n})(?![\d-])", t):
            if len(m.group(1)) <= MAX_DIG_PTO and len(m.group(2)) <= MAX_DIG_NRO:
                return m.group(1), m.group(2)
        mp = re.search(rf"(?:PUNTO\s+DE\s+VENTA|PTO\.?\s*(?:DE\s*)?VTA\.?)\s*[:#-]?\s*(?<!\d)(0*{p})(?!\d)", t, re.I)
        mn = re.search(rf"(?:NRO|N[°ºO]|NUMERO)\.?\s*[:#-]?\s*(?<!\d)(0*{n})(?!\d)", t, re.I)
        if mp and mn and len(mp.group(1)) <= MAX_DIG_PTO and len(mn.group(1)) <= MAX_DIG_NRO:
            return mp.group(1), mn.group(1)
    return None

def resolver_pto_nro(pl, o, nombre_pn):
    obs = []
    qp, qn = a_entero(pl.get("ptoVta")), a_entero(pl.get("nroCmp"))
    qr_ok = qp is not None and qn is not None and 1 <= qp <= 10 ** MAX_DIG_PTO - 1 and 1 <= qn <= 10 ** MAX_DIG_NRO - 1
    cands = o.get("pto_nro_cands", [])
    np_, nn_ = nombre_pn

    def cruce_nombre(p, n):
        if np_ is not None and (int(p), int(n)) != (np_, nn_):
            obs.append(f"PTO-NRO ({p}-{n}) distinto del nombre de archivo ({np_}-{nn_})")

    def aviso_recorte(c):
        if c.get("recortado"):
            obs.append(f"OCR: se quitaron ceros sobrantes ({c['crudo']} -> {c['pto']}-{c['nro']})")

    if qr_ok:
        for c in sorted(cands, key=lambda c: (-c["score"], c.get("seq", 0), c["pos"])):
            if (int(c["pto"]), int(c["nro"])) == (qp, qn):
                aviso_recorte(c)
                cruce_nombre(c["pto"], c["nro"])
                return c["pto"], c["nro"], "OK_QR_Y_TEXTO", obs, ["QR", "TEXTO"]
        a = buscar_anclado(o.get("textos", []), qp, qn)
        if a:
            cruce_nombre(*a)
            return a[0], a[1], "OK_QR_Y_TEXTO_ANCLADO", obs, ["QR", "TEXTO"]
        pto, nro = rellenar(qp, ANCHO_PTO_DEFECTO), rellenar(qn, ANCHO_NRO_DEFECTO)
        obs.append(f"PTO-NRO tomado del QR ({qp}-{qn}); no se pudo ver el formato impreso: ceros completados a "
                   f"{ANCHO_PTO_DEFECTO}/{ANCHO_NRO_DEFECTO} digitos, verificar contra la factura")
        m = elegir_candidato(cands)
        if m:
            obs.append(f"el texto indica {m['pto']}-{m['nro']} y no coincide con el QR: se usa el QR")
        cruce_nombre(pto, nro)
        return pto, nro, "QR_SIN_FORMATO_IMPRESO", obs, ["QR"]

    if cands:
        m = elegir_candidato(cands)
        aviso_recorte(m)
        otros = {(int(c["pto"]), int(c["nro"])) for c in cands if c["score"] >= 80}
        if len(otros) > 1:
            obs.append("hay mas de un PTO-NRO con etiqueta en el PDF (" +
                       ", ".join(f"{p}-{n}" for p, n in sorted(otros)) + "): verificar")
        if len(m["nro"]) != MAX_DIG_NRO or len(m["pto"]) not in (4, 5):
            obs.append(f"PTO-NRO con formato inusual ({m['pto']}-{m['nro']}): verificar contra la factura")
        estado = "SOLO_TEXTO"
        if np_ is not None and (int(m["pto"]), int(m["nro"])) != (np_, nn_):
            estado = "SOLO_TEXTO_DISCREPA_NOMBRE"
            cruce_nombre(m["pto"], m["nro"])
        return m["pto"], m["nro"], estado, obs, ["TEXTO"]

    if np_ is not None:
        obs.append(f"PTO-NRO tomado del NOMBRE de archivo ({np_}-{nn_}): ceros completados a "
                   f"{ANCHO_PTO_DEFECTO}/{ANCHO_NRO_DEFECTO}, verificar contra la factura")
        return rellenar(np_, ANCHO_PTO_DEFECTO), rellenar(nn_, ANCHO_NRO_DEFECTO), "SOLO_NOMBRE_ARCHIVO", obs, ["NOMBRE_ARCHIVO"]
    return "NO_DETECTADO", "NO_DETECTADO", "NO_DETECTADO", obs, []

def campos_texto(lineas_o, cae_esp="", emp_cuits=frozenset(), emp_nombres=frozenset()):
    out = campos_vacios()
    if not lineas_o:
        return out
    lineas = [sin_acentos(l) for l in lineas_o]
    plano_o = " | ".join(lineas_o)
    plano = " | ".join(lineas)

    pats = [
        (r"(?i:NOTA\s+DE\s+CREDITO| N\.?\s*C\.? ).{0,45}? ([ABCM]) ", "NC"),
        (r"(?i:NOTA\s+DE\s+DEBITO| N\.?\s*D\.? ).{0,45}? ([ABCM]) ", "ND"),
        (r"(?i:FACTURA| FCA? |COMPROBANTE\s+TIPO\s+FACTURA).{0,45}? ([ABCM]) ", "FACTURA"),
        (r"(?i:RECIBO).{0,45}? ([ABCM]) ", "RECIBO"),
        (r"(?i:LIQUIDACION).{0,45}? ([ABCM]) ", "LIQUIDACION"),
    ]
    mejor = None
    for pat, base in pats:
        m = re.search(pat, plano)
        if m and (mejor is None or m.start() < mejor[0]):
            mejor = (m.start(), f"{base} {m.group(1)}")
    if mejor:
        out["tipo"] = mejor[1]

    cab = " | ".join(lineas[:40])
    patrones_codigo = (
        r"(?:COD(?:IGO)?|TIPO)\s*(?:DE\s*)?(?:COMPROBANTE|CMP)\s*[:#N°º.*-]*\s*0*(\d{1,3})(?!\d)",
        r"COD(?:IGO)?\.?\s*(?:N(?:RO)?\s*)?[°º*.#:-]*\s*0*(\d{1,3})(?!\d)",
    )
    for pc in patrones_codigo:
        mc = re.search(pc, cab, re.I)
        if not mc:
            continue
        cod = int(mc.group(1))
        if cod not in TIPOS:
            continue
        l1 = _letra(TIPOS[cod])
        l2 = _letra(out["tipo"])
        if not l1 or not l2 or l1 == l2:
            out["tipo_codigo"] = cod
            out["tipo"] = TIPOS[cod]
            break

    out["pto_nro_cands"] = candidatos_pto_nro(lineas, plano)
    mejor_pn = elegir_candidato(out["pto_nro_cands"])
    if mejor_pn:
        out["pto"], out["nro"] = mejor_pn["pto"], mejor_pn["nro"]

    m = re.search(r"C\s*\.?\s*A\s*\.?\s*E\s*(?:NRO|N[°ºO]|NUMERO)?\.?\s*[:#-]?\s*(\d(?:[ .-]?\d){13})", plano, re.I)
    if m:
        out["cae"] = digs(m.group(1))
    if cae_esp and len(cae_esp) == 14:
        out["cae_ok"] = bool(re.search(r"(?<!\d)" + r"[ .\-]?".join(cae_esp) + r"(?!\d)", plano))

    cand = []
    pos_cuits = []
    for i, l in enumerate(lineas):
        for m in re.finditer(r"(?<!\d)(\d{2})[- ]?(\d{8})[- ]?(\d)(?!\d)", l):
            c = "".join(m.groups())
            if c[:2] in PREF_CUIT and cuit_valido(c):
                pos_cuits.append((i, c))
                score = 30 - i
                if re.search(r"CUIT\s*(?:EMISOR)?\s*[:#-]?", l, re.I):
                    score += 100
                if re.search(r"CLIENTE|RECEPTOR|CUIT\s+DESTINATARIO", l, re.I):
                    score -= 200
                cand.append((score, c))
                if c not in out["cuits"]:
                    out["cuits"].append(c)
    mejor_c = {}
    for s, c in cand:
        mejor_c[c] = max(s, mejor_c.get(c, -10 ** 9))
    out["cuits_emp"] = [c for c in out["cuits"] if c in emp_cuits]
    out["cuits_emisor"] = [c for c, s in sorted(mejor_c.items(), key=lambda x: -x[1]) if c not in emp_cuits and s > -100]
    if out["cuits_emisor"]:
        out["cuit"] = out["cuits_emisor"][0]

    fin_cabecera = next((i for i, l in enumerate(lineas)
                         if re.search(r" SENOR(?:ES|\(ES\))? | CLIENTE | RECEPTOR | DESTINATARIO |POR\s+CUENTA\s+Y\s+ORDEN|REMITIMOS\s+A", l, re.I)),
                        len(lineas))
    en_cab = [(i, c) for i, c in pos_cuits
              if i < fin_cabecera and c not in emp_cuits
              and not re.search(r"BRUT|IIBB|C\.?M\.?:|CONVENIO|INSCRIPCION", lineas[i], re.I)]
    con_etiqueta = [(i, c) for i, c in en_cab
                    if re.search(r"C\.?\s?U\.?\s?I\.?\s?T", lineas[i], re.I) or (i > 0 and re.search(r"C\.?\s?U\.?\s?I\.?\s?T\.?\s*[:#N°º*-]*\s*$", lineas[i - 1], re.I))]
    elegido = (con_etiqueta or en_cab or [(None, "")])[0][1]
    out["cuit_cabecera"] = elegido
    if elegido:
        out["cuit"] = elegido

    def es_receptor(n):
        n = norm(n)
        return any(n == e or n.startswith(e + "_") or e.startswith(n + "_") for e in emp_nombres if e and n)

    for p in (r"RAZON\s+SOCIAL\s*[:#-]?\s*([^|]{2,100})", r"DENOMINACION\s*[:#-]?\s*([^|]{2,100})"):
        for m in re.finditer(p, plano, re.I):
            if re.search(r"APELLIDO|NOMBRE|CLIENTE|RECEPTOR|SENOR", plano[max(0, m.start() - 30):m.start()], re.I):
                continue
            sa = plano[m.start(1):m.end(1)]
            k = re.search(r"CUIT|DOMICILIO|CONDICION|INGRESOS", sa, re.I)
            nom = plano_o[m.start(1):m.start(1) + (k.start() if k else len(sa))].strip(" :-")
            if len(nom) >= 3 and not es_receptor(nom):
                out["proveedor"] = nom
                break
        if out["proveedor"]:
            break
    if not out["proveedor"]:
        for lo, l in zip(lineas_o[:15], lineas[:15]):
            if (len(l) <= 70 and ":" not in l and len(re.sub(r"\D", "", l)) < 5
                    and re.search(r" (?:S\.?\s?A\.?\s?S?|S\.?\s?R\.?\s?L|S\.?\s?C\.?\s?A|S\.?\s?A\.?\s?I\.?\s?C\.?[A-Z.]*|S\.?\s?H|LTDA|COOPERATIVA(?:\s+\w+)?)\.?$", l, re.I)
                    and not es_receptor(lo)):
                out["proveedor"] = lo.strip(" :-")
                break

    if re.search(r" (?:USD|DOLARES?|MONEDA\s*[:#-]?\s*DOL) |U\s*\$\s*S|US\$", plano, re.I):
        out["moneda"] = "USD"
    elif re.search(r" (?:ARS|PESOS?|MONEDA\s*[:#-]?\s*PES) |\$", plano, re.I):
        out["moneda"] = "ARS"

    out["montos"] = montos_en_lineas(lineas)
    out["candidatos_total"] = candidatos_junto_a_total(lineas)

    for pf in (r"FECHA\s+(?:DE\s+)?EMISION\s*[:#-]?\s*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})",
               r"EMISION\s*[:#-]?\s*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})",
               r"FECHA\s*[:#-]\s*(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})"):
        m = re.search(pf, plano, re.I)
        if m and fecha(m.group(1)):
            out["fecha_emision"] = fecha(m.group(1))
            break
    if not out["fecha_emision"]:
        for l in lineas[:30]:
            if re.search(r"VTO|VENC|DESDE|HASTA|INICIO|PERIODO|CAE|VALIDO", l, re.I):
                continue
            m = re.search(r"(?<![\d/.\-])(\d{1,2}[/.\-]\d{1,2}[/.\-]\d{2,4})(?![\d/.\-])", l)
            if m and fecha(m.group(1)):
                out["fecha_emision"] = fecha(m.group(1))
                break
    return out

def fusionar(dst, src, vista):
    for c in src["pto_nro_cands"]:
        if not any((c["pto"], c["nro"], c["score"]) == (x["pto"], x["nro"], x["score"]) for x in dst["pto_nro_cands"]):
            dst["pto_nro_cands"].append(dict(c, seq=len(dst["pto_nro_cands"])))
    for k, v in src.items():
        if k in ("fuente", "caracteres", "pto", "nro", "pto_nro_cands", "textos", "importe", "importe_motivo", "montos", "candidatos_total"):
            continue
        if k in ("cuits", "cuits_emisor", "cuits_emp"):
            dst[k] += [x for x in v if x not in dst[k]]
        elif k == "cae_ok":
            dst[k] = dst[k] or v
        elif dst[k] in (None, "") and v not in (None, ""):
            dst[k] = v
            dst["fuente"][k] = vista

def extraer_ocr(pdf, cae_esp="", m=None):
    emp_c = set(m.emp) if m else set()
    emp_n = {norm(v) for v in m.emp.values()} if m else set()
    out = campos_vacios()
    try:
        vistas = vistas_texto_pdf(pdf)
    except Exception:
        vistas = []

    por_vista = []
    for nombre, lin in vistas:
        r = campos_texto(lin, cae_esp, emp_c, emp_n)
        por_vista.append(r)
        fusionar(out, r, nombre)
    out["caracteres"] = sum(len(x.replace(" ", "")) for x in vistas[0][1]) if vistas else 0
    mejor_pn = elegir_candidato(out["pto_nro_cands"])
    out["pto"], out["nro"] = (mejor_pn["pto"], mejor_pn["nro"]) if mejor_pn else ("", "")
    out["textos"] = [" | ".join(lin) for _, lin in vistas]

    if not out["tipo_codigo"]:
        cod_r, tipo_r = tipo_por_recuadro(pdf)
        l_prev = _letra(out["tipo"])
        if cod_r and (not l_prev or l_prev == _letra(TIPOS[cod_r])):
            out["tipo_codigo"], out["tipo"] = cod_r, TIPOS[cod_r]
            out["fuente"]["tipo_codigo"] = "RECUADRO"
        elif tipo_r and not out["tipo"]:
            out["tipo"] = tipo_r
            out["fuente"]["tipo"] = "RECUADRO"
    k_min = 1 if _letra(out["tipo"]) in ("B", "C") else 2
    imp, motivo = determinar_importe_total([r["montos"] for r in por_vista], [r["candidatos_total"] for r in por_vista], k_min)
    out["importe"], out["importe_motivo"] = imp, motivo
    if imp is not None:
        out["fuente"]["importe"] = motivo
        if out["tipo"].startswith("NC "):
            out["importe"] = -abs(imp)
    return out

# =============================================================================
# 7. MAESTROS Y LISTADO DE FACTURAS (CORREGIDO)
# =============================================================================
@dataclass
class Maestro:
    carpeta: Path

    def __post_init__(self):
        self.prov = {}
        self.emp = {}
        self.fechas = {}
        self.prov_por_nombre = {}
        self.tablas = 0

    def col(self, hs, g):
        return next((h for h in hs if h in ALIASES[g]), None)

    def cargar(self):
        for f in sorted(self.carpeta.rglob("*.xlsx")):
            if f.name.startswith("~$"):
                continue
            wb = load_workbook(f, read_only=True, data_only=True)
            for ws in wb.worksheets:
                it = ws.iter_rows(values_only=True)
                h0 = next(it, None)
                if not h0:
                    continue
                hs = [norm(x) or f"COL_{i}" for i, x in enumerate(h0)]
                self.tablas += 1
                es_maestro_empresas = "EMPRESA" in norm(f.stem)
                cc = self.col(hs, "cuit_receptor" if es_maestro_empresas else "cuit_proveedor")
                cp = self.col(hs, "proveedor")
                ce = self.col(hs, "empresa")
                ca = self.col(hs, "archivo")
                cr = self.col(hs, "fecha_recepcion")
                cm = self.col(hs, "fecha_modificacion")
                for row in it:
                    d = {hs[i]: (row[i] if i < len(row) else None) for i in range(len(hs))}
                    c = digs(d.get(cc)) if cc else ""
                    if c and cp and d.get(cp) and not es_maestro_empresas:
                        nombre_proveedor = str(d[cp]).strip()
                        self.prov[c] = nombre_proveedor
                        self.prov_por_nombre[norm(nombre_proveedor)] = c
                    if c and ce and es_maestro_empresas and d.get(ce):
                        self.emp[c] = str(d[ce]).strip()
                    if ca and d.get(ca):
                        self.fechas[norm(Path(str(d[ca])).name)] = {
                            "recepcion": fecha(d.get(cr)) if cr else None,
                            "modificacion": fecha(d.get(cm)) if cm else None,
                        }
            wb.close()

def raiz():
    s = Path(__file__).resolve().parent
    for p in (s, *s.parents):
        if norm(p.name) == norm("98. AUTOMATIZACION"):
            return p
    return s.parent

def buscar(nombre, carpetas):
    o = norm(Path(nombre).stem)
    for c in carpetas:
        if c.exists():
            for f in c.rglob("*.xlsx"):
                if norm(f.stem) == o:
                    return f
    raise FileNotFoundError(nombre)

def listado(p):
    wb = load_workbook(p, read_only=True, data_only=True)
    ws = wb["CAE Extraidos"] if "CAE Extraidos" in wb.sheetnames else wb.active
    it = ws.iter_rows(values_only=True)
    hs = [norm(x) for x in next(it)]
    por_cae, conteo = {}, Counter()
    for r in it:
        d = {hs[i]: (r[i] if i < len(r) else None) for i in range(len(hs))}
        c = digs(d.get("CAE"))
        if len(c) != 14:
            continue
        conteo[c] += 1
        por_cae.setdefault(c, {"cae": c})
    wb.close()
    return list(por_cae.values()), {c: n for c, n in conteo.items() if n > 1}

def extraer_caes_objetivo(texto, objetivos):
    """Busca CAE objetivo en texto OCR ignorando espacios, puntos o guiones."""
    encontrados = set()
    texto_limpio = re.sub(r"\D", "", texto or "")
    for c in objetivos:
        if c in texto_limpio:
            encontrados.add(c)
    return encontrados

def indice_cae(carpeta, caes_objetivo):
    """Lee cada PDF de forma recursiva y arma el indice CAE -> lista de PDF."""
    objetivos = set(caes_objetivo)
    indice = {c: [] for c in objetivos}
    errores = []
    pdfs = sorted(carpeta.rglob("*.pdf"))
    for pdf in pdfs:
        try:
            with fitz.open(pdf) as doc:
                texto = "".join(pag.get_text("text") or "" for pag in doc)
            for c in extraer_caes_objetivo(texto, objetivos):
                indice[c].append(pdf)
        except Exception as e:
            errores.append({"archivo": pdf.name, "ruta": str(pdf), "error": str(e)})
    return indice, pdfs, errores

def resolver_por_cae(cae, indice):
    candidatos = sorted(indice.get(cae, []), key=lambda p: (len(str(p)), str(p).lower()))
    return (candidatos[0] if candidatos else None), candidatos

# =============================================================================
# 8. LECTURA DEL QR FISCAL (ARCA / AFIP)
# =============================================================================
def payload_url(url):
    try:
        url_limpia = url.strip()
        u = urlparse(url_limpia)
        dominio = u.netloc.lower().split(":")[0]
        ruta = u.path

        dominios_autorizados = ("arca.gob.ar", "afip.gob.ar", "servicioscf.afip.gob.ar")
        if not any(dominio == d or dominio.endswith("." + d) for d in dominios_autorizados):
            return None

        if "/fe/qr" not in ruta.lower():
            return None

        p = parse_qs(u.query).get("p", [None])[0]

        if not p:
            m = re.search(r"/U?p=([^/?#&]+)", ruta, re.IGNORECASE)
            if m:
                p = m.group(1)

        if not p:
            return None

        p += "=" * (-len(p) % 4)
        datos = json.loads(base64.urlsafe_b64decode(p).decode("utf-8-sig"))

        requeridos = {"cuit", "ptoVta", "tipoCmp", "nroCmp", "importe", "moneda", "codAut"}
        if not requeridos.issubset(datos):
            return None

        return datos
    except Exception:
        return None

def qrs(img):
    res = []
    gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    vars = [img, gris, cv2.equalizeHist(gris)]
    for modo in (cv2.THRESH_BINARY, cv2.THRESH_BINARY_INV):
        _, th = cv2.threshold(gris, 0, 255, modo | cv2.THRESH_OTSU)
        vars.append(th)
    for v in vars:
        try:
            for b in zxingcpp.read_barcodes(v, formats=zxingcpp.BarcodeFormat.QRCode):
                if b.text:
                    res.append(b.text)
        except Exception:
            pass
    return list(dict.fromkeys(x.strip() for x in res if x and x.strip()))

def localizar_qr_en_imagen(img):
    try:
        for b in zxingcpp.read_barcodes(img, formats=zxingcpp.BarcodeFormat.QRCode):
            pos = b.position
            xs = [int(pt.x) for pt in (pos.top_left, pos.top_right, pos.bottom_right, pos.bottom_left)]
            ys = [int(pt.y) for pt in (pos.top_left, pos.top_right, pos.bottom_right, pos.bottom_left)]
            x0, y0, x1, y1 = min(xs), min(ys), max(xs) + 1, max(ys) + 1
            if x1 > x0 and y1 > y0:
                return (x0, y0, x1, y1)
    except Exception:
        pass
    try:
        ok, puntos = cv2.QRCodeDetector().detect(img)
        if ok and puntos is not None:
            pts = np.asarray(puntos).reshape(-1, 2)
            x0, y0 = np.floor(pts.min(axis=0)).astype(int)
            x1, y1 = np.ceil(pts.max(axis=0)).astype(int)
            if x1 > x0 and y1 > y0:
                return (x0, y0, x1, y1)
    except Exception:
        pass
    return None

def _probar_qr(img, cae, r, pagina, metodo, dpi=""):
    for u in qrs(img):
        x = payload_url(u)
        if not x:
            continue
        r.update(payload=x, pagina=pagina, metodo=metodo, dpi=dpi, motivo="QR_VALIDO")
        return True
    return False

def _pix_a_bgr(pix):
    arr = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)
    return cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)

def qr(pdf, cae):
    r = {"payload": None, "pagina": "", "metodo": "", "dpi": "", "motivo": "QR_NO_LEIDO"}
    try:
        with fitz.open(pdf) as d:
            total = len(d)
            for i, p in enumerate(d):
                for im in p.get_images(full=True):
                    try:
                        a = np.frombuffer(d.extract_image(im[0])["image"], np.uint8)
                        img = cv2.imdecode(a, cv2.IMREAD_COLOR)
                        if img is None:
                            continue
                        if _probar_qr(img, cae, r, i + 1, "IMAGEN_EMBEBIDA"):
                            return r
                    except Exception:
                        pass
                bbox = None
                for esc in (1, 2):
                    try:
                        pix0 = p.get_pixmap(matrix=fitz.Matrix(esc, esc), alpha=False)
                        bbox = localizar_qr_en_imagen(_pix_a_bgr(pix0))
                    except Exception:
                        bbox = None
                    if bbox:
                        break
                if bbox:
                    x0, y0, x1, y1 = bbox
                    margen = max(8, int(max(x1 - x0, y1 - y0) * 0.15))
                    x0 = max(0, x0 - margen)
                    y0 = max(0, y0 - margen)
                    x1 = min(pix0.width, x1 + margen)
                    y1 = min(pix0.height, y1 + margen)
                    sx = p.rect.width / pix0.width
                    sy = p.rect.height / pix0.height
                    rect_qr = fitz.Rect(x0 * sx, y0 * sy, x1 * sx, y1 * sy) & p.rect
                    if not rect_qr.is_empty:
                        pix = p.get_pixmap(dpi=QR_DPI, clip=rect_qr, alpha=False)
                        if _probar_qr(_pix_a_bgr(pix), cae, r, i + 1, "QR_RECORTADO", QR_DPI):
                            return r
                if QR_PAGINA_COMPLETA and (i < 3 or i == total - 1):
                    try:
                        pix = p.get_pixmap(dpi=QR_DPI, alpha=False)
                        if _probar_qr(_pix_a_bgr(pix), cae, r, i + 1, "PAGINA_COMPLETA", QR_DPI):
                            return r
                    except Exception:
                        pass
        return r
    except Exception as e:
        r["motivo"] = f"ERROR_PDF:{e}"
        return r

# =============================================================================
# 9. ARMADO DE CADA FILA DEL RESULTADO
# =============================================================================
def empresa(pdf, payload, m, cuits_texto=()):
    rec = digs(payload.get("nroDocRec"))
    if rec in m.emp:
        return m.emp[rec]
    for c in cuits_texto:
        if c in m.emp:
            return m.emp[c]
    t = norm(str(pdf.parent) + " " + pdf.name)
    for tok, val in (("MASAS", "MASAS"), ("BAT", "BATATALES"), ("SEMILLERO", "SEMILLERO"),
                     ("_SC", "SEMILLERO"), ("GER", "GERMINAGRO"), ("AYERZA", "AYERZA ALEJANDRO")):
        if tok in t:
            return val
    return "NO_IDENTIFICADA"

def fila(pdf, cae_esp, q, m):
    pl = q.get("payload") or {}
    obs = []
    nombres_empresas = {norm(v) for v in m.emp.values()}

    # --- Fechas de modificacion y vencimiento desde archivo/maestros ---
    df = m.fechas.get(norm(pdf.name), {})
    fm = df.get("modificacion")
    vto_nombre = None
    mn = re.search(r" VTO\s*(\d{1,2})[-_. /](\d{1,2})(?:[-_. /](\d{4}|\d{2}))?(?!\d)", norm(pdf.stem).replace("_", " "), re.I)
    if mn:
        referencia_titulo = fm or date.today()
        y = int(mn.group(3)) if mn.group(3) else referencia_titulo.year
        y = y + 2000 if y < 100 else y
        try:
            vto_nombre = date(y, int(mn.group(2)), int(mn.group(1)))
        except Exception:
            vto_nombre = None
    if vto_nombre is not None:
        base = vto_nombre
        venc = vto_nombre
        dias = "NOMBRE_ARCHIVO"
    elif fm is not None:
        base = fm
        venc = fm + timedelta(days=28)
        dias = "MODIFICACION_MAS_28_DIAS"
    else:
        base = "NO_ENCONTRADO_EN_MAESTRO_FECHAS"
        venc = "NO_CALCULADO"
        dias = "SIN_FECHA"

    # =========================================================================
    # CASO 1: EL QR PUDO SER PROCESADO (Se toman EXCLUSIVAMENTE los datos del QR)
    # =========================================================================
    if pl:
        # 1. Punto de Venta y Numero de Comprobante
        qp = a_entero(pl.get("ptoVta"))
        qn = a_entero(pl.get("nroCmp"))
        pto = rellenar(qp, ANCHO_PTO_DEFECTO) if qp is not None else "NO_DETECTADO"
        nro = rellenar(qn, ANCHO_NRO_DEFECTO) if qn is not None else "NO_DETECTADO"
        estado_pn = "OK_QR"

        # 2. CUIT Emisor y Proveedor
        c_qr = digs(pl.get("cuit"))
        c = c_qr if c_qr else "NO_DETECTADO"
        cuit_factura = c_qr if c_qr else "NO_DETECTADO"
        prov = m.prov.get(c) if c != "NO_DETECTADO" and c not in m.emp else None
        if not prov:
            prov = "NO_ENCONTRADO_EN_MAESTRO"
        if c in m.emp or norm(prov) in nombres_empresas:
            obs.append("ERROR DE INTEGRIDAD: empresa receptora y proveedor coinciden")
            c = "NO_DETECTADO"
            prov = "NO_ENCONTRADO_EN_MAESTRO"

        # 3. Empresa Receptora
        emp_rec = empresa(pdf, pl, m, ())

        # 4. Tipo de Comprobante
        cod_qr = a_entero(pl.get("tipoCmp"))
        tc = TIPOS.get(cod_qr) if cod_qr in TIPOS else "NO_DETECTADO"

        # 5. Importe Total
        pi = pl.get("importe")
        imp = float(pi) if isinstance(pi, (int, float)) else numero(pi)
        if tc.startswith("NC ") and imp is not None:
            imp = -abs(imp)

        # 6. Moneda
        mon_raw = norm(pl.get("moneda"))
        if mon_raw in {"PES", "ARS", "PESO", "PESOS"}:
            mon = "ARS"
        elif mon_raw in {"DOL", "USD", "DOLAR", "DOLARES"}:
            mon = "USD"
        else:
            mon = mon_raw or "NO_DETECTADA"

        # 7. Fecha de Emision
        fe = fecha(pl.get("fecha")) or "NO_DETECTADA"

        # 8. CAE
        cae_final = digs(pl.get("codAut")) or cae_esp

        # 9. Fuente de Datos
        fuente_datos = "QR"

    # =========================================================================
    # CASO 2: EL QR NO PUDO SER PROCESADO (Se recurre al texto OCR del PDF)
    # =========================================================================
    else:
        o = extraer_ocr(pdf, cae_esp, m)
        nombre_pn = datos_nombre(pdf.name)
        orig = []
        tag = lambda campo: "TEXTO"

        pto, nro, estado_pn, obs_pn, fuentes_pn = resolver_pto_nro({}, o, nombre_pn)
        obs += obs_pn
        orig += fuentes_pn

        c_qr = digs(pl.get("cuit"))
        candidatos = []
        pares_candidatos = [(c_qr, "QR"), (o["cuit_cabecera"], "TEXTO")]
        pares_candidatos += [(x, "TEXTO") for x in o["cuits_emisor"]]
        for candidato, fuente in pares_candidatos:
            candidato = digs(candidato)
            if (candidato and candidato in m.prov and candidato not in m.emp
                    and candidato not in [x[0] for x in candidatos]):
                candidatos.append((candidato, fuente))
        c = candidatos[0][0] if candidatos else "NO_DETECTADO"
        if candidatos:
            orig.append(candidatos[0][1])
        elif c_qr and c_qr in m.emp:
            obs.append("CUIT del QR descartado: corresponde a la empresa receptora")
        elif o["cuits"]:
            obs.append("Ningun CUIT legible del PDF coincide con el maestro de proveedores")
        else:
            obs.append("No se encontro un CUIT emisor legible para validar contra el maestro de proveedores")

        if c != "NO_DETECTADO":
            cuit_factura = c
        else:
            cuit_factura = "NO_DETECTADO"
            pares_factura = [(c_qr, "QR"), (o["cuit_cabecera"], "TEXTO")] + [(x, "TEXTO") for x in o["cuits_emisor"]]
            for candidato, fuente in pares_factura:
                candidato = digs(candidato)
                if candidato and cuit_valido(candidato) and candidato not in m.emp:
                    cuit_factura = candidato
                    break

        tcod = TIPOS.get(o["tipo_codigo"])
        tc = tcod or o["tipo"] or "NO_DETECTADO"
        if tc != "NO_DETECTADO":
            orig.append(tag("tipo_codigo" if tcod else "tipo"))

        imp = o["importe"]
        if imp is None:
            obs.append("IMPORTE sin verificar (" + MOTIVOS_IMPORTE.get(o["importe_motivo"], "sin datos") + "): cargar manualmente")
        if tc.startswith("NC ") and imp is not None:
            imp = -abs(imp)

        mon = o["moneda"] or "NO_DETECTADA"
        if o["moneda"]:
            orig.append(tag("moneda"))
        elif o["caracteres"] >= MIN_CARACTERES_TEXTO:
            mon = "ARS"
            obs.append("MONEDA no especificada en el PDF: se asume ARS")

        emp_rec = empresa(pdf, pl, m, o["cuits_emp"])
        prov = m.prov.get(c) if c != "NO_DETECTADO" and c not in m.emp else None
        if not prov:
            prov = "NO_ENCONTRADO_EN_MAESTRO"
        if c in m.emp or norm(prov) in nombres_empresas:
            obs.append("ERROR DE INTEGRIDAD: empresa receptora y proveedor coinciden")
            c = "NO_DETECTADO"
            prov = "NO_ENCONTRADO_EN_MAESTRO"

        fe = o["fecha_emision"] or "NO_DETECTADA"
        cae_final = o["cae"] or cae_esp

        if o["caracteres"] >= MIN_CARACTERES_TEXTO:
            if not o["cae_ok"]:
                obs.append("el CAE del listado no figura en el texto del PDF: verificar que sea el PDF correcto")
        obs.insert(0, "QR: " + str(q.get("motivo") or "QR_NO_LEIDO"))
        if o["caracteres"] < MIN_CARACTERES_TEXTO:
            obs.append("el PDF no tiene texto legible (sin capa de texto): no hay de donde extraer datos")

        fuente_datos = "+".join(x for x in ("TEXTO", "NOMBRE_ARCHIVO") if x in orig) or "SIN_DATOS"

    reg = {
        "EMPRESA_RECEPTORA": emp_rec, "PROVEEDOR": prov, "TIPO_COMPROBANTE": tc,
        "PUNTO_VENTA": pto, "NUMERO_COMPROBANTE": nro,
        "FECHA_REFERENCIA_VENCIMIENTO": base, "FECHA_VENCIMIENTO_CALCULADA": venc, "ORIGEN_VENCIMIENTO": dias,
        "IMPORTE_TOTAL": imp if imp is not None else "NO_DETECTADO",
        "MONEDA": mon, "CUIT_EMISOR_MAESTRO_DE_PROVEEDORES": c, "CUIT_EMISOR_FACTURA": cuit_factura, "CAE": cae_final, "NOMBRE_ARCHIVO": pdf.name, "RUTA_COMPLETA": str(pdf),
        "FECHA_MODIFICACION_ARCHIVO": fm or "NO_ENCONTRADO_EN_MAESTRO_FECHAS",
        "ESTADO_QR": "QR_PROCESADO" if pl else "QR_NO_PROCESADO",
        "FECHA_EMISION": fe,
        "VALIDACION_PTO_NRO": estado_pn,
    }
    falt = campos_faltantes(reg)
    if falt:
        obs.append("FALTAN: " + ", ".join(falt))
    reg["FUENTE_DATOS"] = fuente_datos
    reg["OBSERVACIONES"] = " | ".join(obs)
    return reg

# =============================================================================
# 10. GUARDADO DEL EXCEL
# =============================================================================
def guardar(regs, salida):
    wb = Workbook()
    ws = wb.active
    ws.title = "RESULTADO"
    ws.append(COLUMNAS)
    for r in regs:
        ws.append([r.get(c, "") for c in COLUMNAS])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i in range(2, ws.max_row + 1):
        for nom in ("FECHA_REFERENCIA_VENCIMIENTO", "FECHA_VENCIMIENTO_CALCULADA", "FECHA_MODIFICACION_ARCHIVO", "FECHA_EMISION"):
            ws.cell(i, COLUMNAS.index(nom) + 1).number_format = "dd/mm/yyyy"
        ws.cell(i, COLUMNAS.index("IMPORTE_TOTAL") + 1).number_format = '#,##0.00'
        for nom in ("PUNTO_VENTA", "NUMERO_COMPROBANTE", "CUIT_EMISOR_MAESTRO_DE_PROVEEDORES", "CUIT_EMISOR_FACTURA", "CAE"):
            ws.cell(i, COLUMNAS.index(nom) + 1).number_format = "@"
        ws.cell(i, COLUMNAS.index("FECHA_HORA_PROCESAMIENTO") + 1).number_format = "dd/mm/yyyy hh:mm:ss"
    salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(salida)

def guardar_log_xlsx(detalles, resumen, salida):
    wb = Workbook()
    ws = wb.active
    ws.title = "DETALLE"
    columnas = ["ORDEN", "PORCENTAJE", "CAE", "ESTADO_BUSQUEDA", "ESTADO_QR", "NOMBRE_ARCHIVO",
                "RUTA_COMPLETA", "COINCIDENCIAS_CAE", "TIPO_COMPROBANTE", "PUNTO_VENTA",
                "NUMERO_COMPROBANTE", "IMPORTE_TOTAL", "MONEDA", "CUIT_EMISOR_MAESTRO_DE_PROVEEDORES", "CUIT_EMISOR_FACTURA",
                "FUENTE_DATOS", "OBSERVACIONES"]
    ws.append(columnas)
    for d in detalles:
        ws.append([d.get(c, "") for c in columnas])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i in range(2, ws.max_row + 1):
        ws.cell(i, 2).number_format = "0.00%"
        ws.cell(i, 3).number_format = "@"
        ws.cell(i, 12).number_format = '#,##0.00'

    wr = wb.create_sheet("RESUMEN")
    wr.append(["CAMPO", "VALOR"])
    for k, v in resumen.items():
        wr.append([k, v])
    for c in wr[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for i in range(2, wr.max_row + 1):
        if wr.cell(i, 1).value in ("INICIO", "FIN"):
            wr.cell(i, 2).number_format = "dd/mm/yyyy hh:mm:ss"
        elif wr.cell(i, 1).value == "DURACION_SEGUNDOS":
            wr.cell(i, 2).number_format = "0.00"
    ws.column_dimensions["F"].width = 45
    ws.column_dimensions["G"].width = 80
    ws.column_dimensions["Q"].width = 100
    wr.column_dimensions["A"].width = 32
    wr.column_dimensions["B"].width = 70
    salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(salida)

def agregar_resumen_a_resultado(resumen, salida):
    wb = load_workbook(salida)
    if "LOG_PROCESO" in wb.sheetnames:
        del wb["LOG_PROCESO"]
    ws = wb.create_sheet("LOG_PROCESO")
    ws.append(["CAMPO", "VALOR"])
    for k, v in resumen.items():
        ws.append([k, v])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center")
    for i in range(2, ws.max_row + 1):
        ws.cell(i, 2).alignment = Alignment(horizontal="left")
        if ws.cell(i, 1).value == "DURACION_SEGUNDOS":
            ws.cell(i, 2).number_format = "0.00"
        elif ws.cell(i, 1).value in ("DURACION_HH_MM_SS", "HORA_INICIO", "HORA_FIN", "INICIO_COMPLETO", "FIN_COMPLETO", "FECHA_EJECUCION"):
            ws.cell(i, 2).number_format = "@"
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 70
    wb.save(salida)

# =============================================================================
# 11. VALIDACION DE ARCHIVOS ABIERTOS
# =============================================================================
def archivo_bloqueado(path: Path) -> bool:
    if not path.exists():
        return False
    try:
        import msvcrt
        with open(path, "r+b") as f:
            try:
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
                msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
                return False
            except OSError:
                return True
    except (PermissionError, OSError):
        return True

def validar_archivos_cerrados(archivos):
    return [Path(a) for a in archivos if a and archivo_bloqueado(Path(a))]

# =============================================================================
# 12. PROGRAMA PRINCIPAL
# =============================================================================
def main():
    inicio = datetime.now()
    t0 = time.perf_counter()
    r = raiz()
    cf = r.parent / CARPETA_FACTURAS
    if not cf.exists():
        cf = r / CARPETA_FACTURAS
    logs = r / CARPETA_LOGS
    logs.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(message)s", handlers=[logging.StreamHandler(sys.stdout)], force=True)
    log = logging.getLogger("CAE")
    listado_path = buscar(LISTADO_NOMBRE, [r / CARPETA_INTERMEDIA, r])
    maestros = [f for f in (r / CARPETA_MAESTROS).rglob("*.xlsx") if not f.name.startswith("~$")]
    salida = r / CARPETA_RESULTADO / SALIDA_NOMBRE
    log_xlsx = logs / LOG_NOMBRE

    bloqueados = validar_archivos_cerrados([listado_path, *maestros] + ([salida] if salida.exists() else []) + ([log_xlsx] if log_xlsx.exists() else []))
    print("\n" + "=" * 70)
    print("OBTENCION DE DATOS FISCALES DESDE CAE")
    print("=" * 70)
    print("Validando que los archivos necesarios esten cerrados...")
    if bloqueados:
        print("\nNO SE PUEDE INICIAR EL PROCESO.")
        print("Cerra estos archivos:")
        for a in bloqueados:
            print(f" - {a}")
        print("No se proceso ninguna factura.")
        return 2
    print("Validacion de archivos: OK")

    print("Cargando maestros y CAE unicos...")
    m = Maestro(r / CARPETA_MAESTROS)
    m.cargar()
    ls, repetidos_listado = listado(listado_path)
    total = len(ls)
    print("Indexando PDF por CAE desde la capa OCR...")
    idx, pdfs, errores_pdf = indice_cae(cf, [x["cae"] for x in ls])

    regs, detalles = [], []
    qr_ok = qr_no = no_encontrado = duplicado_pdf = 0
    print(f"CAE unicos a procesar : {total}")
    print(f"CAE repetidos listado : {len(repetidos_listado)}")
    print(f"PDF revisados         : {len(pdfs)}")
    print(f"Maestros cargados     : {m.tablas}")
    print("Busqueda principal    : CAE dentro del texto OCR")
    print("-" * 70)

    for n, it in enumerate(ls, 1):
        cae = it["cae"]
        porcentaje = n / total * 100 if total else 100.0
        pdf, candidatos = resolver_por_cae(cae, idx)
        if not pdf:
            no_encontrado += 1
            qr_no += 1
            log.warning("[%6.2f%%] [%s/%s] CAE_NO_ENCONTRADO | CAE: %s", porcentaje, n, total, cae)
            detalles.append({"ORDEN": n, "PORCENTAJE": porcentaje / 100, "CAE": cae,
                             "ESTADO_BUSQUEDA": "CAE_NO_ENCONTRADO", "ESTADO_QR": "QR_NO_PROCESADO",
                             "COINCIDENCIAS_CAE": 0, "OBSERVACIONES": "CAE no encontrado en la capa OCR de ningun PDF"})
            continue

        estado = "CAE_ENCONTRADO"
        obs_busqueda = ""
        if len(candidatos) > 1:
            duplicado_pdf += 1
            estado = "CAE_DUPLICADO_EN_PDF"
            obs_busqueda = "CAE encontrado en varios PDF; se proceso uno solo: " + " | ".join(str(x) for x in candidatos)

        q = qr(pdf, cae)
        reg = fila(pdf, cae, q, m)
        reg["ESTADO_QR"] = "QR_PROCESADO" if q.get("payload") else "QR_NO_PROCESADO"
        if obs_busqueda:
            reg["OBSERVACIONES"] = " | ".join(x for x in (obs_busqueda, reg["OBSERVACIONES"]) if x)
        regs.append(reg)
        if reg["ESTADO_QR"] == "QR_PROCESADO":
            qr_ok += 1
        else:
            qr_no += 1

        detalles.append({"ORDEN": n, "PORCENTAJE": porcentaje / 100, "CAE": cae, "ESTADO_BUSQUEDA": estado,
                         "ESTADO_QR": reg["ESTADO_QR"], "NOMBRE_ARCHIVO": pdf.name, "RUTA_COMPLETA": str(pdf),
                         "COINCIDENCIAS_CAE": len(candidatos), "TIPO_COMPROBANTE": reg["TIPO_COMPROBANTE"],
                         "PUNTO_VENTA": reg["PUNTO_VENTA"], "NUMERO_COMPROBANTE": reg["NUMERO_COMPROBANTE"],
                         "IMPORTE_TOTAL": reg["IMPORTE_TOTAL"], "MONEDA": reg["MONEDA"], "CUIT_EMISOR_MAESTRO_DE_PROVEEDORES": reg["CUIT_EMISOR_MAESTRO_DE_PROVEEDORES"],
                         "CUIT_EMISOR_FACTURA": reg["CUIT_EMISOR_FACTURA"],
                         "FUENTE_DATOS": reg["FUENTE_DATOS"], "OBSERVACIONES": reg["OBSERVACIONES"]})
        log.info("[%6.2f%%] [%s/%s] %s | CAE: %s | ARCHIVO: %s | Tipo: %s | Pto: %s | Nro: %s | Importe: %s | Moneda: %s | CUIT: %s | Origen: %s%s",
                 porcentaje, n, total, estado, cae, pdf.name, reg["TIPO_COMPROBANTE"], reg["PUNTO_VENTA"],
                 reg["NUMERO_COMPROBANTE"], reg["IMPORTE_TOTAL"], reg["MONEDA"], reg["CUIT_EMISOR_MAESTRO_DE_PROVEEDORES"], reg["FUENTE_DATOS"],
                 " | Obs: " + reg["OBSERVACIONES"] if reg["OBSERVACIONES"] else "")

    for e in errores_pdf:
        detalles.append({"ESTADO_BUSQUEDA": "ERROR_LECTURA_PDF", "NOMBRE_ARCHIVO": e["archivo"],
                         "RUTA_COMPLETA": e["ruta"], "OBSERVACIONES": e["error"]})

    fin = datetime.now()
    fecha_hora_fin_str = fin.strftime("%d/%m/%Y %H:%M:%S")

    for reg in regs:
        reg["FECHA_HORA_PROCESAMIENTO"] = fecha_hora_fin_str

    guardar(regs, salida)
    comp = sum(1 for g in regs if not campos_faltantes(g))
    duracion = time.perf_counter() - t0
    segundos_totales = int(round(duracion))
    horas, resto = divmod(segundos_totales, 3600)
    minutos, segundos = divmod(resto, 60)
    duracion_formateada = f"{horas:02d}:{minutos:02d}:{segundos:02d}"
    resumen = {"FECHA_EJECUCION": inicio.strftime("%d/%m/%Y"),
               "HORA_INICIO": inicio.strftime("%H:%M:%S"),
               "HORA_FIN": fin.strftime("%H:%M:%S"),
               "INICIO_COMPLETO": inicio.strftime("%d/%m/%Y %H:%M:%S"),
               "FIN_COMPLETO": fin.strftime("%d/%m/%Y %H:%M:%S"),
               "DURACION_SEGUNDOS": round(duracion, 2),
               "DURACION_HH_MM_SS": duracion_formateada, "CAE_UNICOS_LISTADO": total,
               "CAE_REPETIDOS_EN_LISTADO": len(repetidos_listado), "PDF_REVISADOS": len(pdfs),
               "PDF_CON_ERROR_LECTURA": len(errores_pdf), "FILAS_RESULTADO": len(regs), "FILAS_COMPLETAS": comp,
               "FILAS_CON_FALTANTES": len(regs) - comp, "CAE_NO_ENCONTRADO": no_encontrado,
               "CAE_DUPLICADO_EN_PDF": duplicado_pdf, "QR_PROCESADO": qr_ok, "QR_NO_PROCESADO": qr_no,
               "ARCHIVO_RESULTADO": str(salida)}
    for k in CRITICOS:
        resumen[f"FALTA_{k}"] = sum(1 for g in regs if g.get(k) in FALTA)
    guardar_log_xlsx(detalles, resumen, log_xlsx)
    agregar_resumen_a_resultado(resumen, salida)

    print("-" * 70)
    print("PROCESO FINALIZADO")
    print(f"CAE unicos           : {total}")
    print(f"Procesadas           : {len(regs)}")
    print(f"CAE no encontrados   : {no_encontrado}")
    print(f"CAE duplicados PDF   : {duplicado_pdf}")
    print(f"QR_PROCESADO         : {qr_ok}")
    print(f"QR_NO_PROCESADO      : {qr_no}")
    print(f"Duracion             : {timedelta(seconds=int(round(duracion)))}")
    print("Avance final         : 100.00%")
    print(f"Resultado            : {salida}")
    print(f"Log XLSX             : {log_xlsx}")
    print("=" * 70)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())