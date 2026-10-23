# -*- coding: utf-8 -*-
"""
DEVOLVER OCR (REEMPLAZO DEL ORIGINAL)
-------------------------------------
Toma los PDFs con OCR generados por procesar_ocr_version.py (carpetas
05. Logs\\OCR_PARA_DEVOLVER_<ID>) y REEMPLAZA con ellos al PDF original
en 01. FACTURAS PENDIENTES CONTABILIZAR (AI). El archivo conserva el nombre
original (sin sufijo _OCR).

Resultado final: en cada carpeta queda UNA sola copia (la que tiene OCR).
Nada queda duplicado:
  - el PDF OCR se MUEVE desde 05. Logs (el backup se elimina al terminar);
  - si hay un <nombre>_OCR.pdf junto al original (formato del proceso anterior)
    y es identico al OCR, tambien se elimina;
  - las carpetas OCR_PARA_DEVOLVER_<ID> se eliminan completas (con subcarpetas) al terminar.

La fuente de verdad son los logs de OCR:
    05. Logs\\ocr_log_<ID>.xlsx

Controles antes de tocar un original (si alguno falla NO se reemplaza nada):
  - el original sigue sin texto (no lo cambiaron por otro / no ya tiene OCR);
  - original y OCR tienen la misma cantidad de paginas;
  - el original no fue modificado despues de generarse el OCR;
  - el OCR tiene texto reconocible y se puede abrir.
El reemplazo es atomico: se copia a un temporal junto al original, se valida y
recien ahi se pisa el original. Si algo falla, el original queda intacto.

Uso:
    python devolucion_ocr_facturas_version.py                  -> todas las corridas
    python devolucion_ocr_facturas_version.py 20261002094900   -> solo esa corrida
    python devolucion_ocr_facturas_version.py --simular        -> no toca nada, solo informa

Es seguro ejecutarlo varias veces (un original que ya tiene texto no se vuelve a tocar).

Registro de cada ejecucion:
    05. Logs\\Historial Devolucion OCR\\devolucion_log_<ID>.xlsx / .txt
    05. Logs\\Historial Devolucion OCR\\devolucion_estadisticas_<ID>.txt
"""
import os
import re
import sys
import shutil
import filecmp
import time
import traceback
from collections import OrderedDict
from datetime import datetime
import pymupdf as fitz
from openpyxl import Workbook, load_workbook

# ==========================================================
# ID UNICO DE PROCESO (YYYYMMDDHHMMSS)
# ==========================================================

ID_PROCESO = datetime.now().strftime("%Y%m%d%H%M%S")

# ==========================================================
# CONFIGURACION
# ==========================================================

DIRECTORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
DIRECTORIO_RAIZ = os.path.dirname(DIRECTORIO_SCRIPT)

# Carpeta real de facturas: vive en la raiz de la biblioteca "Facturacion Inteligente (IA)",
# al lado de "98. AUTOMATIZACION" (DIRECTORIO_RAIZ), no dentro de ella.
NOMBRE_CARPETA_FACTURAS = "01. FACTURAS PENDIENTES CONTABILIZAR (AI)"
CARPETA_FACTURAS = os.path.join(
    os.path.dirname(DIRECTORIO_RAIZ), NOMBRE_CARPETA_FACTURAS
)
CARPETA_RESULTADOS = os.path.join(DIRECTORIO_RAIZ, "04. Resultados")

# Todos los logs se guardan directamente en 05. Logs
CARPETA_LOGS = os.path.join(DIRECTORIO_RAIZ, "05. Logs")

# Los PDFs temporales generados por el proceso OCR siguen identificados
# mediante OCR_PARA_DEVOLVER_<ID>
PREFIJO_OCR_SALIDA = "OCR_PARA_DEVOLVER"

# Los logs y estadisticas de devolucion se guardan directamente en
# 05. Logs\\Historial Devolucion OCR (sin subcarpeta por ID de proceso;
# el ID ya forma parte del nombre de cada archivo).
CARPETA_DEVOLUCIONES = os.path.join(CARPETA_LOGS, "Historial Devolucion OCR")
CARPETA_CORRIDA = CARPETA_DEVOLUCIONES

os.makedirs(CARPETA_LOGS, exist_ok=True)
os.makedirs(CARPETA_DEVOLUCIONES, exist_ok=True)

ARCHIVO_LOG = os.path.join(CARPETA_CORRIDA, f"devolucion_log_{ID_PROCESO}.xlsx")
ARCHIVO_LOG_TXT = os.path.join(CARPETA_CORRIDA, f"devolucion_log_{ID_PROCESO}.txt")
ARCHIVO_ESTADISTICAS = os.path.join(
    CARPETA_CORRIDA,
    f"devolucion_estadisticas_{ID_PROCESO}.txt"
)

# Estados del log de OCR cuyo PDF con OCR debe devolverse.
# OCR_REALIZADO es el que genera procesar_ocr_version.py. OCR_PENDIENTE_DEVOLVER
# ya no lo genera ningun script actual: se mantiene SOLO para poder devolver
# logs de OCR de versiones anteriores que todavia tengan filas con ese estado.
ESTADOS_DEVOLVIBLES = {"OCR_REALIZADO", "OCR_PENDIENTE_DEVOLVER"}

FLUSH_CADA = 50  # filas de log que se acumulan antes de volcar al Excel

# Seguridad: no reemplazar un original que se modifico DESPUES de generarse el OCR
# (alguien pudo editarlo o cambiarlo por otro). Tolerancia en segundos.
VALIDAR_FECHA_MODIFICACION = True
TOLERANCIA_FECHA_SEG = 5

# ==========================================================
# MENSAJES EN PANTALLA + LOG DE TEXTO EN VIVO
# ==========================================================

def formato_hhmmss(segundos):
    segundos = int(segundos)
    minutos, seg = divmod(segundos, 60)
    horas, minutos = divmod(minutos, 60)
    return f"{horas:02d}:{minutos:02d}:{seg:02d}"


def mensaje(texto=""):
    """Muestra el mensaje en pantalla y lo agrega al log de texto."""
    print(texto, flush=True)
    try:
        with open(ARCHIVO_LOG_TXT, "a", encoding="utf-8") as f:
            hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{hora}] [{ID_PROCESO}] {texto}\n")
    except Exception:
        pass

# ==========================================================
# LOG XLSX (uno por corrida)
# ==========================================================

ENCABEZADO = [
    "Fecha",
    "IdCorridaOCR",          # corrida de OCR de donde viene el archivo
    "Archivo",               # ruta relativa a 01. FACTURAS PENDIENTES CONTABILIZAR (AI)
    "Estado",
    "Detalle",
    "DuracionSegundos",
    "RutaBackupOCR",         # de donde se tomo el PDF con OCR (se mueve, ya no queda ahi)
    "RutaFinal",             # PDF que queda en 01. FACTURAS (ahora con OCR)
    "TamanoOriginalBytes",   # tamano del PDF original que se reemplazo
    "TamanoOCRBytes",        # tamano del PDF con OCR que quedo
    "Paginas",
    "CopiaOCRJuntoEliminada",  # SI/NO: se borro un <nombre>_OCR.pdf duplicado
    "IdProceso"
]


def guardar_log(filas):
    if not filas:
        return

    if os.path.isfile(ARCHIVO_LOG):
        wb = load_workbook(ARCHIVO_LOG)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Log"
        ws.append(ENCABEZADO)

    for fila in filas:
        ws.append(list(fila) + [ID_PROCESO])

    wb.save(ARCHIVO_LOG)


def volcar_log(filas_log):
    """Guarda las filas pendientes. Si falla, las conserva y reintenta luego."""
    try:
        guardar_log(filas_log)
        filas_log.clear()
    except PermissionError:
        mensaje(f"AVISO: no se pudo escribir {os.path.basename(ARCHIVO_LOG)} "
                "(esta abierto en Excel?). Cierrelo para que se guarde el registro.")
    except Exception as e:
        mensaje(f"AVISO: no se pudo guardar {os.path.basename(ARCHIVO_LOG)}: {e}")


def nueva_fila(id_corrida, relativo, estado, detalle="", duracion=0,
               backup="", final="", tam_original="", tam_ocr="", paginas="",
               copia_eliminada=""):
    return [
        datetime.now(), id_corrida, relativo, estado, detalle, duracion,
        backup, final, tam_original, tam_ocr, paginas, copia_eliminada
    ]

# ==========================================================
# LECTURA DE LOS LOGS DE OCR
# ==========================================================

def listar_logs_ocr(filtro_id=None):
    """Busca directamente en 05. Logs archivos ocr_log_<ID>.xlsx.

    Devuelve una lista de tuplas (id_corrida, ruta_xlsx).
    """
    logs = []
    if not os.path.isdir(CARPETA_LOGS):
        return logs

    for archivo in sorted(os.listdir(CARPETA_LOGS)):
        m = re.fullmatch(r"ocr_log_(\d{14})\.xlsx", archivo, re.IGNORECASE)
        if not m:
            continue

        id_corrida = m.group(1)
        if filtro_id and id_corrida != filtro_id:
            continue

        logs.append((id_corrida, os.path.join(CARPETA_LOGS, archivo)))

    return logs


def leer_log_ocr(ruta_xlsx):
    """Lee el xlsx de OCR y devuelve una lista de diccionarios por columna."""
    columnas = ["Archivo", "Estado", "Detalle", "RutaOrigen", "RutaActual"]
    filas = []

    wb = load_workbook(ruta_xlsx, read_only=True, data_only=True)
    try:
        ws = wb.active
        iterador = ws.iter_rows(values_only=True)
        encabezado = next(iterador, None)

        if not encabezado:
            return filas

        indice = {str(n).strip(): i for i, n in enumerate(encabezado) if n}

        for fila in iterador:
            if not fila or not any(fila):
                continue

            registro = {}
            for col in columnas:
                i = indice.get(col)
                valor = fila[i] if i is not None and i < len(fila) else None
                registro[col] = "" if valor is None else str(valor).strip()

            filas.append(registro)
    finally:
        wb.close()

    return filas

# ==========================================================
# UTILIDADES DE ARCHIVOS / PDF
# ==========================================================

def esta_dentro(base, ruta):
    base = os.path.normcase(os.path.abspath(base))
    ruta = os.path.normcase(os.path.abspath(ruta))
    return ruta.startswith(base + os.sep)


def contar_paginas(ruta):
    try:
        with fitz.open(ruta, filetype="pdf") as doc:
            return doc.page_count
    except Exception:
        return None


def tiene_texto(ruta):
    try:
        with fitz.open(ruta, filetype="pdf") as doc:
            for pagina in doc:
                texto = pagina.get_text()
                if texto and texto.strip():
                    return True
    except Exception:
        pass
    return False


def borrar_si_existe(ruta):
    try:
        if os.path.isfile(ruta):
            os.remove(ruta)
    except Exception:
        pass


def carpetas_backup(solo_id=None):
    """Carpetas OCR_PARA_DEVOLVER* de 05. Logs."""
    resultado = []

    if not os.path.isdir(CARPETA_LOGS):
        return resultado

    for nombre in sorted(os.listdir(CARPETA_LOGS)):
        ruta = os.path.join(CARPETA_LOGS, nombre)

        if not os.path.isdir(ruta) or not nombre.startswith(PREFIJO_OCR_SALIDA):
            continue
        if solo_id and nombre != f"{PREFIJO_OCR_SALIDA}_{solo_id}":
            continue

        resultado.append(ruta)

    return resultado


def resolver_relativo(fila):
    """Ruta relativa a 01. FACTURAS PENDIENTES CONTABILIZAR (AI) del PDF original."""
    relativo = fila.get("Archivo", "")

    if not relativo and fila.get("RutaOrigen"):
        try:
            candidato = os.path.relpath(fila["RutaOrigen"], CARPETA_FACTURAS)
            if not candidato.startswith(".."):
                relativo = candidato
        except ValueError:
            pass

    return relativo


def localizar_ocr(fila, relativo, ids_corrida, todas_las_carpetas):
    """Busca el PDF con OCR: primero la ruta del log, luego por ID y por
    estructura relativa (sirve si cambio el usuario/PC)."""
    rel_ocr = os.path.splitext(relativo)[0] + "_OCR.pdf"

    candidatos = []

    if fila.get("RutaActual"):
        candidatos.append(fila["RutaActual"])

    for i in ids_corrida:
        candidatos.append(
            os.path.join(CARPETA_LOGS, f"{PREFIJO_OCR_SALIDA}_{i}", rel_ocr)
        )

    for carpeta in todas_las_carpetas:
        candidatos.append(os.path.join(carpeta, rel_ocr))

    for c in candidatos:
        if c and os.path.isfile(c):
            return c

    return None

# ==========================================================
# LIMPIEZA DE CARPETAS OCR_PARA_DEVOLVER_<ID>
# ==========================================================

def _contar_pdfs(carpeta):
    total = 0
    for _, _, archivos in os.walk(carpeta):
        total += sum(1 for a in archivos if a.lower().endswith(".pdf"))
    return total


def borrar_carpetas_backup(filas_log, solo_id=None):
    """Elimina COMPLETAS (carpeta raiz + todas sus subcarpetas y archivos) las
    carpetas OCR_PARA_DEVOLVER_<ID> una vez terminada la devolucion.

    Si al borrar queda algun PDF que no se pudo devolver (conflicto, OCR sin
    texto, huerfano, etc.) se deja constancia de cada uno en el log antes de
    eliminarlo. El original nunca se toca: si hacia falta, el proximo OCR
    volvera a generar esa version.

    Devuelve (carpetas_eliminadas, pdfs_descartados, carpetas_que_siguen)."""
    eliminadas = 0
    pdfs_descartados = 0
    siguen = 0

    for carpeta in carpetas_backup(solo_id):
        # Dejar registro de lo que se descarta (PDFs que no se devolvieron)
        restantes = []
        for raiz, _, archivos in os.walk(carpeta):
            for a in archivos:
                if a.lower().endswith(".pdf"):
                    restantes.append(os.path.join(raiz, a))

        for ruta in restantes:
            filas_log.append(nueva_fila(
                "", os.path.relpath(ruta, CARPETA_LOGS), "BACKUP_DESCARTADO",
                "PDF con OCR que no se devolvio y se elimino junto con su carpeta "
                "OCR_PARA_DEVOLVER. El original no se modifico.",
                backup=ruta
            ))

        # Borrado completo, con reintentos por si OneDrive tiene archivos tomados
        for intento in range(3):
            shutil.rmtree(carpeta, ignore_errors=True)
            if not os.path.exists(carpeta):
                break
            time.sleep(1)

        if os.path.exists(carpeta):
            siguen += 1
            mensaje(f"AVISO: no se pudo eliminar por completo {carpeta} "
                    "(archivo en uso o sincronizando con OneDrive). Borrar a mano.")
        else:
            eliminadas += 1
            pdfs_descartados += len(restantes)
            mensaje(f"Carpeta OCR_PARA_DEVOLVER eliminada completamente: {carpeta}"
                    + (f" ({len(restantes)} PDF sin devolver descartados)" if restantes else ""))

    return eliminadas, pdfs_descartados, siguen

# ==========================================================
# VERIFICACION FINAL DE 01. FACTURAS PENDIENTES CONTABILIZAR (AI)
# ==========================================================

def verificar_facturas(filas_log):
    """Cuenta los PDFs de facturas y detecta duplicados que siguen existiendo:
    <nombre>_OCR.pdf junto a su original (formato del proceso anterior)."""
    total = 0
    ocr_junto = 0
    ocr_sin_original = 0

    for raiz, _, archivos in os.walk(CARPETA_FACTURAS):
        nombres = {a.lower() for a in archivos}

        for a in archivos:
            n = a.lower()

            if not n.endswith(".pdf"):
                continue

            total += 1

            if not n.endswith("_ocr.pdf"):
                continue

            ruta = os.path.join(raiz, a)
            base = a[:-len("_OCR.pdf")]

            if (base + ".pdf").lower() in nombres:
                ocr_junto += 1
                filas_log.append(nueva_fila(
                    "", os.path.relpath(ruta, CARPETA_FACTURAS),
                    "DUPLICADO_OCR_JUNTO",
                    "Sigue existiendo un _OCR.pdf junto a su original. No se elimino "
                    "porque no hay un OCR en backup que lo respalde o es distinto.",
                    final=ruta
                ))
            else:
                ocr_sin_original += 1
                filas_log.append(nueva_fila(
                    "", os.path.relpath(ruta, CARPETA_FACTURAS),
                    "OCR_SIN_ORIGINAL",
                    "Hay un _OCR.pdf en 01. FACTURAS PENDIENTES CONTABILIZAR (AI) sin su PDF original al lado",
                    final=ruta
                ))

    return total, ocr_junto, ocr_sin_original

# ==========================================================
# MAIN
# ==========================================================

def main():
    inicio_global = time.time()
    inicio_dt = datetime.now()

    # ---- argumentos: [ID corrida] [--simular]
    simular = False
    filtro_id = None

    for arg in (a.strip() for a in sys.argv[1:]):
        if arg == "--simular":
            simular = True
        elif re.fullmatch(r"\d{14}", arg):
            filtro_id = arg
        else:
            print(f"Argumento no valido: {arg}")
            print("Uso: python devolucion_ocr_facturas_version.py [YYYYMMDDHHMMSS] [--simular]")
            return 1

    if not os.path.isdir(CARPETA_FACTURAS):
        mensaje("ERROR: no se encuentra la carpeta de facturas:")
        mensaje(CARPETA_FACTURAS)
        mensaje("Verifique que este sincronizada y que el nombre sea exacto.")
        return 1

    filas_log = []
    cont = OrderedDict([
        ("REEMPLAZADO", 0),
        ("YA_REEMPLAZADO", 0),
        ("CONFLICTO_ORIGINAL", 0),
        ("OCR_SIN_TEXTO", 0),
        ("ORIGINAL_NO_ENCONTRADO", 0),
        ("OCR_NO_ENCONTRADO", 0),
        ("SIMULADO", 0),
        ("ERROR", 0),
    ])
    avisos_total = 0
    copias_eliminadas = 0
    huerfanos = 0
    errores_ocr_previos = 0
    corridas_leidas = 0
    a_procesar = 0
    bytes_liberados = 0
    interrumpido = False

    modo = "SIMULACION" if simular else "REAL"

    mensaje("====================================")
    mensaje("REEMPLAZO DE ORIGINALES POR PDFs CON OCR")
    mensaje("====================================")
    mensaje(f"ID de proceso: {ID_PROCESO}")
    mensaje(f"Carpeta de registro: {CARPETA_CORRIDA}")
    if simular:
        mensaje("*** MODO SIMULACION: no se modifica ni se borra ningun archivo ***")
    if filtro_id:
        mensaje(f"Solo se procesa la corrida de OCR: {filtro_id}")
    mensaje("(Para detener el proceso en forma segura: Ctrl + C)")
    mensaje("------------------------------------")

    # ------------------------------------------------------
    # 1) Leer los logs de OCR y armar la lista de tareas
    # ------------------------------------------------------
    tareas = []   # (id_corrida, ids_corrida, relativo, fila)
    backups_usados = set()

    logs = listar_logs_ocr(filtro_id)

    if not logs:
        mensaje("No se encontraron logs de OCR para procesar en: " + CARPETA_LOGS)

    for id_corrida, ruta_xlsx in logs:
        try:
            filas = leer_log_ocr(ruta_xlsx)
        except Exception as e:
            cont["ERROR"] += 1
            mensaje(f"ERROR: no se pudo leer {ruta_xlsx}: {e}")
            filas_log.append(nueva_fila(
                id_corrida, "", "ERROR", f"No se pudo leer el log de OCR: {e}",
                backup=ruta_xlsx
            ))
            continue

        corridas_leidas += 1

        # El ID de la corrida es el de la carpeta; el del nombre del archivo
        # puede diferir si se renombraron carpetas a mano.
        m = re.search(r"ocr_log_(\d{14})", os.path.basename(ruta_xlsx), re.IGNORECASE)
        ids_corrida = [id_corrida]
        if m and m.group(1) not in ids_corrida:
            ids_corrida.append(m.group(1))

        pendientes = 0
        for fila in filas:
            estado = fila.get("Estado", "")

            if estado == "ERROR":
                errores_ocr_previos += 1

            if estado in ESTADOS_DEVOLVIBLES:
                pendientes += 1
                tareas.append((id_corrida, ids_corrida, resolver_relativo(fila), fila))

        mensaje(f"Corrida OCR {id_corrida}: {len(filas)} filas en el log, "
                f"{pendientes} con OCR para aplicar")

    a_procesar = len(tareas)
    todas_las_carpetas = carpetas_backup()

    mensaje(f"Total de PDFs con OCR a aplicar sobre su original: {a_procesar}")
    mensaje("------------------------------------")

    # ------------------------------------------------------
    # 2) Reemplazar cada original por su version con OCR
    # ------------------------------------------------------
    try:
        for numero, (id_corrida, ids_corrida, relativo, fila) in enumerate(tareas, start=1):
            prefijo = f"[{numero}/{a_procesar}]"
            inicio = time.time()

            def registrar(estado, detalle="", backup="", final="",
                          tam_original="", tam_ocr="", paginas="", copia=""):
                cont[estado] += 1
                filas_log.append(nueva_fila(
                    id_corrida, relativo, estado, detalle,
                    round(time.time() - inicio, 2),
                    backup, final, tam_original, tam_ocr, paginas, copia
                ))
                mensaje(f"{prefijo} {relativo or '(sin ruta)'} -> {estado}"
                        + (f": {detalle}" if detalle else ""))

            if not relativo:
                registrar("ERROR", "El log no indica la ruta del archivo original")
                continue

            original = os.path.join(CARPETA_FACTURAS, relativo)
            # Copia "_OCR.pdf" junto al original (la dejaba el proceso anterior)
            copia_junto = os.path.splitext(original)[0] + "_OCR.pdf"

            try:
                # Seguridad: nunca escribir fuera de 01. FACTURAS PENDIENTES CONTABILIZAR (AI)
                if not esta_dentro(CARPETA_FACTURAS, original):
                    registrar("ERROR", "La ruta del original queda fuera de "
                              "01. FACTURAS PENDIENTES CONTABILIZAR (AI)", final=original)
                    continue

                # El original debe seguir en su lugar
                if not os.path.isfile(original):
                    registrar("ORIGINAL_NO_ENCONTRADO",
                              "El PDF original ya no esta (movido, renombrado o borrado). "
                              "No se reemplaza nada.", final=original)
                    continue

                # Buscar el PDF con OCR en el backup
                ruta_ocr = localizar_ocr(fila, relativo, ids_corrida, todas_las_carpetas)

                if not ruta_ocr:
                    if tiene_texto(original):
                        registrar("YA_REEMPLAZADO",
                                  "El original ya contiene texto y el backup ya no existe: "
                                  "se considera reemplazado en una corrida anterior",
                                  final=original)
                    else:
                        registrar("OCR_NO_ENCONTRADO",
                                  "No se encontro el PDF con OCR en ninguna carpeta "
                                  "OCR_PARA_DEVOLVER* y el original sigue sin texto",
                                  backup=fila.get("RutaActual", ""), final=original)
                    continue

                backups_usados.add(os.path.normcase(os.path.abspath(ruta_ocr)))
                tam_ocr = os.path.getsize(ruta_ocr)
                tam_orig = os.path.getsize(original)

                # El original ya tiene texto pero el backup sigue ahi: no se toca nada
                if tiene_texto(original):
                    registrar("CONFLICTO_ORIGINAL",
                              "El original ya tiene texto pero el backup sigue en 05. Logs. "
                              "No se toca nada; revisar y borrar el backup a mano si corresponde.",
                              backup=ruta_ocr, final=original,
                              tam_original=tam_orig, tam_ocr=tam_ocr)
                    continue

                pag_ocr = contar_paginas(ruta_ocr)
                pag_orig = contar_paginas(original)

                if pag_ocr is None:
                    registrar("ERROR", "El PDF con OCR del backup no se puede abrir",
                              backup=ruta_ocr, final=original,
                              tam_original=tam_orig, tam_ocr=tam_ocr)
                    continue

                if pag_orig is None:
                    registrar("CONFLICTO_ORIGINAL",
                              "El original no se puede abrir; no se reemplaza",
                              backup=ruta_ocr, final=original,
                              tam_original=tam_orig, tam_ocr=tam_ocr, paginas=pag_ocr)
                    continue

                if pag_ocr != pag_orig:
                    registrar("CONFLICTO_ORIGINAL",
                              f"Cantidad de paginas distinta (original {pag_orig}, OCR {pag_ocr}); "
                              "no se reemplaza",
                              backup=ruta_ocr, final=original,
                              tam_original=tam_orig, tam_ocr=tam_ocr, paginas=pag_ocr)
                    continue

                if (VALIDAR_FECHA_MODIFICACION and
                        os.path.getmtime(original) > os.path.getmtime(ruta_ocr) + TOLERANCIA_FECHA_SEG):
                    registrar("CONFLICTO_ORIGINAL",
                              "El original fue modificado despues de generarse el OCR "
                              "(posible edicion o archivo distinto con el mismo nombre); no se reemplaza",
                              backup=ruta_ocr, final=original,
                              tam_original=tam_orig, tam_ocr=tam_ocr, paginas=pag_ocr)
                    continue

                if not tiene_texto(ruta_ocr):
                    # Se deja todo como esta: el backup evita que procesar_ocr lo reintente
                    registrar("OCR_SIN_TEXTO",
                              "El OCR no contiene texto reconocible (hoja en blanco o ilegible). "
                              "No se reemplaza; el original y el backup quedan sin cambios. Revisar.",
                              backup=ruta_ocr, final=original,
                              tam_original=tam_orig, tam_ocr=tam_ocr, paginas=pag_ocr)
                    continue

                # Copia duplicada junto al original (proceso anterior): solo si es identica
                hay_copia = os.path.isfile(copia_junto)
                copia_identica = hay_copia and filecmp.cmp(copia_junto, ruta_ocr, shallow=False)

                if simular:
                    extra = ""
                    if copia_identica:
                        extra = " y se eliminaria el _OCR.pdf duplicado junto al original"
                    elif hay_copia:
                        extra = " (hay un _OCR.pdf distinto junto al original: no se tocaria)"
                    registrar("SIMULADO",
                              "Se reemplazaria el original por la version OCR, "
                              "se borraria el backup" + extra,
                              backup=ruta_ocr, final=original,
                              tam_original=tam_orig, tam_ocr=tam_ocr, paginas=pag_ocr,
                              copia="SI" if copia_identica else "NO")
                    continue

                # Reemplazo seguro: temporal junto al original, validacion y luego se pisa
                tmp = original + ".ocr.tmp"
                try:
                    shutil.copy2(ruta_ocr, tmp)

                    if os.path.getsize(tmp) != tam_ocr:
                        raise IOError("el tamano copiado no coincide con el del backup")

                    if contar_paginas(tmp) != pag_ocr:
                        raise ValueError("el PDF copiado no se puede abrir o cambio de paginas")

                    os.replace(tmp, original)   # atomico: nunca queda sin archivo
                except BaseException:
                    borrar_si_existe(tmp)
                    raise

                # A partir de aca el original ya es la version con OCR.
                # Se eliminan los duplicados.
                avisos = []
                liberado = tam_orig - tam_ocr   # el original viejo fue reemplazado

                try:
                    os.remove(ruta_ocr)
                    liberado += tam_ocr
                except OSError as e:
                    avisos.append(f"no se pudo borrar el backup ({e}); queda duplicado en 05. Logs")

                copia_borrada = False
                if copia_identica:
                    try:
                        os.remove(copia_junto)
                        copia_borrada = True
                        copias_eliminadas += 1
                        liberado += tam_ocr
                    except OSError as e:
                        avisos.append(f"no se pudo borrar el _OCR.pdf duplicado ({e})")
                elif hay_copia:
                    avisos.append("hay un _OCR.pdf distinto junto al original; no se borro, revisar")

                bytes_liberados += liberado

                if avisos:
                    avisos_total += 1

                registrar("REEMPLAZADO",
                          ("AVISO: " + "; ".join(avisos)) if avisos else "",
                          backup=ruta_ocr, final=original,
                          tam_original=tam_orig, tam_ocr=tam_ocr, paginas=pag_ocr,
                          copia="SI" if copia_borrada else "NO")

            except KeyboardInterrupt:
                raise

            except Exception as e:
                registrar("ERROR", str(e), final=original)

            if len(filas_log) >= FLUSH_CADA:
                volcar_log(filas_log)

    except KeyboardInterrupt:
        interrumpido = True
        mensaje("")
        mensaje("*** Proceso detenido por el usuario. Se guarda lo realizado hasta ahora. ***")

    # ------------------------------------------------------
    # 3) Huerfanos: PDFs en el backup que ningun log reclamo
    # ------------------------------------------------------
    if not interrumpido:
        for carpeta in carpetas_backup(filtro_id):
            for raiz, _, archivos in os.walk(carpeta):
                for a in archivos:
                    if not a.lower().endswith(".pdf"):
                        continue

                    ruta = os.path.join(raiz, a)

                    if os.path.normcase(os.path.abspath(ruta)) in backups_usados:
                        continue

                    huerfanos += 1
                    filas_log.append(nueva_fila(
                        "", os.path.relpath(ruta, CARPETA_LOGS), "HUERFANO_EN_BACKUP",
                        "PDF en una carpeta de backup que ningun log de OCR (OCR_REALIZADO) "
                        "reclama. No se toca; revisar.",
                        backup=ruta
                    ))

    # ------------------------------------------------------
    # 4) Eliminar las carpetas OCR_PARA_DEVOLVER_<ID> (con todas sus subcarpetas)
    # ------------------------------------------------------
    carpetas_borradas = 0
    pdfs_descartados = 0
    carpetas_siguen = 0
    if not simular and not interrumpido:
        carpetas_borradas, pdfs_descartados, carpetas_siguen = borrar_carpetas_backup(
            filas_log, filtro_id
        )

    # ------------------------------------------------------
    # 5) Verificacion final de 01. FACTURAS PENDIENTES CONTABILIZAR (AI)
    # ------------------------------------------------------
    total_pdfs, ocr_junto, ocr_sin_original = verificar_facturas(filas_log)

    volcar_log(filas_log)

    # ------------------------------------------------------
    # 6) Estadisticas y resumen
    # ------------------------------------------------------
    fin_dt = datetime.now()
    duracion_total = time.time() - inicio_global
    mb_liberados = round(bytes_liberados / (1024 * 1024), 1)

    try:
        with open(ARCHIVO_ESTADISTICAS, "w", encoding="utf-8") as f:
            f.write("====================================\n")
            f.write("ESTADISTICAS DEVOLUCION OCR\n")
            f.write("====================================\n")
            f.write(f"ID_PROCESO: {ID_PROCESO}\n")
            f.write(f"Estado: {'INTERRUMPIDO' if interrumpido else 'COMPLETO'}\n")
            f.write(f"Modo: {modo}\n")
            f.write(f"Inicio: {inicio_dt.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Fin: {fin_dt.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Tiempo Total: {formato_hhmmss(duracion_total)}\n")
            f.write("\n")
            f.write(f"Corridas OCR leidas: {corridas_leidas}\n")
            f.write(f"PDFs con OCR a aplicar: {a_procesar}\n")
            f.write(f"Originales reemplazados: {cont['REEMPLAZADO']}\n")
            f.write(f"Ya reemplazados: {cont['YA_REEMPLAZADO']}\n")
            f.write(f"Conflictos: {cont['CONFLICTO_ORIGINAL']}\n")
            f.write(f"OCR sin texto: {cont['OCR_SIN_TEXTO']}\n")
            f.write(f"Originales no encontrados: {cont['ORIGINAL_NO_ENCONTRADO']}\n")
            f.write(f"OCR no encontrados: {cont['OCR_NO_ENCONTRADO']}\n")
            f.write(f"Errores: {cont['ERROR']}\n")
            f.write(f"Avisos: {avisos_total}\n")
            f.write(f"Copias OCR eliminadas: {copias_eliminadas}\n")
            f.write(f"PDFs huerfanos en backup: {huerfanos}\n")
            f.write(f"Carpetas OCR_PARA_DEVOLVER eliminadas: {carpetas_borradas}\n")
            f.write(f"PDFs sin devolver descartados con la carpeta: {pdfs_descartados}\n")
            f.write(f"Carpetas OCR_PARA_DEVOLVER que no se pudieron eliminar: {carpetas_siguen}\n")
            f.write(f"Espacio liberado (MB): {mb_liberados}\n")
            f.write(f"PDFs en facturas: {total_pdfs}\n")
            f.write(f"_OCR junto al original: {ocr_junto}\n")
            f.write(f"_OCR sin original: {ocr_sin_original}\n")
    except Exception as e:
        mensaje(f"AVISO: no se pudo guardar el archivo de estadisticas: {e}")

    mensaje("")
    mensaje("====================================")
    mensaje(f"RESUMEN ({modo})" + (" - INTERRUMPIDO" if interrumpido else ""))
    mensaje("====================================")
    mensaje(f"Corridas de OCR leidas: {corridas_leidas}")
    mensaje(f"PDFs con OCR a aplicar: {a_procesar}")
    if simular:
        mensaje(f"Se reemplazarian: {cont['SIMULADO']}")
    else:
        mensaje(f"Originales reemplazados por su version OCR: {cont['REEMPLAZADO']}")
        mensaje(f"Copias _OCR.pdf duplicadas eliminadas: {copias_eliminadas}")
        mensaje(f"Espacio liberado aprox.: {mb_liberados} MB")
        mensaje(f"Carpetas OCR_PARA_DEVOLVER eliminadas: {carpetas_borradas}")
        mensaje(f"PDFs sin devolver descartados con la carpeta: {pdfs_descartados}")
        if carpetas_siguen:
            mensaje(f"AVISO: {carpetas_siguen} carpeta(s) OCR_PARA_DEVOLVER no se pudieron eliminar")
    mensaje(f"Ya reemplazados en corridas anteriores: {cont['YA_REEMPLAZADO']}")
    mensaje(f"Conflictos (no se tocaron): {cont['CONFLICTO_ORIGINAL']}")
    mensaje(f"OCR sin texto (no se tocaron): {cont['OCR_SIN_TEXTO']}")
    mensaje(f"Originales no encontrados: {cont['ORIGINAL_NO_ENCONTRADO']}")
    mensaje(f"OCR no encontrados en backup: {cont['OCR_NO_ENCONTRADO']}")
    mensaje(f"Errores: {cont['ERROR']}")
    mensaje(f"Avisos: {avisos_total}")
    mensaje("------------------------------------")
    mensaje("DATOS PARA REVISAR")
    mensaje(f"PDFs huerfanos en backup (sin log que los reclame): {huerfanos}")
    mensaje(f"Errores registrados en los logs de OCR leidos: {errores_ocr_previos}")
    mensaje(f"PDFs en 01. FACTURAS PENDIENTES CONTABILIZAR (AI): {total_pdfs}")
    mensaje(f"  _OCR.pdf que siguen junto a su original (duplicados): {ocr_junto}")
    mensaje(f"  _OCR.pdf sin su original: {ocr_sin_original}")
    mensaje("------------------------------------")
    mensaje(f"Inicio: {inicio_dt.strftime('%Y-%m-%d %H:%M:%S')}  "
            f"Fin: {fin_dt.strftime('%Y-%m-%d %H:%M:%S')}")
    mensaje(f"Tiempo total: {formato_hhmmss(duracion_total)} "
            f"({round(duracion_total / 60, 2)} min)")
    mensaje("====================================")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        mensaje("")
        mensaje("*** Ocurrio un error inesperado. Avise al area de sistemas con este detalle: ***")
        mensaje(traceback.format_exc())
        raise SystemExit(1)
