from pathlib import Path
from datetime import datetime
import re
import sys

try:
    import pymupdf as fitz
    from openpyxl import Workbook
    from openpyxl.styles import Font
except ImportError as exc:
    sys.exit(
        "Fallo de dependencias:\n"
        "Ejecutar: pip install pymupdf openpyxl\n\n"
        f"{exc}"
    )

SCRIPT_DIR = Path(__file__).resolve().parent
NOMBRE_CARPETA_AUTOMATIZACION = "98. AUTOMATIZACION"
NOMBRE_CARPETA_RESULTADOS = "04. Resultados"
PREFIJO_CARPETA_FACTURAS = "01. FACTURAS PENDIENTES"

def ubicar_raiz_automatizacion() -> Path:
    for carpeta in [SCRIPT_DIR, *SCRIPT_DIR.parents]:
        if carpeta.name.upper() == NOMBRE_CARPETA_AUTOMATIZACION.upper():
            return carpeta
    return SCRIPT_DIR.parent

RAIZ_AUTOMATIZACION = ubicar_raiz_automatizacion()

def ubicar_carpeta_facturas() -> Path:
    if len(sys.argv) > 1:
        ruta = Path(sys.argv[1]).expanduser().resolve()
        if not ruta.is_dir():
            sys.exit(f"Directorio inválido: {ruta}")
        return ruta

    sitio = RAIZ_AUTOMATIZACION.parent
    for carpeta in sorted(sitio.iterdir()):
        if carpeta.is_dir() and carpeta.name.upper().startswith(PREFIJO_CARPETA_FACTURAS):
            return carpeta

    sys.exit(f"Directorio origen no localizado '{PREFIJO_CARPETA_FACTURAS}...' en:\n{sitio}")

def ubicar_carpeta_resultados() -> Path:
    destino = RAIZ_AUTOMATIZACION / NOMBRE_CARPETA_RESULTADOS
    destino.mkdir(parents=True, exist_ok=True)
    return destino

RAIZ = ubicar_carpeta_facturas()
CARPETA_SALIDA = ubicar_carpeta_resultados()
SALIDA = CARPETA_SALIDA / "Listado_CAE_14_digitos.xlsx"

# Expresión regular flexible para capturar "CAE", "C.A.E." o variaciones OCR como "CAF"
ETIQUETA_CAE = re.compile(r"c\.?\s*a\.?\s*[ef]\.?", re.IGNORECASE)
BLOQUE_NUMERICO = re.compile(r"(?<!\d)(\d{14,})(?!\d)")

# Ventana de proximidad ampliada para cubrir desfases de layout en el flujo de texto del PDF
VENTANA_PROXIMIDAD = 2500

def procesar_candidato(numero: str) -> str | None:
    """
    Valida un string de números candidato a CAE según sus posiciones de derecha a izquierda:
    - Posiciones 1 a 13 (desde la derecha): cualquier número entre 0 y 9.
    - Posición 14 (desde la derecha): diferente de '0' (<> '0').
    - Posiciones posteriores a la 14 (hacia la izquierda): deben ser todas iguales a '0'.
      Si alguna posición posterior a 14 es diferente a '0', el string NO es válido y se descarta.
    """
    if len(numero) < 14:
        return None

    # Posición 14 desde la derecha (índice -14 en Python)
    posicion_14 = numero[-14]
    if posicion_14 == '0':
        return None

    # Posiciones posteriores a la 14 (de la 15 en adelante hacia la izquierda, corte :-14)
    posteriores = numero[:-14]
    if posteriores and any(c != '0' for c in posteriores):
        return None

    # Si cumple todas las condiciones, retorna los 14 dígitos del CAE
    return numero[-14:]

def extraer_y_buscar_caes(pdf_path: Path):
    encontrados = set()
    tiene_texto = False
    error_io = None

    try:
        with fitz.open(pdf_path) as documento:
            for pagina in documento:
                texto = pagina.get_text("text")
                
                if texto.strip():
                    tiene_texto = True

                # Excluir páginas que no contienen la expresión de etiqueta CAE / CAF
                iter_etiquetas = list(ETIQUETA_CAE.finditer(texto))
                if not iter_etiquetas:
                    continue
                
                # Pre-procesar todos los bloques numéricos válidos en la página entera
                candidatos_validos = []
                for match_num in BLOQUE_NUMERICO.finditer(texto):
                    valido = procesar_candidato(match_num.group(1))
                    if valido:
                        candidatos_validos.append((valido, match_num.start(), match_num.end()))
                
                # 1) Buscar en la ventana de líneas (línea actual, una línea por encima y una línea por debajo)
                lineas = texto.splitlines()
                for idx, linea in enumerate(lineas):
                    if ETIQUETA_CAE.search(linea):
                        lineas_contexto = []
                        if idx > 0:
                            lineas_contexto.append(lineas[idx - 1])
                        lineas_contexto.append(linea)
                        if idx + 1 < len(lineas):
                            lineas_contexto.append(lineas[idx + 1])
                        
                        bloque_contexto = "\n".join(lineas_contexto)
                        for match_num in BLOQUE_NUMERICO.finditer(bloque_contexto):
                            valido = procesar_candidato(match_num.group(1))
                            if valido:
                                encontrados.add(valido)

                # 2) Evaluar direccionalmente respecto a cada etiqueta encontrada con ventana ampliada
                for etiqueta in iter_etiquetas:
                    cae_encontrado = None
                    
                    # Buscar a la DERECHA
                    candidatos_derecha = [
                        c for c in candidatos_validos
                        if 0 <= (c[1] - etiqueta.end()) <= VENTANA_PROXIMIDAD
                    ]
                    if candidatos_derecha:
                        candidatos_derecha.sort(key=lambda c: c[1] - etiqueta.end())
                        cae_encontrado = candidatos_derecha[0][0]
                    
                    # Buscar a la IZQUIERDA si no se encontró a la derecha
                    if not cae_encontrado:
                        candidatos_izquierda = [
                            c for c in candidatos_validos
                            if 0 <= (etiqueta.start() - c[2]) <= VENTANA_PROXIMIDAD
                        ]
                        if candidatos_izquierda:
                            candidatos_izquierda.sort(key=lambda c: etiqueta.start() - c[2])
                            cae_encontrado = candidatos_izquierda[0][0]
                    
                    if cae_encontrado:
                        encontrados.add(cae_encontrado)

    except Exception as exc:
        error_io = f"ERROR_IO: {exc}"

    return sorted(encontrados), tiene_texto, error_io

def ruta_relativa_carpeta(pdf: Path) -> str:
    padre = pdf.relative_to(RAIZ).parent
    return "" if str(padre) == "." else str(padre)

def obtener_timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def main() -> None:
    print(f"Directorio de escaneo activo: {RAIZ}")

    filas_cae = []
    log_eventos = []
    caes_globales = set()
    metricas = {"revisados": 0, "sin_texto": 0, "sin_cae": 0, "duplicados": 0, "errores": 0}

    for pdf in sorted(RAIZ.rglob("*.pdf"), key=lambda p: str(p).lower()):
        metricas["revisados"] += 1
        ruta_rel = ruta_relativa_carpeta(pdf)
        
        caes, tiene_texto, error = extraer_y_buscar_caes(pdf)

        if error:
            metricas["errores"] += 1
            log_eventos.append([obtener_timestamp(), pdf.name, ruta_rel, "ERROR LECTURA", error])
            continue

        if not tiene_texto:
            metricas["sin_texto"] += 1
            log_eventos.append([obtener_timestamp(), pdf.name, ruta_rel, "OMITIDO", "Sin texto extraíble (Requiere OCR)"])
            continue

        if not caes:
            metricas["sin_cae"] += 1
            log_eventos.append([obtener_timestamp(), pdf.name, ruta_rel, "OMITIDO", "Ningún CAE válido localizado en ventana de proximidad"])
            continue

        for cae in caes:
            if cae in caes_globales:
                metricas["duplicados"] += 1
                log_eventos.append([obtener_timestamp(), pdf.name, ruta_rel, "DUPLICADO DESCARTADO", cae])
                continue
            
            caes_globales.add(cae)
            filas_cae.append((pdf.name, str(pdf.resolve()), ruta_rel, cae))
            log_eventos.append([obtener_timestamp(), pdf.name, ruta_rel, "PROCESADO EXITOSAMENTE", cae])

    # ========================================================
    # CREACIÓN DEL ARCHIVO DE SALIDA UNIFICADO
    # ========================================================
    libro = Workbook()

    # --- Hoja 1: Resultados CAE ---
    hoja_cae = libro.active
    hoja_cae.title = "CAE Extraidos"
    hoja_cae.append(["Nombre de archivo", "Ruta completa", "Ruta relativa", "CAE"])
    
    for celda in hoja_cae[1]:
        celda.font = Font(bold=True)

    for fila in filas_cae:
        hoja_cae.append(fila)

    for celda in hoja_cae["D"][1:]:
        celda.number_format = "@"

    hoja_cae.freeze_panes = "A2"
    hoja_cae.auto_filter.ref = hoja_cae.dimensions

    for columna, ancho in {"A": 45, "B": 110, "C": 65, "D": 18}.items():
        hoja_cae.column_dimensions[columna].width = ancho

    # --- Hoja 2: Log de Eventos ---
    hoja_log = libro.create_sheet(title="Log de Eventos")
    hoja_log.append(["Timestamp", "Archivo", "Ruta Relativa", "Estado", "Detalle / CAE detectado"])
    
    for celda in hoja_log[1]:
        celda.font = Font(bold=True)

    for evento in log_eventos:
        hoja_log.append(evento)
        
    for celda in hoja_log["E"][1:]:
        celda.number_format = "@"

    hoja_log.freeze_panes = "A2"
    hoja_log.auto_filter.ref = hoja_log.dimensions

    for columna, ancho in {"A": 22, "B": 45, "C": 40, "D": 30, "E": 55}.items():
        hoja_log.column_dimensions[columna].width = ancho

    # ========================================================
    # GUARDADO_DE_DATOS
    # ========================================================
    try:
        libro.save(SALIDA)
    except PermissionError as e:
        sys.exit(f"Excepción de acceso I/O. El archivo de salida está bloqueado.\nDetalle: {e}")

    print(f"Salida de datos escrita en: {SALIDA}")

if __name__ == "__main__":
    main()