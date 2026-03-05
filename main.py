"""
╔══════════════════════════════════════════════════════════════╗
║   PDF Formulario Agentes / Directivos  →  Excel             ║
╚══════════════════════════════════════════════════════════════╝
"""
# ToDo: verificar que el mapeo de directivos va bien
# ToDo: arreglar el bug de valores intercambiados en agentes (acepto/no procede)

# Dependencias: pip install pdfplumber openpyxl

import re
import shutil
import sys
import io
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

import pdfplumber
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string


# ------------------------------------------------------------
#  RUTA BASE: funciona como .py y como .exe --onefile
# ------------------------------------------------------------
def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent


BASE_DIR = get_base_dir()

# ------------------------------------------------------------
#  CONFIGURACIÓN DE PLANTILLAS
# ------------------------------------------------------------
PLANTILLAS = {
    "Directivos": ("_plantillas/formulario_directivos.xlsm", "Formulario Directivos", 50),
    "Agentes (1)": ("_plantillas/formulario_agentes1.xlsm", "formulario agentes", 40),
    "Agentes (2)": ("_plantillas/formulario_agentes2.xlsm", "formulario agentes", 40),
}

MAPEO_AGENTES2 = {
    "PRIMER APELLIDO": "B6",
    "SEGUNDO APELLIDO": "B7",
    "NOMBRE": "B8",
    "TIPO DE DOCUMENTO": "B9",
    "Nº DE DOCUMENTO": "B10",
    "SEXO": "B11",
    "FECHA DE NACIMIENTO": "B12",
    "DIRECCION": "B13",
    "CIUDAD": "B14",
    "CODIGO POSTAL": "B15",
    "CCAA": "B16",
    "PROVINCIA": "B17",
    "TELÉFONO": "B18",
    "EMAIL": "B19",
    "RESIDE EN LOCALIDAD <5000": "B20",
    "PERSONA CON DISCAPACIDAD": "B21",
    "PERFIL LINKEDIN": "B22",
    "NIVEL DE ESTUDIOS": "B25",
    "FORMACIÓN DIGITALIZACIÓN": "B26",
    "FORMACIÓN GESTIÓN PROYECTOS": "B27",
    "AÑOS EXPERIENCIA LABORAL": "B30",
    "EXPERIENCIA DIGITALIZACIÓN": "B31",
    "SITUACIÓN LABORAL": "B32",
    "NOMBRE EMPRESA": "B34",
    "NIF EMPRESA": "B35",
    "DEPARTAMENTO EMPRESA": "B36",
    "PUESTO EMPRESA": "B37",
    "ACTIVIDAD EMPRESA": "B38",
    "TAMAÑO EMPRESA": "B39",
    "DIRECCIÓN EMPRESA": "B40",
    "CIUDAD EMPRESA": "B41",
    "CODIGO POSTAL EMPRESA": "B42",
    "CCAA EMPRESA": "B43",
    "PROVINCIA EMPRESA": "B44",
    "TELÉFONO EMPRESA": "B45",
    "WEB EMPRESA": "B46",
    "ANTIGÜEDAD EMPRESA": "B47",
    "FACTURACIÓN EMPRESA": "B48",
    "ÁMBITO RURAL EMPRESA": "B49",
    "MADUREZ DIGITAL EMPRESA": "B50",
    "SOSTENIBILIDAD EMPRESA": "B51",
    "PLAN DIGITAL EMPRESA": "B52",
    "MUJER DIRECTIVA EMPRESA": "B53",
    "PORCENTAJE MUJERES EMPRESA": "B54",
    "MOTIVACIÓN": "B57",
    "ACEPTO CONDICIONADO": "B79",
    "DECLARO PYME": "B80",
    "DECLARO REALIDAD": "B81",
    "DECLARO NO RECIBIDA": "B82",
    "DECLARO CONFLICTO": "B83",
    "AUTORIZO DATOS": "B84",
    "ACEPTO DISCAPACIDAD": "B85",
}

MAPEO_AGENTES1 = MAPEO_AGENTES2.copy()
MAPEO_AGENTES1["ACEPTO CONDICIONADO"] = "B80"

MAPEO_DIRECTIVOS = {
    # DATOS PERSONALES (Filas 6-23)
    "PRIMER APELLIDO": "B6",
    "SEGUNDO APELLIDO": "B7",
    "NOMBRE": "B8",
    "TIPO DE DOCUMENTO": "B9",
    "Nº DE DOCUMENTO": "B10",
    "SEXO": "B11",
    "FECHA DE NACIMIENTO": "B12",
    "DIRECCION": "B13",
    "CIUDAD": "B14",
    "CODIGO POSTAL": "B15",
    "CCAA": "B16",
    "PROVINCIA": "B17",
    "TELÉFONO": "B18",
    "EMAIL": "B19",
    "RESIDE EN LOCALIDAD <5000": "B20",
    "PERSONA CON DISCAPACIDAD": "B21",
    "NIVEL DE ESTUDIOS": "B22",
    "TITULACION": "B23",
    # DATOS PROFESIONALES (Filas 26-29)
    "NOMBRE EMPRESA": "B26",
    "RELACIÓN CON LA EMPRESA": "B27",
    "DEPARTAMENTO EMPRESA": "B28",
    "PUESTO EMPRESA": "B29",
    # DATOS DE LA EMPRESA (Filas 32-51)
    "NIF EMPRESA": "B32",
    "ACTIVIDAD EMPRESA": "B33",
    "TAMAÑO EMPRESA": "B34",
    "DIRECCIÓN EMPRESA": "B35",
    "CIUDAD EMPRESA": "B36",
    "CODIGO POSTAL EMPRESA": "B37",
    "CCAA EMPRESA": "B38",
    "PROVINCIA EMPRESA": "B39",
    "TELÉFONO EMPRESA": "B40",
    "WEB EMPRESA": "B41",
    "ANTIGÜEDAD EMPRESA": "B42",
    "FACTURACIÓN EMPRESA": "B43",
    "ÁMBITO RURAL EMPRESA": "B44",
    "MADUREZ DIGITAL EMPRESA": "B45",
    "CANALES RELACION EMPRESA": "B46",
    "PERFIL TIC EMPRESA": "B47",
    "SOSTENIBILIDAD EMPRESA": "B48",
    "PLAN DIGITAL EMPRESA": "B49",
    "MUJER DIRECTIVA EMPRESA": "B50",
    "PORCENTAJE MUJERES EMPRESA": "B51",
    # MOTIVACIÓN (Fila 54)
    "MOTIVACIÓN": "B54",
    # DECLARACIONES (Filas 85-91)
    "ACEPTO CONDICIONADO": "B85",
    "DECLARO PYME": "B86",
    "DECLARO REALIDAD": "B87",
    "DECLARO NO RECIBIDA": "B88",
    "DECLARO CONFLICTO": "B89",
    "AUTORIZO DATOS": "B90",
    "ACEPTO DISCAPACIDAD": "B91",
}


def _set_icon(ventana):
    ico = BASE_DIR / "icono.ico"
    if ico.exists():
        try:
            ventana.iconbitmap(str(ico))
        except Exception:
            pass


class VentanaLog:
    def __init__(self, titulo="Procesando PDFs"):
        self.ventana = tk.Toplevel()
        self.ventana.title(titulo)
        self.ventana.geometry("1000x700")
        self.ventana.transient()
        _set_icon(self.ventana)

        frame = tk.Frame(self.ventana)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.texto = scrolledtext.ScrolledText(
            frame, wrap=tk.WORD, width=120, height=35, font=("Consolas", 9)
        )
        self.texto.pack(fill=tk.BOTH, expand=True)

        btn_frame = tk.Frame(self.ventana)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        # Marca de agua
        tk.Label(
            self.ventana,
            text="Hecho por RaulRDA.com, Pablo Álvarez y Pelayo Fernández",
            font=("Arial", 8),
            fg="gray"
        ).pack(side=tk.BOTTOM, pady=2)

        self.btn_cerrar = tk.Button(
            btn_frame, text="Cerrar", command=self.cerrar, state=tk.DISABLED
        )
        self.btn_cerrar.pack(side=tk.RIGHT)

        self.ventana.protocol("WM_DELETE_WINDOW", self.cerrar)
        self.procesando = True

    def log(self, mensaje, nivel="INFO"):
        import time
        timestamp = time.strftime("%H:%M:%S")
        linea = f"[{timestamp}] {nivel}: {mensaje}\n"
        self.texto.insert(tk.END, linea)
        self.texto.see(tk.END)
        self.ventana.update()
        if sys.stdout:
            print(f"{nivel}: {mensaje}")

    def error(self, mensaje):
        self.log(mensaje, "ERROR")
        messagebox.showerror("Error", mensaje)

    def warning(self, mensaje):
        self.log(mensaje, "ADVERTENCIA")

    def info(self, mensaje):
        self.log(mensaje, "INFO")

    def success(self, mensaje):
        self.log(mensaje, "ÉXITO")

    def completado(self):
        self.procesando = False
        self.btn_cerrar.config(state=tk.NORMAL)
        self.info("=" * 70)
        self.success("PROCESO COMPLETADO")

    def cerrar(self):
        if not self.procesando:
            self.ventana.destroy()


def limpiar_texto(texto: str) -> str:
    if not texto or not isinstance(texto, str):
        return ""
    texto = texto.replace('\n', ' ').replace('\r', ' ')
    return re.sub(r'\s+', ' ', texto).strip()


def seleccionar_tipo_plantilla():
    ventana = tk.Toplevel()
    ventana.title("Seleccionar tipo de plantilla")
    ventana.geometry("350x180")
    ventana.transient()
    ventana.grab_set()
    _set_icon(ventana)

    tk.Label(ventana, text="Elija el tipo de plantilla:", font=("Arial", 11)).pack(pady=15)

    tipo_var = tk.StringVar()
    combobox = ttk.Combobox(
        ventana, textvariable=tipo_var,
        values=list(PLANTILLAS.keys()), state="readonly", width=25
    )
    combobox.pack(pady=5)
    combobox.current(0)

    resultado = []

    def aceptar():
        resultado.append(tipo_var.get())
        ventana.destroy()

    def cancelar():
        resultado.append(None)
        ventana.destroy()

    frame_botones = tk.Frame(ventana)
    frame_botones.pack(pady=15)
    tk.Button(frame_botones, text="Aceptar", command=aceptar, width=10).pack(side=tk.LEFT, padx=10)
    tk.Button(frame_botones, text="Cancelar", command=cancelar, width=10).pack(side=tk.RIGHT, padx=10)

    tk.Label(
        ventana,
        text="Hecho por RaulRDA.com, Pablo Álvarez y Pelayo Fernández",
        font=("Arial", 8),
        fg="gray"
    ).pack(side=tk.BOTTOM, pady=2)

    ventana.wait_window()
    return resultado[0] if resultado else None


def seleccionar_carpeta_salida():
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta donde guardar los resultados")
    return Path(carpeta) if carpeta else None


def verificar_hoja_excel(archivo_excel, nombre_hoja_buscar, log):
    try:
        wb = openpyxl.load_workbook(archivo_excel, keep_vba=True, read_only=True)
        hojas = wb.sheetnames
        wb.close()
        log.info(f"📊 Hojas disponibles: {', '.join(hojas)}")
        for hoja in hojas:
            if hoja.lower() == nombre_hoja_buscar.lower():
                log.info(f"✓ Hoja encontrada: '{hoja}'")
                return hoja
        for hoja in hojas:
            if nombre_hoja_buscar.lower() in hoja.lower():
                log.info(f"✓ Hoja encontrada: '{hoja}'")
                return hoja
        log.error(f"✗ No se encontró la hoja '{nombre_hoja_buscar}'")
        return None
    except Exception as e:
        log.error(f"Error al leer Excel: {e}")
        return None


def extraer_tablas_pdf(pdf_path: str, tipo: str, log) -> dict:
    """
    Extrae datos del PDF según el tipo de formulario (Directivos o Agentes).

    Diferencias entre formularios:
    - Directivos: TIPO DE DOCUMENTO y Nº DOCUMENTO están INVERTIDOS en el PDF.
    - Directivos: DEPARTAMENTO y PUESTO/CARGO no llevan "(empresa)" en la etiqueta.
    - Directivos: las declaraciones (Acepto) están en la columna 7, no en la 1.
    - Directivos: campos extra: TITULACION, RELACIÓN CON LA EMPRESA,
                  CANALES RELACION EMPRESA, PERFIL TIC EMPRESA.
    """
    es_directivos = (tipo == "Directivos")
    log.info("🔍 Extrayendo datos del PDF...")
    datos = {}
    todas_etiquetas = []
    ultima_clave = None

    with pdfplumber.open(pdf_path) as pdf:
        for pagina_num, pagina in enumerate(pdf.pages, 1):
            tablas = pagina.extract_tables()
            if not tablas:
                continue
            log.info(f"  Página {pagina_num}: {len(tablas)} tabla(s)")

            for tabla_num, tabla in enumerate(tablas, 1):
                log.info(f"    Tabla {tabla_num} ({len(tabla)} filas):")
                for fila in tabla:
                    if not fila or len(fila) < 2:
                        ultima_clave = None
                        continue

                    etiqueta_raw = limpiar_texto(fila[0]) if fila[0] else None

                    # ── Buscar valor: col 1 para agentes, col 1 o última no vacía para directivos ──
                    valor = limpiar_texto(fila[1]) if fila[1] else ""
                    if es_directivos and not valor:
                        # Las declaraciones de directivos tienen el valor en columnas más a la derecha
                        for celda in fila[2:]:
                            v = limpiar_texto(celda) if celda else ""
                            if v:
                                valor = v
                                break

                    # FIX agentes: fila sin etiqueta → FORMACIÓN GESTIÓN PROYECTOS
                    if not etiqueta_raw:
                        if ultima_clave == "FORMACIÓN DIGITALIZACIÓN" and valor:
                            log.info(f"      [FIX] → FORMACIÓN GESTIÓN PROYECTOS = '{valor}'")
                            datos["FORMACIÓN GESTIÓN PROYECTOS"] = valor
                        ultima_clave = None
                        continue

                    todas_etiquetas.append(etiqueta_raw)
                    log.info(f"      '{etiqueta_raw[:60]}' = '{valor[:50]}'")

                    if not valor:
                        ultima_clave = None
                        continue

                    e = etiqueta_raw.upper()

                    # ── DATOS PERSONALES ──────────────────────────────────────
                    if "PRIMER APELLIDO" in e:
                        datos["PRIMER APELLIDO"] = valor
                    elif "SEGUNDO APELLIDO" in e:
                        datos["SEGUNDO APELLIDO"] = valor
                    elif "NOMBRE" in e and "EMPRESA" not in e and "RAZÓN" not in e and "SOCIAL" not in e:
                        datos["NOMBRE"] = valor
                    elif "TIPO DE DOCUMENTO" in e:
                        # Directivos: el PDF pone el número en este campo y el tipo en el siguiente
                        if es_directivos:
                            datos["Nº DE DOCUMENTO"] = valor
                        else:
                            datos["TIPO DE DOCUMENTO"] = valor
                    elif "Nº" in e and "DOCUMENTO" in e:
                        # Directivos: el PDF pone el tipo (NIE/NIF) en este campo
                        if es_directivos:
                            datos["TIPO DE DOCUMENTO"] = valor
                        else:
                            datos["Nº DE DOCUMENTO"] = valor
                    elif "SEXO" in e:
                        datos["SEXO"] = valor
                    elif "FECHA DE NACIMIENTO" in e:
                        datos["FECHA DE NACIMIENTO"] = valor
                    elif "DIRECCION" in e and "EMPRESA" not in e:
                        datos["DIRECCION"] = valor
                    elif "CIUDAD" in e and "EMPRESA" not in e:
                        datos["CIUDAD"] = valor
                    elif "CODIGO POSTAL" in e and "EMPRESA" not in e:
                        datos["CODIGO POSTAL"] = valor
                    elif "CCAA" in e and "EMPRESA" not in e:
                        datos["CCAA"] = valor
                    elif "PROVINCIA" in e and "EMPRESA" not in e:
                        datos["PROVINCIA"] = valor
                    elif "TELÉFONO" in e and "EMPRESA" not in e:
                        datos["TELÉFONO"] = valor
                    elif "EMAIL" in e:
                        datos["EMAIL"] = valor
                    elif "RESIDE EN UNA LOCALIDAD" in e or "HABITANTES INFERIOR" in e:
                        datos["RESIDE EN LOCALIDAD <5000"] = valor
                    elif "PERSONA CON DISCAPACIDAD" in e:
                        datos["PERSONA CON DISCAPACIDAD"] = valor
                    elif "NIVEL DE ESTUDIOS" in e:
                        datos["NIVEL DE ESTUDIOS"] = valor

                    # ── SOLO DIRECTIVOS ───────────────────────────────────────
                    elif "TITULACION" in e:
                        datos["TITULACION"] = valor
                    elif "RELACIÓN CON LA EMPRESA" in e:
                        datos["RELACIÓN CON LA EMPRESA"] = valor

                    # ── SOLO AGENTES ──────────────────────────────────────────
                    elif "LINKEDIN" in e:
                        datos["PERFIL LINKEDIN"] = valor
                    elif "COMPLEMENTARIA" in e and "DIGITALIZACIÓN" in e and "GESTIÓN" not in e:
                        datos["FORMACIÓN DIGITALIZACIÓN"] = valor
                        ultima_clave = "FORMACIÓN DIGITALIZACIÓN"
                        continue
                    elif "GESTIÓN" in e and "PROYECTOS" in e:
                        datos["FORMACIÓN GESTIÓN PROYECTOS"] = valor
                    elif "AÑOS DE EXPERIENCIA LABORAL" in e:
                        datos["AÑOS EXPERIENCIA LABORAL"] = valor
                    elif "EXPERIENCIA LABORAL EN PUESTOS" in e or "EXPERIENCIA EN DIGITALIZACIÓN" in e:
                        datos["EXPERIENCIA DIGITALIZACIÓN"] = valor
                    elif "SITUACIÓN LABORAL ACTUAL" in e:
                        datos["SITUACIÓN LABORAL"] = valor

                    # ── DATOS EMPRESA ─────────────────────────────────────────
                    elif "NOMBRE EMPRESA" in e or ("NOMBRE" in e and "RAZÓN SOCIAL" in e):
                        datos["NOMBRE EMPRESA"] = valor
                    elif "NIF EMPRESA" in e or e.strip() == "NIF":
                        datos["NIF EMPRESA"] = valor
                    elif "DEPARTAMENTO" in e:
                        # Directivos: etiqueta es solo "DEPARTAMENTO" (sin "EMPRESA")
                        datos["DEPARTAMENTO EMPRESA"] = valor
                    elif "PUESTO" in e or "CARGO" in e:
                        # Directivos: etiqueta es solo "PUESTO/CARGO" (sin "EMPRESA")
                        datos["PUESTO EMPRESA"] = valor
                    elif "ACTIVIDAD DE LA EMPRESA" in e:
                        datos["ACTIVIDAD EMPRESA"] = valor
                    elif "TAMAÑO EMPRESA" in e:
                        datos["TAMAÑO EMPRESA"] = valor
                    elif "DIRECCIÓN" in e and "EMPRESA" in e:
                        datos["DIRECCIÓN EMPRESA"] = valor
                    elif "CIUDAD" in e and "EMPRESA" in e:
                        datos["CIUDAD EMPRESA"] = valor
                    elif "CODIGO POSTAL" in e and "EMPRESA" in e:
                        datos["CODIGO POSTAL EMPRESA"] = valor
                    elif "CCAA" in e and "EMPRESA" in e:
                        datos["CCAA EMPRESA"] = valor
                    elif "PROVINCIA" in e and "EMPRESA" in e:
                        datos["PROVINCIA EMPRESA"] = valor
                    elif "TELÉFONO" in e and "EMPRESA" in e:
                        datos["TELÉFONO EMPRESA"] = valor
                    elif "PAGINA WEB" in e:
                        datos["WEB EMPRESA"] = valor
                    elif "ANTIGÜEDAD DE LA EMPRESA" in e:
                        datos["ANTIGÜEDAD EMPRESA"] = valor
                    elif "FACTURACIÓN ÚLTIMO AÑO" in e:
                        datos["FACTURACIÓN EMPRESA"] = valor
                    elif "AMBITO RURAL" in e:
                        datos["ÁMBITO RURAL EMPRESA"] = valor
                    elif "NIVEL DE MADUREZ DIGITAL" in e:
                        datos["MADUREZ DIGITAL EMPRESA"] = valor
                    elif "CANALES DE RELACION DE LA EMPRESA" in e:
                        datos["CANALES RELACION EMPRESA"] = valor
                    elif "PROFESIONALES CON PERFIL TIC" in e:
                        datos["PERFIL TIC EMPRESA"] = valor
                    elif "POLITICAS DE SOSTENIBILIDAD" in e:
                        datos["SOSTENIBILIDAD EMPRESA"] = valor
                    elif "POLÍTICAS O PLANES" in e or "PLANES DE TRANSFORMACIÓN DIGITAL" in e:
                        datos["PLAN DIGITAL EMPRESA"] = valor
                    elif "MÁXIMA RESPONSABLE" in e or "EQUIPO DIRECTIVO ES MUJER" in e:
                        datos["MUJER DIRECTIVA EMPRESA"] = valor
                    elif "PORCENTAJE DE MUJERES" in e:
                        datos["PORCENTAJE MUJERES EMPRESA"] = valor

                    # ── MOTIVACIÓN ────────────────────────────────────────────
                    elif "DESCRIBIR MOTIVACIÓN" in e or (e == "MOTIVACIÓN"):
                        datos["MOTIVACIÓN"] = valor

                    # ── DECLARACIONES ─────────────────────────────────────────
                    elif "ACEPTO LOS TÉRMINOS" in e or ("ACEPTO" in e and "TÉRMINOS" in e):
                        datos["ACEPTO CONDICIONADO"] = valor
                    elif "EMPRESA DONDE" in e or "SOY DIRECTIVO" in e or ("DECLARO" in e and "PYME" in e):
                        datos["DECLARO PYME"] = valor
                    elif "DECLARO RESPONSABLEMENTE QUE TODA" in e or ("INFORMACIÓN" in e and "SOLICITUD" in e and "DECLARO" in e):
                        datos["DECLARO REALIDAD"] = valor
                    elif "NO HE RECIBIDO" in e:
                        datos["DECLARO NO RECIBIDA"] = valor
                    elif "NO ME ENCUENTRO INCURSO" in e or "CONFLICTO DE INTERÉS" in e:
                        datos["DECLARO CONFLICTO"] = valor
                    elif "AUTORIZO TRATAMIENTO" in e or ("AUTORIZO" in e and "DATOS" in e):
                        datos["AUTORIZO DATOS"] = valor
                    elif "DATO RELATIVO A DISCAPACIDAD" in e:
                        datos["ACEPTO DISCAPACIDAD"] = valor

                    ultima_clave = None

    log.info("\n  📋 ETIQUETAS ÚNICAS ENCONTRADAS:")
    for etiqueta in sorted(set(todas_etiquetas)):
        log.info(f"    - {etiqueta}")
    log.info(f"\n  Total campos extraídos: {len(datos)}")
    return datos


def escribir_valor_celda(ws, coordenada, valor, log):
    try:
        for merged_range in ws.merged_cells.ranges:
            if coordenada in merged_range:
                col_letter, row_num = coordinate_from_string(coordenada)
                col_idx = column_index_from_string(col_letter)
                if col_idx == merged_range.min_col and row_num == merged_range.min_row:
                    break
                else:
                    log.info(f"      ⚠ Celda {coordenada} combinada - omitida")
                    return False
        ws[coordenada] = str(valor).strip()
        return True
    except Exception as ex:
        log.error(f"      Error escribiendo en {coordenada}: {str(ex)}")
        return False


def rellenar_excel(mapa_datos, mapeo_celdas, plantilla, salida, nombre_hoja, log):
    log.info("📝 Generando archivo Excel...")
    shutil.copy2(plantilla, salida)
    wb = openpyxl.load_workbook(salida, keep_vba=True)

    if nombre_hoja not in wb.sheetnames:
        log.error(f"Hoja '{nombre_hoja}' no encontrada. Disponibles: {wb.sheetnames}")
        return 0

    ws = wb[nombre_hoja]
    celdas_ok = 0
    celdas_omitidas = 0

    log.info("  Escribiendo en Excel:")
    for campo, celda in mapeo_celdas.items():
        if campo in mapa_datos and mapa_datos[campo]:
            valor = mapa_datos[campo]
            log.info(f"    {campo} -> {celda}: '{valor[:30]}'")
            if escribir_valor_celda(ws, celda, valor, log):
                celdas_ok += 1
            else:
                celdas_omitidas += 1

    wb.save(salida)
    log.info(f"  Celdas rellenadas: {celdas_ok}")
    if celdas_omitidas > 0:
        log.warning(f"  Celdas omitidas (combinadas): {celdas_omitidas}")
    return celdas_ok


def obtener_nombre_archivo(mapa_datos):
    nombre = mapa_datos.get("NOMBRE", "")
    apellido1 = mapa_datos.get("PRIMER APELLIDO", "")
    partes = [p for p in [apellido1, nombre] if p]
    nombre_completo = "_".join(partes) if partes else "SinNombre"
    nombre_completo = re.sub(r'[\\/*?:"<>|]', "", nombre_completo)
    return f"Formulario {nombre_completo}.xlsm"


def main():
    # FIX: en modo --windowed el exe no tiene consola y sys.stdout es None
    if sys.stdout and hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    root = tk.Tk()
    root.withdraw()
    _set_icon(root)

    # 1. Tipo de plantilla
    tipo = seleccionar_tipo_plantilla()
    if not tipo:
        messagebox.showinfo("Info", "Operación cancelada")
        return

    archivo_plantilla, nombre_hoja_buscar, limite = PLANTILLAS[tipo]
    mapeo = MAPEO_DIRECTIVOS if tipo == "Directivos" else (MAPEO_AGENTES2 if tipo == "Agentes (2)" else MAPEO_AGENTES1)

    ruta_plantilla = BASE_DIR / archivo_plantilla
    if not ruta_plantilla.exists():
        messagebox.showerror("Error", f"No se encuentra la plantilla:\n{ruta_plantilla}")
        return

    # 2. Seleccionar PDFs
    pdfs = filedialog.askopenfilenames(
        title=f"Selecciona hasta {limite} PDFs",
        filetypes=[("PDF", "*.pdf")]
    )
    if not pdfs:
        return
    if len(pdfs) > limite:
        messagebox.showerror("Error", f"Máximo {limite} PDFs para este tipo")
        return

    # 3. Carpeta de salida
    carpeta_salida = seleccionar_carpeta_salida()
    if not carpeta_salida:
        messagebox.showinfo("Info", "Operación cancelada: no se seleccionó carpeta de destino")
        return

    log = VentanaLog(f"Procesando - {tipo}")
    log.info(f"📂 Plantilla:       {ruta_plantilla}")
    log.info(f"📁 Carpeta salida:  {carpeta_salida}")
    log.info(f"📑 PDFs:            {len(pdfs)}")

    nombre_hoja = verificar_hoja_excel(str(ruta_plantilla), nombre_hoja_buscar, log)
    if not nombre_hoja:
        return

    exitosos = 0
    for i, pdf in enumerate(pdfs, 1):
        pdf_path = Path(pdf)
        log.info("=" * 70)
        log.info(f"📄 {i}/{len(pdfs)}: {pdf_path.name}")

        try:
            datos = extraer_tablas_pdf(str(pdf_path), tipo, log)
            if not datos:
                log.error("No se extrajeron datos")
                continue

            log.info("\n📊 DATOS EXTRAÍDOS:")
            for campo in sorted(datos.keys()):
                log.info(f"  {campo}: {datos[campo][:50]}")

            nombre_salida = obtener_nombre_archivo(datos)
            ruta_salida = carpeta_salida / nombre_salida
            contador = 1
            while ruta_salida.exists():
                nombre_base = ruta_salida.stem
                ruta_salida = carpeta_salida / f"{nombre_base}_{contador}.xlsm"
                contador += 1

            celdas = rellenar_excel(datos, mapeo, str(ruta_plantilla),
                                    str(ruta_salida), nombre_hoja, log)
            log.success(f"✅ Generado: {ruta_salida.name} ({celdas} celdas)")
            exitosos += 1

        except Exception as ex:
            log.error(f"Error: {ex}")
            import traceback
            traceback.print_exc()

    log.info("=" * 70)
    log.success(f"✅ Completado: {exitosos}/{len(pdfs)}")
    log.completado()
    messagebox.showinfo("Éxito", f"Procesados: {exitosos}/{len(pdfs)}\nGuardados en:\n{carpeta_salida}")


if __name__ == "__main__":
    main()