"""
╔════════════════════════════════════════════════════════════╗
║   PDF Formulario Agentes / Directivos  →  Excel            ║
║   Grupo ATU © 2026                                         ║
║   Hecho por RaulRDA.com, Pablo Álvarez y Pelayo Fernández  ║
╚════════════════════════════════════════════════════════════╝
"""
# INSTALACIÓN:
#? pip install pdfplumber openpyxl

# USO:
#? python main.py

# CONVERTIR A EXE (opcional):
#? py -m PyInstaller --onefile --windowed --name "PDF_a_Excel" --icon="icono.ico" --add-data "_plantillas;_plantillas" --add-data "icono.ico;." --collect-all pdfplumber --collect-all pdfminer main.py

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


# ── RUTA BASE ────────────────────────────────────────────────
def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent

BASE_DIR = get_base_dir()

# ── PALETA ATU ───────────────────────────────────────────────
ATU_NAVY    = "#1C2B4A"
ATU_ORANGE  = "#E8721C"
ATU_LIGHT   = "#F5F7FA"
ATU_WHITE   = "#FFFFFF"
ATU_TEXT    = "#2D3748"
ATU_MID     = "#4A6FA5"
ATU_SUCCESS = "#2E7D32"
ATU_ERROR   = "#C62828"
ATU_WARN    = "#E65100"
ATU_BORDER  = "#CBD5E0"

# ── PLANTILLAS ───────────────────────────────────────────────
PLANTILLAS = {
    "Directivos":  ("_plantillas/formulario_directivos.xlsm", "Formulario Directivos", 50),
    "Agentes (1)": ("_plantillas/formulario_agentes1.xlsm",   "formulario agentes",    40),
    "Agentes (2)": ("_plantillas/formulario_agentes2.xlsm",   "formulario agentes",    40),
}

MAPEO_AGENTES2 = {
    "PRIMER APELLIDO": "B6", "SEGUNDO APELLIDO": "B7", "NOMBRE": "B8",
    "TIPO DE DOCUMENTO": "B9", "Nº DE DOCUMENTO": "B10", "SEXO": "B11",
    "FECHA DE NACIMIENTO": "B12", "DIRECCION": "B13", "CIUDAD": "B14",
    "CODIGO POSTAL": "B15", "CCAA": "B16", "PROVINCIA": "B17",
    "TELÉFONO": "B18", "EMAIL": "B19", "RESIDE EN LOCALIDAD <5000": "B20",
    "PERSONA CON DISCAPACIDAD": "B21", "PERFIL LINKEDIN": "B22",
    "NIVEL DE ESTUDIOS": "B25", "FORMACIÓN DIGITALIZACIÓN": "B26",
    "FORMACIÓN GESTIÓN PROYECTOS": "B27", "AÑOS EXPERIENCIA LABORAL": "B30",
    "EXPERIENCIA DIGITALIZACIÓN": "B31", "SITUACIÓN LABORAL": "B32",
    "NOMBRE EMPRESA": "B34", "NIF EMPRESA": "B35",
    "DEPARTAMENTO EMPRESA": "B36", "PUESTO EMPRESA": "B37",
    "ACTIVIDAD EMPRESA": "B38", "TAMAÑO EMPRESA": "B39",
    "DIRECCIÓN EMPRESA": "B40", "CIUDAD EMPRESA": "B41",
    "CODIGO POSTAL EMPRESA": "B42", "CCAA EMPRESA": "B43",
    "PROVINCIA EMPRESA": "B44", "TELÉFONO EMPRESA": "B45",
    "WEB EMPRESA": "B46", "ANTIGÜEDAD EMPRESA": "B47",
    "FACTURACIÓN EMPRESA": "B48", "ÁMBITO RURAL EMPRESA": "B49",
    "MADUREZ DIGITAL EMPRESA": "B50", "SOSTENIBILIDAD EMPRESA": "B51",
    "PLAN DIGITAL EMPRESA": "B52", "MUJER DIRECTIVA EMPRESA": "B53",
    "PORCENTAJE MUJERES EMPRESA": "B54", "MOTIVACIÓN": "B57",
    "ACEPTO CONDICIONADO": "B79", "DECLARO PYME": "B80",
    "DECLARO REALIDAD": "B81", "DECLARO NO RECIBIDA": "B82",
    "DECLARO CONFLICTO": "B83", "AUTORIZO DATOS": "B84",
    "ACEPTO DISCAPACIDAD": "B85",
}

MAPEO_AGENTES1 = MAPEO_AGENTES2.copy()
MAPEO_AGENTES1.pop("DECLARO PYME")
MAPEO_AGENTES1["ACEPTO CONDICIONADO"] = "B80"

MAPEO_DIRECTIVOS = {
    "PRIMER APELLIDO": "B6", "SEGUNDO APELLIDO": "B7", "NOMBRE": "B8",
    "TIPO DE DOCUMENTO": "B9", "Nº DE DOCUMENTO": "B10", "SEXO": "B11",
    "FECHA DE NACIMIENTO": "B12", "DIRECCION": "B13", "CIUDAD": "B14",
    "CODIGO POSTAL": "B15", "CCAA": "B16", "PROVINCIA": "B17",
    "TELÉFONO": "B18", "EMAIL": "B19", "RESIDE EN LOCALIDAD <5000": "B20",
    "PERSONA CON DISCAPACIDAD": "B21", "NIVEL DE ESTUDIOS": "B22",
    "TITULACION": "B23", "NOMBRE EMPRESA": "B26",
    "RELACIÓN CON LA EMPRESA": "B27", "DEPARTAMENTO EMPRESA": "B28",
    "PUESTO EMPRESA": "B29", "NIF EMPRESA": "B32",
    "ACTIVIDAD EMPRESA": "B33", "TAMAÑO EMPRESA": "B34",
    "DIRECCIÓN EMPRESA": "B35", "CIUDAD EMPRESA": "B36",
    "CODIGO POSTAL EMPRESA": "B37", "CCAA EMPRESA": "B38",
    "PROVINCIA EMPRESA": "B39", "TELÉFONO EMPRESA": "B40",
    "WEB EMPRESA": "B41", "ANTIGÜEDAD EMPRESA": "B42",
    "FACTURACIÓN EMPRESA": "B43", "ÁMBITO RURAL EMPRESA": "B44",
    "MADUREZ DIGITAL EMPRESA": "B45", "CANALES RELACION EMPRESA": "B46",
    "PERFIL TIC EMPRESA": "B47", "SOSTENIBILIDAD EMPRESA": "B48",
    "PLAN DIGITAL EMPRESA": "B49", "MUJER DIRECTIVA EMPRESA": "B50",
    "PORCENTAJE MUJERES EMPRESA": "B51", "MOTIVACIÓN": "B54",
    "ACEPTO CONDICIONADO": "B85", "DECLARO PYME": "B86",
    "DECLARO REALIDAD": "B87", "DECLARO NO RECIBIDA": "B88",
    "DECLARO CONFLICTO": "B89", "AUTORIZO DATOS": "B90",
    "ACEPTO DISCAPACIDAD": "B91",
}


# ── UTILIDADES ───────────────────────────────────────────────
def _set_icon(ventana):
    ico = BASE_DIR / "icono.ico"
    if ico.exists():
        try:
            ventana.iconbitmap(str(ico))
        except Exception:
            pass

def limpiar_texto(texto: str) -> str:
    if not texto or not isinstance(texto, str):
        return ""
    return re.sub(r'\s+', ' ', texto.replace('\n', ' ').replace('\r', ' ')).strip()


# ── VENTANA SELECTOR ─────────────────────────────────────────
def seleccionar_tipo_plantilla():
    ventana = tk.Toplevel()
    ventana.title("Grupo ATU — Conversor PDF → Excel")
    ventana.geometry("420x280")
    ventana.resizable(False, False)
    ventana.configure(bg=ATU_NAVY)
    ventana.transient()
    ventana.grab_set()
    _set_icon(ventana)

    # Header
    header = tk.Frame(ventana, bg=ATU_NAVY, pady=18)
    header.pack(fill=tk.X)
    tk.Label(header, text="GRUPO ATU", font=("Arial", 18, "bold"),
             bg=ATU_NAVY, fg=ATU_WHITE).pack()
    tk.Label(header, text="Conversor de formularios PDF a Excel",
             font=("Arial", 9), bg=ATU_NAVY, fg=ATU_MID).pack()

    # Separador naranja
    tk.Frame(ventana, bg=ATU_ORANGE, height=3).pack(fill=tk.X)

    # Cuerpo
    body = tk.Frame(ventana, bg=ATU_LIGHT, pady=20)
    body.pack(fill=tk.BOTH, expand=True)

    tk.Label(body, text="Selecciona el tipo de plantilla:",
             font=("Arial", 10, "bold"), bg=ATU_LIGHT, fg=ATU_TEXT).pack(pady=(0, 8))

    tipo_var = tk.StringVar()
    style = ttk.Style()
    style.configure("ATU.TCombobox", fieldbackground=ATU_WHITE, background=ATU_WHITE)
    combobox = ttk.Combobox(body, textvariable=tipo_var,
                            values=list(PLANTILLAS.keys()),
                            state="readonly", width=28, font=("Arial", 10))
    combobox.pack(pady=4)
    combobox.current(0)

    resultado = []

    def aceptar():
        resultado.append(tipo_var.get())
        ventana.destroy()

    def cancelar():
        resultado.append(None)
        ventana.destroy()

    btn_frame = tk.Frame(body, bg=ATU_LIGHT)
    btn_frame.pack(pady=14)

    tk.Button(btn_frame, text="  Aceptar  ", command=aceptar,
              bg=ATU_ORANGE, fg=ATU_WHITE, font=("Arial", 10, "bold"),
              relief=tk.FLAT, cursor="hand2", padx=10, pady=6).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_frame, text="  Cancelar  ", command=cancelar,
              bg=ATU_BORDER, fg=ATU_TEXT, font=("Arial", 10),
              relief=tk.FLAT, cursor="hand2", padx=10, pady=6).pack(side=tk.LEFT, padx=8)

    # Pie
    tk.Frame(ventana, bg=ATU_ORANGE, height=3).pack(fill=tk.X, side=tk.BOTTOM)
    tk.Label(ventana, text="Hecho por RaulRDA.com · Pablo Álvarez · Pelayo Fernández",
             font=("Arial", 7), bg=ATU_NAVY, fg=ATU_MID).pack(side=tk.BOTTOM, pady=4)

    ventana.wait_window()
    return resultado[0] if resultado else None


def seleccionar_carpeta_salida():
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta de destino")
    return Path(carpeta) if carpeta else None


# ── VENTANA LOG ──────────────────────────────────────────────
class VentanaLog:
    def __init__(self, titulo="Procesando"):
        self.ventana = tk.Toplevel()
        self.ventana.title(f"Grupo ATU — {titulo}")
        self.ventana.geometry("860x600")
        self.ventana.configure(bg=ATU_NAVY)
        self.ventana.transient()
        _set_icon(self.ventana)

        # Header
        header = tk.Frame(self.ventana, bg=ATU_NAVY, pady=10, padx=16)
        header.pack(fill=tk.X)
        tk.Label(header, text="GRUPO ATU", font=("Arial", 14, "bold"),
                 bg=ATU_NAVY, fg=ATU_WHITE).pack(side=tk.LEFT)
        self.lbl_estado = tk.Label(header, text="⏳ Procesando...",
                                   font=("Arial", 9), bg=ATU_NAVY, fg=ATU_ORANGE)
        self.lbl_estado.pack(side=tk.RIGHT)

        tk.Frame(self.ventana, bg=ATU_ORANGE, height=3).pack(fill=tk.X)

        # Área de log
        log_frame = tk.Frame(self.ventana, bg=ATU_LIGHT, padx=12, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.texto = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=("Consolas", 9),
            bg="#1E2433", fg="#C8D0E0", insertbackground=ATU_WHITE,
            relief=tk.FLAT, borderwidth=0, padx=8, pady=6
        )
        self.texto.pack(fill=tk.BOTH, expand=True)

        # Tags de color para los distintos niveles
        self.texto.tag_config("INFO",    foreground="#8BA7C7")
        self.texto.tag_config("OK",      foreground="#4CAF50")
        self.texto.tag_config("ERROR",   foreground="#EF5350")
        self.texto.tag_config("WARN",    foreground="#FF9800")
        self.texto.tag_config("TITLE",   foreground=ATU_ORANGE, font=("Consolas", 9, "bold"))
        self.texto.tag_config("ARCHIVO", foreground="#E0E0E0", font=("Consolas", 9, "bold"))

        # Barra inferior
        tk.Frame(self.ventana, bg=ATU_ORANGE, height=3).pack(fill=tk.X)
        bottom = tk.Frame(self.ventana, bg=ATU_NAVY, pady=8, padx=12)
        bottom.pack(fill=tk.X)

        tk.Label(bottom, text="Hecho por RaulRDA.com · Pablo Álvarez · Pelayo Fernández",
                 font=("Arial", 7), bg=ATU_NAVY, fg=ATU_MID).pack(side=tk.LEFT)

        self.btn_cerrar = tk.Button(
            bottom, text="  Cerrar  ", command=self.cerrar,
            bg=ATU_BORDER, fg=ATU_TEXT, font=("Arial", 9),
            relief=tk.FLAT, cursor="hand2", pady=4, state=tk.DISABLED
        )
        self.btn_cerrar.pack(side=tk.RIGHT)

        self.ventana.protocol("WM_DELETE_WINDOW", self.cerrar)
        self.procesando = True

    def _escribir(self, texto, tag="INFO"):
        self.texto.insert(tk.END, texto + "\n", tag)
        self.texto.see(tk.END)
        self.ventana.update()
        if sys.stdout:
            print(texto)

    def info(self, msg):    self._escribir(f"  {msg}", "INFO")
    def ok(self, msg):      self._escribir(f"✓ {msg}", "OK")
    def error(self, msg):   self._escribir(f"✗ {msg}", "ERROR")
    def warning(self, msg): self._escribir(f"⚠ {msg}", "WARN")
    def titulo(self, msg):  self._escribir(f"\n{'─'*60}\n  {msg}", "TITLE")
    def archivo(self, msg): self._escribir(f"  📄 {msg}", "ARCHIVO")

    # Alias de compatibilidad
    def log(self, msg, nivel="INFO"): self.info(msg)
    def success(self, msg): self.ok(msg)

    def completado(self, fallidos: list):
        self.procesando = False
        self.lbl_estado.config(text="✅ Completado", fg="#4CAF50")
        self._escribir(f"\n{'═'*60}", "TITLE")
        self._escribir("  PROCESO COMPLETADO", "OK")
        if fallidos:
            self._escribir(f"\n  ⚠ No se pudieron convertir estos archivos ({len(fallidos)}):", "WARN")
            for f in fallidos:
                self._escribir(f"    • {f}", "ERROR")
        self._escribir("", "INFO")
        self.btn_cerrar.config(state=tk.NORMAL, bg=ATU_ORANGE, fg=ATU_WHITE,
                               font=("Arial", 9, "bold"))

    def cerrar(self):
        if not self.procesando:
            self.ventana.destroy()


# ── LÓGICA EXCEL ─────────────────────────────────────────────
def verificar_hoja_excel(archivo_excel, nombre_hoja_buscar, log):
    try:
        wb = openpyxl.load_workbook(archivo_excel, keep_vba=True, read_only=True)
        hojas = wb.sheetnames
        wb.close()
        for hoja in hojas:
            if hoja.lower() == nombre_hoja_buscar.lower():
                return hoja
        for hoja in hojas:
            if nombre_hoja_buscar.lower() in hoja.lower():
                return hoja
        log.error(f"Hoja '{nombre_hoja_buscar}' no encontrada en la plantilla")
        return None
    except Exception as e:
        log.error(f"Error al leer Excel: {e}")
        return None


def escribir_valor_celda(ws, coordenada, valor, log):
    try:
        for merged_range in ws.merged_cells.ranges:
            if coordenada in merged_range:
                col_letter, row_num = coordinate_from_string(coordenada)
                col_idx = column_index_from_string(col_letter)
                if col_idx == merged_range.min_col and row_num == merged_range.min_row:
                    break
                else:
                    return False
        ws[coordenada] = str(valor).strip()
        return True
    except Exception:
        return False


def rellenar_excel(mapa_datos, mapeo_celdas, plantilla, salida, nombre_hoja, log):
    shutil.copy2(plantilla, salida)
    wb = openpyxl.load_workbook(salida, keep_vba=True)
    if nombre_hoja not in wb.sheetnames:
        log.error(f"Hoja '{nombre_hoja}' no encontrada")
        return 0
    ws = wb[nombre_hoja]
    celdas_ok = 0
    for campo, celda in mapeo_celdas.items():
        if campo in mapa_datos and mapa_datos[campo]:
            if escribir_valor_celda(ws, celda, mapa_datos[campo], log):
                celdas_ok += 1
    wb.save(salida)
    return celdas_ok


# ── EXTRACCIÓN PDF ───────────────────────────────────────────
def extraer_tablas_pdf(pdf_path: str, tipo: str, log) -> dict:
    es_directivos = (tipo == "Directivos")
    datos = {}
    ultima_clave = None
    n_tablas = 0

    with pdfplumber.open(pdf_path) as pdf:
        for pagina_num, pagina in enumerate(pdf.pages, 1):
            tablas = pagina.extract_tables()
            if not tablas:
                continue
            n_tablas += len(tablas)

            for tabla in tablas:
                for fila in tabla:
                    if not fila or len(fila) < 2:
                        ultima_clave = None
                        continue

                    etiqueta_raw = limpiar_texto(fila[0]) if fila[0] else None
                    valor = limpiar_texto(fila[1]) if fila[1] else ""
                    if not valor:
                        for celda in fila[2:]:
                            v = limpiar_texto(celda) if celda else ""
                            if v:
                                valor = v
                                break

                    if not etiqueta_raw:
                        if ultima_clave == "FORMACIÓN DIGITALIZACIÓN" and valor:
                            datos["FORMACIÓN GESTIÓN PROYECTOS"] = valor
                        ultima_clave = None
                        continue

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
                        datos["Nº DE DOCUMENTO" if es_directivos else "TIPO DE DOCUMENTO"] = valor
                    elif "Nº" in e and "DOCUMENTO" in e:
                        datos["TIPO DE DOCUMENTO" if es_directivos else "Nº DE DOCUMENTO"] = valor
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
                        datos["DEPARTAMENTO EMPRESA"] = valor
                    elif "PUESTO" in e or "CARGO" in e:
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
                    elif "DESCRIBIR MOTIVACIÓN" in e or e == "MOTIVACIÓN":
                        datos["MOTIVACIÓN"] = valor

                    # ── DECLARACIONES (solo Agentes1 puede leer DECLARO PYME del PDF) ──
                    elif "ACEPTO LOS TÉRMINOS" in e or ("ACEPTO" in e and "TÉRMINOS" in e):
                        if tipo == "Agentes (1)": datos["ACEPTO CONDICIONADO"] = valor
                    elif "EMPRESA DONDE" in e or "SOY DIRECTIVO" in e or ("DECLARO" in e and "PYME" in e):
                        if tipo == "Agentes (1)": datos["DECLARO PYME"] = valor
                    elif "DECLARO RESPONSABLEMENTE QUE TODA" in e:
                        if tipo == "Agentes (1)": datos["DECLARO REALIDAD"] = valor
                    elif "NO HE RECIBIDO" in e:
                        if tipo == "Agentes (1)": datos["DECLARO NO RECIBIDA"] = valor
                    elif "NO ME ENCUENTRO INCURSO" in e or "CONFLICTO DE INTERÉS" in e:
                        if tipo == "Agentes (1)": datos["DECLARO CONFLICTO"] = valor
                    elif "AUTORIZO TRATAMIENTO" in e or ("AUTORIZO" in e and "DATOS" in e):
                        if tipo == "Agentes (1)": datos["AUTORIZO DATOS"] = valor
                    elif "DATO RELATIVO A DISCAPACIDAD" in e:
                        if tipo == "Agentes (1)": datos["ACEPTO DISCAPACIDAD"] = valor

                    ultima_clave = None

    # ── FORZAR DECLARACIONES ──────────────────────────────────
    SIEMPRE_ACEPTO = ["ACEPTO CONDICIONADO", "DECLARO REALIDAD", "DECLARO NO RECIBIDA",
                      "DECLARO CONFLICTO", "AUTORIZO DATOS", "ACEPTO DISCAPACIDAD"]
    if tipo != "Agentes (1)":
        for campo in SIEMPRE_ACEPTO + ["DECLARO PYME"]:
            datos[campo] = "Acepto"
    else:
        for campo in SIEMPRE_ACEPTO:
            datos[campo] = "Acepto"
        datos.setdefault("DECLARO PYME", "Acepto")

    log.info(f"Tablas leídas: {n_tablas} | Campos extraídos: {len(datos)}")
    return datos


def obtener_nombre_archivo(mapa_datos):
    apellido1 = mapa_datos.get("PRIMER APELLIDO", "")
    nombre    = mapa_datos.get("NOMBRE", "")
    partes = [p for p in [apellido1, nombre] if p]
    nombre_completo = "_".join(partes) if partes else "SinNombre"
    nombre_completo = re.sub(r'[\\/*?:"<>|]', "", nombre_completo)
    return f"Formulario {nombre_completo}.xlsm"


# ── MAIN ─────────────────────────────────────────────────────
def main():
    if sys.stdout and hasattr(sys.stdout, 'buffer'):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

    root = tk.Tk()
    root.withdraw()
    _set_icon(root)

    tipo = seleccionar_tipo_plantilla()
    if not tipo:
        return

    archivo_plantilla, nombre_hoja_buscar, limite = PLANTILLAS[tipo]
    mapeo = (MAPEO_DIRECTIVOS if tipo == "Directivos"
             else MAPEO_AGENTES2 if tipo == "Agentes (2)"
             else MAPEO_AGENTES1)

    ruta_plantilla = BASE_DIR / archivo_plantilla
    if not ruta_plantilla.exists():
        messagebox.showerror("Error", f"No se encuentra la plantilla:\n{ruta_plantilla}")
        return

    pdfs = filedialog.askopenfilenames(
        title=f"Selecciona PDFs ({tipo} — máx. {limite})",
        filetypes=[("PDF", "*.pdf")]
    )
    if not pdfs:
        return
    if len(pdfs) > limite:
        messagebox.showerror("Error", f"Máximo {limite} PDFs para '{tipo}'")
        return

    carpeta_salida = seleccionar_carpeta_salida()
    if not carpeta_salida:
        return

    log = VentanaLog(tipo)

    nombre_hoja = verificar_hoja_excel(str(ruta_plantilla), nombre_hoja_buscar, log)
    if not nombre_hoja:
        return

    log.info(f"Plantilla : {ruta_plantilla.name}")
    log.info(f"Destino   : {carpeta_salida}")
    log.info(f"PDFs      : {len(pdfs)}")

    exitosos = 0
    fallidos = []   # lista de nombres de archivos que fallaron

    for i, pdf in enumerate(pdfs, 1):
        pdf_path = Path(pdf)
        log.titulo(f"[{i}/{len(pdfs)}] {pdf_path.name}")

        try:
            datos = extraer_tablas_pdf(str(pdf_path), tipo, log)

            # Verificar que se extrajeron datos reales (no solo declaraciones forzadas).
            # Los campos de declaraciones siempre se rellenan, así que comprobamos
            # que haya al menos un campo de datos personales o empresa.
            CAMPOS_REALES = ["PRIMER APELLIDO", "SEGUNDO APELLIDO", "NOMBRE", "EMAIL",
                             "TELÉFONO", "NIF EMPRESA", "NOMBRE EMPRESA", "DIRECCION"]
            if not any(c in datos for c in CAMPOS_REALES):
                log.error("No se pudieron extraer datos reales (¿PDF escaneado o protegido?)")
                fallidos.append(pdf_path.name)
                continue

            nombre_salida = obtener_nombre_archivo(datos)
            ruta_salida = carpeta_salida / nombre_salida
            contador = 1
            while ruta_salida.exists():
                ruta_salida = carpeta_salida / f"{ruta_salida.stem}_{contador}.xlsm"
                contador += 1

            celdas = rellenar_excel(datos, mapeo, str(ruta_plantilla),
                                    str(ruta_salida), nombre_hoja, log)

            if celdas > 0:
                log.ok(f"Generado: {ruta_salida.name}  ({celdas} celdas)")
                exitosos += 1
            else:
                log.error("No se escribió ninguna celda")
                fallidos.append(pdf_path.name)

        except Exception as ex:
            log.error(f"Error inesperado: {ex}")
            fallidos.append(pdf_path.name)
            import traceback
            traceback.print_exc()

    log.completado(fallidos)

    resumen = f"✅ Convertidos: {exitosos}/{len(pdfs)}"
    if fallidos:
        resumen += f"\n\n⚠ No convertidos ({len(fallidos)}):\n" + "\n".join(f"  • {f}" for f in fallidos)
    resumen += f"\n\nGuardados en:\n{carpeta_salida}"
    messagebox.showinfo("Resultado", resumen)


if __name__ == "__main__":
    main()