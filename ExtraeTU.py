"""
╔════════════════════════════════════════════════════════════╗
║   Extractor de PDFs para Fichas de Inscripción             ║
║   Grupo ATU © 2026                                         ║
║   Hecho por RaulRDA.com, Pablo Álvarez y Pelayo Fernández  ║
╚════════════════════════════════════════════════════════════╝
"""

import shutil
import unicodedata
import difflib
import threading
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ──────────────────────────────────────────────
VERSION = "1.0.0"
CARPETA_OBJETIVO_REF = "10 FICHA INSCRIPCION"   # normalizada (sin tildes, sin guiones bajos)
UMBRAL_SIMILITUD = 0.72
# ──────────────────────────────────────────────

# ── Paleta de colores ─────────────────────────
BG_DARK    = "#0f1117"
BG_PANEL   = "#1a1d27"
BG_INPUT   = "#22263a"
ACCENT     = "#4f8ef7"
ACCENT2    = "#2563eb"
SUCCESS    = "#22c55e"
WARNING    = "#f59e0b"
ERROR      = "#ef4444"
TEXT_MAIN  = "#e8eaf0"
TEXT_DIM   = "#6b7280"
TEXT_LOG   = "#a0aec0"
BORDER     = "#2d3148"
# ─────────────────────────────────────────────


def normalizar(texto: str) -> str:
    """Quita tildes, pasa a mayúsculas, reemplaza guiones/guiones bajos por espacios y colapsa espacios."""
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = texto.upper()
    texto = texto.replace("_", " ").replace("-", " ")
    texto = " ".join(texto.split())
    return texto


def es_carpeta_ficha(nombre: str) -> bool:
    """Devuelve True si el nombre de la carpeta se parece a '10_ FICHA INSCRIPCIÓN'."""
    norm = normalizar(nombre)
    ratio = difflib.SequenceMatcher(None, norm, CARPETA_OBJETIVO_REF).ratio()
    # También aceptamos si contiene las palabras clave principales
    contiene_clave = ("FICHA" in norm and "INSCRIPCION" in norm) or \
                     ("FICHA" in norm and "10" in norm)
    return ratio >= UMBRAL_SIMILITUD or contiene_clave


def encontrar_carpeta_ficha(ruta_alumno: Path):
    """Busca dentro de la carpeta del alumno la subcarpeta de ficha. Devuelve Path o None."""
    for item in ruta_alumno.iterdir():
        if item.is_dir() and es_carpeta_ficha(item.name):
            return item
    return None


def encontrar_pdf(ruta_carpeta: Path):
    """Devuelve el primer PDF encontrado en la carpeta, o None."""
    for item in ruta_carpeta.iterdir():
        if item.is_file() and item.suffix.lower() == ".pdf":
            return item
    return None


# ══════════════════════════════════════════════
#  INTERFAZ GRÁFICA
# ══════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TuExtractorPDF")
        self.resizable(False, False)
        self.configure(bg=BG_DARK)
        self._centrar_ventana(780, 640)

        self._aplicar_estilos()
        self._construir_ui()

    # ── Utilidades ────────────────────────────

    def _centrar_ventana(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _aplicar_estilos(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame", background=BG_DARK)
        style.configure("Panel.TFrame", background=BG_PANEL)

        style.configure("TLabel",
                        background=BG_DARK,
                        foreground=TEXT_MAIN,
                        font=("Segoe UI", 10))

        style.configure("Dim.TLabel",
                        background=BG_PANEL,
                        foreground=TEXT_DIM,
                        font=("Segoe UI", 9))

        style.configure("Title.TLabel",
                        background=BG_DARK,
                        foreground=TEXT_MAIN,
                        font=("Segoe UI", 22, "bold"))

        style.configure("Sub.TLabel",
                        background=BG_DARK,
                        foreground=TEXT_DIM,
                        font=("Segoe UI", 10))

        style.configure("Section.TLabel",
                        background=BG_PANEL,
                        foreground=ACCENT,
                        font=("Segoe UI", 9, "bold"))

        style.configure("Input.TLabel",
                        background=BG_PANEL,
                        foreground=TEXT_MAIN,
                        font=("Segoe UI", 10))

        # Botón principal
        style.configure("Primary.TButton",
                        background=ACCENT,
                        foreground="#ffffff",
                        font=("Segoe UI", 11, "bold"),
                        borderwidth=0,
                        focusthickness=0,
                        padding=(20, 10))
        style.map("Primary.TButton",
                  background=[("active", ACCENT2), ("disabled", "#2d3148")],
                  foreground=[("disabled", TEXT_DIM)])

        # Botón secundario (explorar)
        style.configure("Browse.TButton",
                        background=BG_INPUT,
                        foreground=ACCENT,
                        font=("Segoe UI", 9, "bold"),
                        borderwidth=0,
                        focusthickness=0,
                        padding=(10, 6))
        style.map("Browse.TButton",
                  background=[("active", "#2d3148")])

        # Barra de progreso
        style.configure("Blue.Horizontal.TProgressbar",
                        troughcolor=BG_INPUT,
                        background=ACCENT,
                        borderwidth=0,
                        thickness=8)

    # ── UI Principal ──────────────────────────

    def _construir_ui(self):
        # ── CABECERA ──
        header = tk.Frame(self, bg=BG_DARK)
        header.pack(fill="x", padx=30, pady=(28, 0))

        tk.Label(header, text="TuExtractorPDF",
                 font=("Segoe UI", 24, "bold"),
                 bg=BG_DARK, fg=TEXT_MAIN).pack(anchor="w")
        tk.Label(header, text="Extrae automáticamente las fichas de inscripción de cada alumno",
                 font=("Segoe UI", 10),
                 bg=BG_DARK, fg=TEXT_DIM).pack(anchor="w", pady=(2, 0))

        # separador decorativo
        sep = tk.Frame(self, bg=ACCENT, height=2)
        sep.pack(fill="x", padx=30, pady=(14, 20))

        # ── PANEL CENTRAL ──
        panel = tk.Frame(self, bg=BG_PANEL, bd=0, highlightthickness=1,
                         highlightbackground=BORDER)
        panel.pack(fill="x", padx=30, pady=0)

        self._seccion_ruta(panel, "📁  CARPETA DE ALUMNOS",
                           "Selecciona la carpeta raíz que contiene las subcarpetas de cada alumno",
                           "origen")
        tk.Frame(panel, bg=BORDER, height=1).pack(fill="x", padx=20)
        self._seccion_ruta(panel, "💾  CARPETA DE DESTINO",
                           "Aquí se copiarán todos los PDFs renombrados con el nombre del alumno",
                           "destino")

        # ── BOTÓN EJECUTAR ──
        btn_frame = tk.Frame(self, bg=BG_DARK)
        btn_frame.pack(pady=20)

        self.btn_ejecutar = ttk.Button(btn_frame, text="▶  Extraer PDFs",
                                       style="Primary.TButton",
                                       command=self._iniciar_extraccion)
        self.btn_ejecutar.pack()

        # ── PROGRESO ──
        prog_frame = tk.Frame(self, bg=BG_DARK)
        prog_frame.pack(fill="x", padx=30)

        self.lbl_estado = tk.Label(prog_frame, text="",
                                   font=("Segoe UI", 9),
                                   bg=BG_DARK, fg=TEXT_DIM, anchor="w")
        self.lbl_estado.pack(fill="x")

        self.progreso = ttk.Progressbar(prog_frame, style="Blue.Horizontal.TProgressbar",
                                        mode="determinate", length=720)
        self.progreso.pack(fill="x", pady=(4, 0))

        # ── LOG ──
        log_outer = tk.Frame(self, bg=BG_PANEL, bd=0,
                             highlightthickness=1, highlightbackground=BORDER)
        log_outer.pack(fill="both", expand=True, padx=30, pady=16)

        log_header = tk.Frame(log_outer, bg=BG_PANEL)
        log_header.pack(fill="x", padx=16, pady=(10, 4))
        tk.Label(log_header, text="REGISTRO DE ACTIVIDAD",
                 font=("Segoe UI", 8, "bold"),
                 bg=BG_PANEL, fg=ACCENT).pack(side="left")

        self.txt_log = tk.Text(log_outer, bg=BG_PANEL, fg=TEXT_LOG,
                               font=("Consolas", 9),
                               bd=0, highlightthickness=0,
                               wrap="word", state="disabled",
                               insertbackground=TEXT_MAIN)
        scroll = ttk.Scrollbar(log_outer, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y", padx=(0, 4), pady=(0, 10))
        self.txt_log.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        # Tags de color para el log
        self.txt_log.tag_configure("ok",      foreground=SUCCESS)
        self.txt_log.tag_configure("warn",    foreground=WARNING)
        self.txt_log.tag_configure("err",     foreground=ERROR)
        self.txt_log.tag_configure("info",    foreground=ACCENT)
        self.txt_log.tag_configure("dim",     foreground=TEXT_DIM)

        # ── PIE ──
        tk.Label(self, text=f"v{VERSION}  ·  Grupo ATU © {datetime.date.today().year}",
                 font=("Segoe UI", 8),
                 bg=BG_DARK, fg=TEXT_DIM).pack(pady=(0, 10))

    def _seccion_ruta(self, parent, titulo, descripcion, key):
        frame = tk.Frame(parent, bg=BG_PANEL)
        frame.pack(fill="x", padx=20, pady=14)

        tk.Label(frame, text=titulo,
                 font=("Segoe UI", 9, "bold"),
                 bg=BG_PANEL, fg=ACCENT).pack(anchor="w")
        tk.Label(frame, text=descripcion,
                 font=("Segoe UI", 9),
                 bg=BG_PANEL, fg=TEXT_DIM).pack(anchor="w", pady=(1, 6))

        row = tk.Frame(frame, bg=BG_PANEL)
        row.pack(fill="x")

        var = tk.StringVar()
        setattr(self, f"var_{key}", var)

        entrada = tk.Entry(row, textvariable=var,
                           bg=BG_INPUT, fg=TEXT_MAIN,
                           font=("Segoe UI", 10),
                           bd=0, highlightthickness=1,
                           highlightbackground=BORDER,
                           highlightcolor=ACCENT,
                           insertbackground=TEXT_MAIN,
                           relief="flat")
        entrada.pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 8))

        btn = ttk.Button(row, text="Explorar…", style="Browse.TButton",
                         command=lambda k=key: self._explorar(k))
        btn.pack(side="right")

    # ── Lógica ───────────────────────────────

    def _explorar(self, key):
        ruta = filedialog.askdirectory(title="Seleccionar carpeta")
        if ruta:
            getattr(self, f"var_{key}").set(ruta)

    def _log(self, texto, tag=""):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", texto + "\n", tag)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _limpiar_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _set_estado(self, texto):
        self.lbl_estado.config(text=texto)

    def _iniciar_extraccion(self):
        origen = self.var_origen.get().strip()
        destino = self.var_destino.get().strip()

        if not origen or not destino:
            messagebox.showwarning("Faltan rutas",
                                   "Por favor, selecciona tanto la carpeta de alumnos como la de destino.")
            return
        if not Path(origen).is_dir():
            messagebox.showerror("Carpeta no válida",
                                 "La carpeta de alumnos seleccionada no existe.")
            return
        if not Path(destino).is_dir():
            messagebox.showerror("Carpeta no válida",
                                 "La carpeta de destino seleccionada no existe.")
            return

        self.btn_ejecutar.config(state="disabled")
        self._limpiar_log()
        self.progreso["value"] = 0

        hilo = threading.Thread(target=self._ejecutar,
                                args=(Path(origen), Path(destino)),
                                daemon=True)
        hilo.start()

    def _ejecutar(self, origen: Path, destino: Path):
        alumnos = [p for p in sorted(origen.iterdir()) if p.is_dir()]
        total = len(alumnos)

        if total == 0:
            self.after(0, self._log,
                       "⚠  No se encontraron subcarpetas de alumnos en la carpeta seleccionada.", "warn")
            self.after(0, self.btn_ejecutar.config, {"state": "normal"})
            return

        self.after(0, self._log,
                   f"Iniciando extracción · {total} alumno(s) encontrado(s)\n", "info")

        errores = []       # (nombre_alumno, motivo)
        copiados = []      # nombres de alumnos OK

        for i, ruta_alumno in enumerate(alumnos, 1):
            nombre_alumno = ruta_alumno.name
            self.after(0, self._set_estado, f"Procesando: {nombre_alumno}  ({i}/{total})")
            self.after(0, lambda v=int(i / total * 100): self.progreso.config(value=v))

            # Buscar subcarpeta de ficha
            carpeta_ficha = encontrar_carpeta_ficha(ruta_alumno)
            if carpeta_ficha is None:
                msg = f"✗  {nombre_alumno}  →  No se encontró la carpeta '10_ FICHA INSCRIPCIÓN'"
                self.after(0, self._log, msg, "err")
                errores.append((nombre_alumno, "Carpeta '10_ FICHA INSCRIPCIÓN' no encontrada"))
                continue

            # Buscar PDF dentro
            pdf = encontrar_pdf(carpeta_ficha)
            if pdf is None:
                msg = f"✗  {nombre_alumno}  →  Carpeta encontrada ({carpeta_ficha.name}) pero sin PDF"
                self.after(0, self._log, msg, "err")
                errores.append((nombre_alumno, f"No hay PDF en '{carpeta_ficha.name}'"))
                continue

            # Copiar con el nombre del alumno
            destino_pdf = destino / f"{nombre_alumno}.pdf"
            try:
                shutil.copy2(pdf, destino_pdf)
                msg = f"✓  {nombre_alumno}  →  {destino_pdf.name}"
                self.after(0, self._log, msg, "ok")
                copiados.append(nombre_alumno)
            except Exception as exc:
                msg = f"✗  {nombre_alumno}  →  Error al copiar: {exc}"
                self.after(0, self._log, msg, "err")
                errores.append((nombre_alumno, f"Error al copiar: {exc}"))

        # ── Resumen final ──
        self.after(0, self._log, "", "")
        self.after(0, self._log,
                   f"{'─'*60}", "dim")
        self.after(0, self._log,
                   f"✔  Extraídos correctamente: {len(copiados)}", "ok")
        if errores:
            self.after(0, self._log,
                       f"✘  Con problemas: {len(errores)}", "err")
        self.after(0, self._log,
                   f"{'─'*60}", "dim")

        # Generar log de errores si los hay
        if errores:
            self.after(0, self._generar_log_errores, destino, errores)

        self.after(0, self._set_estado, "Proceso completado.")
        self.after(0, lambda: self.progreso.config(value=100))
        self.after(0, self.btn_ejecutar.config, {"state": "normal"})
        self.after(0, self._mostrar_resumen, len(copiados), errores)

    def _generar_log_errores(self, destino: Path, errores: list):
        fecha = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_log = destino / f"log_errores_{fecha}.txt"
        lineas = [
            "TuExtractorPDF — Log de errores",
            f"Fecha: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
            f"Total con errores: {len(errores)}",
            "=" * 50,
            ""
        ]
        for alumno, motivo in errores:
            lineas.append(f"Alumno : {alumno}")
            lineas.append(f"Motivo : {motivo}")
            lineas.append("-" * 40)

        ruta_log.write_text("\n".join(lineas), encoding="utf-8")
        self._log(f"\n📄 Log de errores guardado en:\n   {ruta_log}", "warn")

    def _mostrar_resumen(self, n_ok: int, errores: list):
        if not errores:
            messagebox.showinfo(
                "Extracción completada",
                f"✅  Se han extraído {n_ok} PDF(s) correctamente.\n\n"
                "Todos los archivos están en la carpeta de destino."
            )
        else:
            messagebox.showwarning(
                "Extracción completada con advertencias",
                f"✅  Extraídos: {n_ok} PDF(s)\n"
                f"⚠   Con problemas: {len(errores)}\n\n"
                "Consulta el log de errores en la carpeta de destino para más detalles."
            )


# ══════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()