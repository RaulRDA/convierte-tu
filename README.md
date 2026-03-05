# PDF a Excel — Grupo ATU

Herramienta de escritorio desarrollada para Grupo ATU que automatiza el proceso de trasladar los datos de los formularios PDF de inscripción del programa Generación Digital a las plantillas Excel correspondientes.

Hasta ahora ese trabajo se hacía a mano campo por campo. Esta aplicación lo hace en segundos.

---

## Qué hace

Lee un formulario PDF rellenado por un alumno, extrae todos sus campos y los vuelca en la plantilla Excel correcta, generando un archivo `.xlsm` listo para entregar. Soporta los tres tipos de formulario del programa:

- Agentes del Cambio — versión 1
- Agentes del Cambio — versión 2
- Personas de Equipos Directivos

---

## Uso

Ejecuta `PDF_a_Excel.exe`, selecciona el tipo de plantilla, elige los PDFs que quieras procesar, indica la carpeta de destino y espera. Al terminar aparece un resumen con los archivos convertidos y, si los hay, los que no se pudieron leer.

Los archivos generados siguen el formato `Formulario Apellido_Nombre.xlsm`.

---

## Detalles técnicos

### Dependencias

```
pdfplumber   — extracción de tablas del PDF
openpyxl     — lectura y escritura de ficheros Excel (.xlsm con VBA)
tkinter      — interfaz gráfica (incluido en la instalación estándar de Python)
```

### Estructura del código

Todo el proyecto vive en un único fichero `pdf_a_excel_agentes.py`. La razón es sencilla: facilita el empaquetado con PyInstaller y reduce la fricción para quien lo mantenga.

**Mapeos de celdas**

Cada tipo de formulario tiene su propio diccionario `MAPEO_*` que relaciona el nombre interno del campo con la celda exacta de la plantilla Excel (`"NOMBRE": "B8"`, etc.). Si la plantilla cambia de estructura, solo hay que tocar el diccionario correspondiente.

**Extracción del PDF**

La función `extraer_tablas_pdf()` itera sobre todas las tablas de todas las páginas usando `pdfplumber`. Cada fila se evalúa contra una cadena de `elif` que identifica el campo por palabras clave en la etiqueta.

Hay tres particularidades importantes documentadas en el código:

1. En el formulario de Directivos, los campos `TIPO DE DOCUMENTO` y `Nº DE DOCUMENTO` aparecen invertidos en el PDF respecto al Excel, así que se cruzan al asignarlos.

2. `FORMACIÓN GESTIÓN PROYECTOS` no tiene etiqueta propia en algunos PDFs — aparece como una fila sin label justo después de `FORMACIÓN DIGITALIZACIÓN`. Se resuelve rastreando `ultima_clave`.

3. Las declaraciones finales (Acepto/No procede) están en columnas distintas según el tipo de formulario, por lo que el valor se busca en toda la fila si la columna 1 viene vacía.

**Declaraciones forzadas**

En Agentes (2) y Directivos todos los términos se fijan a `"Acepto"` independientemente de lo que diga el PDF, ya que ese es el único valor válido en esos formularios. En Agentes (1) solo el campo `DECLARO PYME` se lee del PDF porque puede contener `"No procede"`.

**Detección de PDFs no legibles**

Un PDF escaneado como imagen no tiene texto extraíble, pero las declaraciones forzadas harían que `datos` nunca estuviese vacío. Por eso la validación comprueba que exista al menos un campo real (`NOMBRE`, `EMAIL`, `NIF EMPRESA`, etc.) antes de considerar la extracción exitosa.

**Ruta base**

```python
def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent
```

Cuando PyInstaller empaqueta la app con `--onefile`, extrae los recursos a una carpeta temporal en `sys._MEIPASS`. Esta función devuelve siempre la ruta correcta tanto en desarrollo como en el ejecutable.

### Compilar el ejecutable

```bash
pip install pyinstaller pdfplumber openpyxl

py -m PyInstaller --onefile --windowed --name "PDF_a_Excel" \
  --icon="icono.ico" \
  --add-data "_plantillas;_plantillas" \
  --add-data "icono.ico;." \
  --collect-all pdfplumber \
  --collect-all pdfminer \
  PDF_a_Excel.py
```

El ejecutable resultante queda en `dist/PDF_a_Excel.exe`. Las plantillas y el icono van empaquetados dentro, no hace falta distribuir nada más.

### Añadir un nuevo tipo de formulario

1. Crear el diccionario `MAPEO_NUEVO` con los campos y celdas.
2. Añadir la entrada en `PLANTILLAS` con la ruta de la plantilla, el nombre de la hoja y el límite de PDFs.
3. Actualizar el selector de mapeo en `main()`.
4. Si el nuevo formulario tiene etiquetas distintas para campos ya existentes, añadir las condiciones necesarias en `extraer_tablas_pdf()`.

---

## Limitaciones conocidas

Los PDFs escaneados como imagen no son legibles. `pdfplumber` solo extrae texto de PDFs digitales. Si un alumno entrega un PDF generado escaneando el formulario en papel, la app lo detectará y lo reportará como no convertido.

---

## Autores

Desarrollado por RaulRDA.com, Pablo Álvarez y Pelayo Fernández para Grupo ATU.