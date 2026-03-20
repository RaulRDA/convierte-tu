# ConvierteTU — Grupo ATU

Herramienta de escritorio desarrollada para Grupo ATU que automatiza el proceso de trasladar los datos de los formularios PDF de inscripción del programa Generación Digital a las plantillas Excel correspondientes.

Hasta ahora ese trabajo se hacía a mano campo por campo. Esta aplicación lo hace en segundos.

---

## Qué hace

Lee un formulario PDF rellenado por un alumno, extrae todos sus campos y los vuelca en la plantilla Excel correcta, generando un archivo `.xlsm` listo para entregar. Soporta los tres tipos de formulario del programa:

- Agentes del Cambio — versión 1
- Agentes del Cambio — versión 2
- Personas de Equipos Directivos

Al arrancar, comprueba en segundo plano si hay una versión más reciente disponible y avisa al usuario con opción de ir a descargarla. Si no hay conexión, el programa arranca igualmente sin ningún problema.

---

## Compatibilidad

La aplicación es compatible con **Windows** y **macOS**. Todas las dependencias (pdfplumber, pypdf, openpyxl, requests, tkinter) tienen versiones para ambas plataformas, y el código no usa ninguna API exclusiva de Windows. En macOS, el icono de la ventana no se aplica (el formato `.ico` es exclusivo de Windows), pero la aplicación funciona con normalidad.

---

## Uso

**Windows:** ejecuta `ConvierteTU.exe`.

**macOS:** ejecuta `ConvierteTU.app` o, si usas el código fuente directamente, `python3 main.py`.

En ambos casos: selecciona el tipo de plantilla, elige los PDFs que quieras procesar, indica la carpeta de destino y espera. Al terminar aparece un resumen con los archivos convertidos y, si los hay, los que no se pudieron leer.

Los archivos generados siguen el formato `Formulario Apellido_Nombre.xlsm`.

---

## Detalles técnicos

### Dependencias

```
pdfplumber   — extracción de tablas del PDF
pypdf        — lectura alternativa de texto del PDF
openpyxl     — lectura y escritura de ficheros Excel (.xlsm con VBA)
requests     — comprobación de actualizaciones
tkinter      — interfaz gráfica (incluido en la instalación estándar de Python)
```

### Estructura del código

Todo el proyecto vive en un único fichero `main.py`. La razón es sencilla: facilita el empaquetado con PyInstaller y reduce la fricción para quien lo mantenga.

**Mapeos de celdas**

Cada tipo de formulario tiene su propio diccionario `MAPEO_*` que relaciona el nombre interno del campo con la celda exacta de la plantilla Excel (`"NOMBRE": "B8"`, etc.). Si la plantilla cambia de estructura, solo hay que tocar el diccionario correspondiente.

**Extracción del PDF**

El motor de extracción combina tres estrategias en cascada, ejecutándolas en orden y mezclando los resultados para maximizar los campos recuperados:

1. **Tablas** — `_extraer_por_tablas()` itera sobre todas las tablas de todas las páginas usando `pdfplumber`. Cada fila se evalúa contra una cadena de `elif` que identifica el campo por palabras clave en la etiqueta, usando siempre la cadena normalizada (sin tildes, en mayúsculas) para evitar fallos con variaciones tipográficas del PDF.

2. **Coordenadas X** — `_extraer_por_coordenadas()` analiza la posición horizontal de cada palabra en la página. Todo lo que cae a la izquierda de un umbral fijo se trata como etiqueta; lo que cae a la derecha, como valor. Útil cuando el PDF no tiene tablas reales sino texto posicionado visualmente.

3. **Líneas de texto** — `_extraer_por_lineas()` recorre las líneas en busca de patrones de etiquetas conocidas. Actúa como último recurso y rellena solo los campos que las dos estrategias anteriores no hayan capturado.

Hay dos particularidades importantes documentadas en el código:

1. `FORMACIÓN GESTIÓN PROYECTOS` no tiene etiqueta propia en algunos PDFs — aparece como una fila sin label justo después de `FORMACIÓN DIGITALIZACIÓN`. Se resuelve rastreando `ultima_clave` durante la iteración de tablas.

2. Las declaraciones finales se fijan siempre a `"Acepto"` en Agentes (2) y Directivos, ya que ese es el único valor válido. En Agentes (1) el campo `DECLARO PYME` se lee del PDF porque puede contener `"No procede"`.

**Post-procesado**

Tras la extracción, `postprocesar_campos()` normaliza los valores antes de escribirlos en el Excel: limpia texto residual, resuelve listas de opciones con `/`, extrae el valor SI/NO de campos booleanos, parsea fechas en distintos formatos y vacía los datos de empresa si el participante declara estar desempleado.

**Detección de PDFs no legibles**

Un PDF escaneado como imagen no tiene texto extraíble, pero las declaraciones forzadas harían que `datos` nunca estuviese vacío. Por eso la validación comprueba que exista al menos un campo real (`NOMBRE`, `EMAIL`, `NIF EMPRESA`, etc.) antes de considerar la extracción exitosa.

**Log de errores**

Si algún PDF no se puede convertir, se genera automáticamente un fichero `log_errores_*.txt` en la carpeta de destino con el nombre del archivo, el tipo de error y el detalle completo.

**Actualizaciones automáticas**

`lanzar_comprobacion_actualizacion()` arranca un hilo `daemon` que consulta la API de GitHub al inicio. Si encuentra una versión con tag superior a `VERSION`, programa la notificación en el hilo principal con `root.after()`. Cualquier error de red se ignora silenciosamente para no interrumpir al usuario.

**Ruta base**

```python
def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent
```

Cuando PyInstaller empaqueta la app con `--onefile`, extrae los recursos a una carpeta temporal en `sys._MEIPASS`. Esta función devuelve siempre la ruta correcta tanto en desarrollo como en el ejecutable.

**Versión y URL de actualizaciones**

Las dos únicas constantes que hay que tocar en cada nueva versión están al principio de `main.py`:

```python
VERSION    = "1.1.0"
UPDATE_URL = "<url>/releases/latest"
```

### Compilar el ejecutable

**Windows**

```bash
pip install pyinstaller pdfplumber pypdf openpyxl requests

py -m PyInstaller --onefile --windowed --name "ConvierteTU" \
  --icon="icono.ico" \
  --add-data "_plantillas;_plantillas" \
  --add-data "icono.ico;." \
  --collect-all pdfplumber \
  --collect-all pdfminer \
  main.py
```

El ejecutable resultante queda en `dist/ConvierteTU.exe`.

**macOS**

```bash
pip3 install pyinstaller pdfplumber pypdf openpyxl requests

python3 -m PyInstaller --onefile --windowed --name "ConvierteTU" \
  --add-data "_plantillas:_plantillas" \
  --collect-all pdfplumber \
  --collect-all pdfminer \
  main.py
```

El bundle resultante queda en `dist/ConvierteTU.app`. En macOS se omite el parámetro `--icon` porque el formato `.ico` no es compatible; si quieres icono personalizado, convierte `icono.ico` a `icono.icns` y usa `--icon="icono.icns"`.

En ambas plataformas las plantillas van empaquetadas dentro del ejecutable y no hace falta distribuir nada más.

### Añadir un nuevo tipo de formulario

1. Crear el diccionario `MAPEO_NUEVO` con los campos y celdas y añadirlo a `MAPEO_POR_TIPO`.
2. Añadir la entrada en `PLANTILLAS` con la ruta de la plantilla, el nombre de la hoja y el límite de PDFs.
3. Ampliar `_mapear_campo_tabla()` si el nuevo formulario tiene etiquetas distintas para campos ya existentes.
4. Si hay campos booleanos nuevos, añadirlos a `CAMPOS_SI_NO`. Si hay campos de empresa nuevos, añadirlos a `CAMPOS_EMPRESA`.

---

## Limitaciones conocidas

Los PDFs escaneados como imagen no son legibles. `pdfplumber` y `pypdf` solo extraen texto de PDFs digitales. Si un alumno entrega un PDF generado escaneando el formulario en papel, la app lo detectará y lo reportará como no convertido.

---

## Créditos

Hecho por [RaulRDA](https://raulrda.com) · [Pablo Álvarez](https://github.com/pabloalvf2004) · [Pelayo Fernández](https://github.com/Pelayo89) | Grupo ATU © 2026