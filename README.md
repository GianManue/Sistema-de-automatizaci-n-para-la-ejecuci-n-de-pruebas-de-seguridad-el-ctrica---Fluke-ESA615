  ## Automatización de Reportes - Fluke ESA615

Una aplicación de escritorio desarrollada en Python que automatiza la generación de informes en Excel a partir de los datos crudos extraídos de analizadores de seguridad eléctrica (Fluke ESA615).

_Descripción del Proyecto_ 

Este proyecto elimina la necesidad de transcribir manualmente los resultados de las pruebas de equipos médicos. La aplicación recolecta los datos de las pruebas exportados en formato `.csv`, busca etiquetas clave dentro de una plantilla maestra de Excel y genera un informe final formateado y listo para entregar.

 _Características Principales_

**Inyección de Datos Dinámica:** No utiliza coordenadas rígidas (como `A1` o `B5`). El sistema escanea la plantilla de Excel buscando "códigos" o "etiquetas" específicas y las reemplaza por los valores numéricos correspondientes. Esto hace que el código sea resistente a cambios de diseño en la plantilla.
**Portabilidad Total:** El módulo de procesamiento de datos cuenta con una busqueda absoluta de rutas. El `.exe` generado es completamente portátil; detecta automáticamente su entorno de ejecución para localizar los archivos generados y recursos internos sin depender de rutas estáticas en el disco duro.
**Interfaz Gráfica (GUI):** Interfaz moderna e intuitiva construida con `customtkinter`.
**Empaquetado Standalone:** Compilable en un único archivo `.exe` para Windows, incluyendo la plantilla de Excel como un recurso oculto gracias a PyInstaller.

_Librerías Utilizadas_

Python 3.x
openpyxl: Para la lectura, manipulación y guardado de archivos `.xlsx` sin perder el formato original.
csv: Para el procesamiento de los datos crudos extraídos del equipo Fluke.
CustomTkinter / Tkinter:Para el desarrollo de la interfaz gráfica de usuario.
PyInstaller:Para la compilación y empaquetado del software.
Añadir a estas librerías , las librerías instaladas en python por defecto
Estructura del Proyecto

```text
📁 Automatizacion_ESA615
├── AppFluke.py                # Script principal (Interfaz Gráfica)
├── data_to_excel.py           # Módulo principal (Lógica de inyección Excel y rutas)
├── Pruebas_recibir_info.py    # (Opcional) Módulo de recolección de datos
├── Copia de informe_pruebas.xlsx # Plantilla maestra de Excel
└── ArchivosGenerados/         # Carpeta de salida (Generada automáticamente)
    ├── resultados_esa615.csv
    ├── resultados_bloques_esa615.csv
    └── INFORME_COMPLETO_ESA615.xlsx

Atentamente, el encargado de la elaboración del código para este proyecto, Gian Rojas :D.
