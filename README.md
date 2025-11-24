# Primeros pasos

Este repositorio incluye un módulo VBA que exporta todas las gráficas de un libro de Excel a una carpeta creada automáticamente junto al archivo. Las imágenes se guardan en formato PNG usando el tamaño real del gráfico para evitar pérdida de nitidez o cambios en proporciones.

## Cómo usar la macro
1. Abre el archivo Excel que contiene las gráficas y presiona `ALT + F11` para abrir el Editor de VBA.
2. En el menú **Archivo**, elige **Importar archivo** y selecciona `ExportarGraficas.bas` desde este repositorio.
3. Guarda el libro como `.xlsm` si aún no tiene macros.
4. Ejecuta `ExportarGraficasComoImagenes` desde el Editor de VBA o asígnala a un botón.

La macro creará una carpeta con la fecha y hora actuales (por ejemplo, `2024-06-20_1530_Graficas`) en la misma ubicación del libro y guardará ahí cada gráfica en PNG sin alterar tamaños ni texto.
