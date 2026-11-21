# Conversor XML del ATS a Excel

Aplicación web estática para GitHub Pages. Permite seleccionar un archivo XML del ATS o pegar su contenido y descargar un archivo `.xlsx`.

## Archivos

- `index.html`: estructura de la página.
- `styles.css`: diseño adaptable a computadora y móvil.
- `script.js`: lectura del XML, extracción de secciones y generación del Excel.
- `README.md`: estas instrucciones.

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub.
2. Sube los cuatro archivos a la raíz del repositorio.
3. Entra en **Settings → Pages**.
4. En **Build and deployment**, selecciona **Deploy from a branch**.
5. Selecciona la rama `main` y la carpeta `/ (root)`.
6. Pulsa **Save** y espera a que GitHub Pages publique el sitio.

La dirección tendrá una forma similar a `https://TU-USUARIO.github.io/NOMBRE-DEL-REPOSITORIO/`.

## Cómo utilizarlo

1. Abre la página publicada.
2. Selecciona el XML del ATS o pega el contenido completo en el cuadro.
3. Pulsa **Convertir y descargar Excel**.
4. El archivo `.xlsx` se descargará en el dispositivo.

## Estructura del Excel

El libro contiene cuatro hojas:

1. Datos generales
2. Compras
3. Ventas
4. Ventas establecimiento

Las etiquetas anidadas del XML se convierten en nombres de columna unidos por guion bajo, por ejemplo `pagoExterior_pagoLocExt`. Se aplican encabezados azul oscuro con texto blanco, filtros, primera fila inmovilizada y anchos de columnas basados en el libro de referencia `ReporteATS_2025-10_CDEF.xlsx`.

Se conservan como texto las columnas:
- Compras: A, C, G, H, I, K
- Ventas: A, B
- Ventas establecimiento: A

## Privacidad y requisitos

- El XML se procesa dentro del navegador del usuario; no se carga a un servidor.
- Requiere conexión a Internet para cargar ExcelJS desde jsDelivr (CDN).
- No requiere Python, Excel ni complementos de Excel.
- La aplicación usa ExcelJS 4.4.0 desde CDN para generar el archivo `.xlsx` con estilos.
