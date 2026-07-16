# Conciliación Banco/Clientes en Dólares

Web para cruzar el extracto bancario (Banco Macro, USD) contra el formulario
de transferencias de clientes, con los mismos niveles de match que se usan
manualmente: **ALTO / BAJO / SIN MATCH / REPETIDO**.

Corre 100% en el navegador. Ningún archivo se sube a un servidor: todo el
procesamiento y la generación de los Excel de salida ocurre localmente,
usando la librería [ExcelJS](https://github.com/exceljs/exceljs) vía CDN.

## Cómo publicarla en GitHub Pages

1. Creá un repositorio nuevo en GitHub (puede ser privado).
2. Subí estos tres archivos a la raíz del repo: `index.html`, `app.js`, `README.md`.
3. Andá a **Settings → Pages**.
4. En "Source", elegí **Deploy from a branch**, rama `main`, carpeta `/ (root)`.
5. Guardá. GitHub te va a dar una URL tipo `https://tu-usuario.github.io/tu-repo/`.
6. Entrá a esa URL y ya podés usarla desde cualquier dispositivo.

## Cómo se usa

1. **Subir archivos**: arrastrá o hacé click para cargar el archivo de BANCO
   (extracto Banco Macro) y el de CLIENTES (formulario de transferencias).
2. Apretá **Analizar**.
3. **Cotización BNA compra**: la web detecta todas las fechas presentes en
   el extracto. Para la fecha de hoy intenta autocompletar la cotización
   con una API pública (dolarapi.com); para el resto, cargala manualmente.
   Podés usar "Aplicar a fechas vacías" para poner el mismo valor en todas
   las que falten.
4. Revisá la tabla de resultados y el resumen por nivel.
5. Apretá **Descargar CLIENTES y BANCO (.xlsx)** para bajar los dos archivos
   ya completados, con los mismos formatos, colores y columnas que se usan
   en el proceso manual.

## Qué detecta automáticamente

- **ALTO**: coincide CUIT/CUIL + monto + fecha exacta contra una fila
  disponible del banco.
- **BAJO**: coincidencias parciales — CUIT incompleto (8-10 dígitos) que
  aparece en el campo CONCEPTO del banco, CUIT con 1-2 dígitos distintos
  pero mismo monto, o CUIT exacto con monto/fecha que no coinciden.
- **SIN MATCH**: no se encontró ninguna fila de banco disponible.
- **REPETIDO**: mismo número de cliente + mismo CUIT + mismo monto +
  misma referencia bancaria que otra fila ya cargada.

## Qué NO hace automáticamente

Los casos ambiguos (por ejemplo, dos números de cliente distintos que
casualmente comparten CUIT y monto, o montos con errores de tipeo evidentes)
quedan clasificados como **BAJO** por defecto, sin preguntar. Si necesitás
ajustar alguno de estos casos a mano, podés editar el archivo descargado
directamente en Excel antes de enviarlo.

## Cotización del dólar: automática vs. manual

Una página estática como esta corre en el navegador del usuario, sin
servidor propio. No puede navegar la web libremente como lo haría un asistente
con herramientas de búsqueda, y el sitio oficial del BNA bloquea el acceso
automatizado desde el navegador (CORS). Por eso la cotización de **hoy** se
intenta autocompletar con una API pública (que no siempre coincide al
centavo con el cierre oficial del BNA), y las fechas anteriores requieren
carga manual — igual que se viene haciendo hasta ahora cuando hay
discrepancias entre fuentes.
