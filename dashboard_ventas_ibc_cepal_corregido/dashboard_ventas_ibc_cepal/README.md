# Panel de gestión y ventas IBC · CEPAL

Panel local de análisis basado únicamente en la hoja de detalle de `DETALLE FINAL.xlsx`. Presenta indicadores, gráficos y comparativos por empresa, área, asesor y mes; filtros; detalle de gestiones; y exportación de un informe Excel con resumen, comparativos y registros filtrados.

## Iniciar en Windows

Desde esta carpeta, instala las dependencias e inicia el servidor local:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Abre `http://127.0.0.1:8000`. Si `DETALLE FINAL.xlsx` está en la carpeta Escritorio que contiene este proyecto, el panel lo carga al iniciar. También puedes usar **Cargar Excel** para analizar otro archivo con una única hoja de detalle compatible.

Para configurar una ubicación distinta en PowerShell:

```powershell
$env:SOURCE_XLSX = "C:\ruta\DETALLE FINAL.xlsx"
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

## Indicadores y criterios

- Los resultados guardados, caídos y otras gestiones conservan la clasificación de la columna `CLASIFICACIÓN`; no se vuelven a inferir.
- Los intentos de anulación se cuentan como ventas salvadas + ventas caídas; las tarjetas y el resumen exportado usan esta misma regla.
- Todos los porcentajes del panel y del informe usan como denominador el total de casos registrados en el filtro o grupo correspondiente.
- El plazo de gestión se obtiene de `MOMENTO DÍAS`. Se comparan las ventas salvadas dentro de tres días, fuera de tres días y sin dato.
- El momento de las gestiones respecto al primer pago se obtiene de `MOMENTO PRIMER PAGO`. La existencia de una fecha de primer pago se informa por separado; no se sustituyen fechas o valores faltantes.
- El valor `CONTRATOS` se consulta dentro de **Ver campos** en el detalle de cada gestión.
- Cuando un dato está vacío en el Excel y no existe una coincidencia inequívoca en otra fila, el panel lo marca como `No registrado en el Excel`; no inventa información.
- Los filtros de fecha usan `FECHA GESTIÓN` e incluyen las fechas seleccionadas.
- La exportación aplica los filtros actuales e incluye hojas de resumen, comparación por empresa y área, comparación por asesor, resumen mensual y detalle.

## Privacidad

El panel no guarda análisis en SQLite: el archivo cargado se procesa en memoria y solo queda disponible mientras el proceso del servidor siga activo. La ruta de origen predeterminada apunta al archivo Excel fuera del proyecto; no se copia al repositorio.

El servidor se configura para `127.0.0.1` y no tiene autenticación. No lo expongas a la red local ni a Internet con información de clientes. El archivo Excel exportado contiene los datos de las filas seleccionadas; guárdalo y compártelo de forma segura.

## Pruebas

```powershell
python -m unittest discover -s tests -v
```
