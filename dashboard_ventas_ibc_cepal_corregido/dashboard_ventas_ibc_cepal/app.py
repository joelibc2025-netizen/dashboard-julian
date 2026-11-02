import io
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from engine import (
    CLASSIFICATIONS,
    SOURCE_COLUMNS,
    canonical_header,
    normalize_record,
    summarize,
)


BASE = Path(__file__).resolve().parent
DEFAULT_SOURCE = Path(
    os.getenv("SOURCE_XLSX", str(BASE.parent.parent / "DETALLE FINAL.xlsx"))
)
MAX_UPLOAD_BYTES = 50 * 1024 * 1024
MAX_ROWS = 200_000
REQUIRED_COLUMNS = {
    "FECHA GESTIÓN",
    "EMPRESA",
    "ÁREA",
    "VENDEDOR",
    "CLIENTE",
    "DNI",
    "CLASIFICACIÓN",
    "MOMENTO DÍAS",
    "FECHA PRIMER PAGO",
    "MOMENTO PRIMER PAGO",
}

app = FastAPI(title="Análisis de ventas | IBC · CEPAL", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=BASE / "static"), name="static")
ACTIVE_ROWS: list[dict[str, Any]] = []
ACTIVE_FILENAME = ""
ACTIVE_WARNINGS: list[str] = []


def parse_workbook(raw: bytes) -> tuple[list[dict[str, Any]], list[str]]:
    try:
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    except Exception as exc:
        raise HTTPException(400, "No se pudo leer el archivo. Verifica que sea un Excel .xlsx válido.") from exc

    candidates = []
    for worksheet in workbook.worksheets:
        first_row = next(worksheet.iter_rows(min_row=1, max_row=1), ())
        headers = [canonical_header(cell.value) for cell in first_row]
        if REQUIRED_COLUMNS.issubset(set(headers)):
            candidates.append((worksheet, headers))
    if len(candidates) != 1:
        workbook.close()
        if not candidates:
            raise HTTPException(
                422,
                "No se encontró una hoja con el detalle esperado. Debe incluir columnas como "
                "FECHA GESTIÓN, EMPRESA, ÁREA, VENDEDOR, CLASIFICACIÓN y MOMENTO DÍAS.",
            )
        raise HTTPException(422, "El archivo contiene más de una hoja de detalle compatible; deja solo una.")

    worksheet, headers = candidates[0]
    nonempty_headers = [header for header in headers if header]
    if len(nonempty_headers) != len(set(nonempty_headers)):
        workbook.close()
        raise HTTPException(422, "La hoja contiene encabezados duplicados.")

    rows = []
    invalid_dates = 0
    for row_number, cells in enumerate(worksheet.iter_rows(min_row=2), 2):
        if not any(cell.value is not None and str(cell.value).strip() for cell in cells):
            continue
        if len(rows) >= MAX_ROWS:
            workbook.close()
            raise HTTPException(413, f"El archivo supera el límite de {MAX_ROWS:,} filas.")
        values: dict[str, Any] = {}
        for header, cell in zip(headers, cells):
            if header:
                values[header] = cell.value
                values[f"__format__{header}"] = cell.number_format
        record = normalize_record(values, row_number)
        if any(
            values.get(column) not in (None, "")
            and not record[column]
            for column in ("FECHA GESTIÓN", "FECHA DE ACEPTACIÓN DE PRODUCTO", "FECHA PRIMER PAGO")
        ):
            invalid_dates += 1
        rows.append(record)
    workbook.close()
    if not rows:
        raise HTTPException(422, "La hoja no contiene registros para analizar.")
    warnings = []
    missing_dni = sum(not row["DNI"] for row in rows)
    missing_company = sum(not row["EMPRESA"] for row in rows)
    if missing_dni:
        warnings.append(f"{missing_dni} filas sin DNI")
    if missing_company:
        warnings.append(f"{missing_company} filas sin empresa")
    if invalid_dates:
        warnings.append(f"{invalid_dates} filas con alguna fecha no reconocida")
    return rows, warnings


def set_active(rows: list[dict[str, Any]], filename: str, warnings: list[str]) -> None:
    global ACTIVE_ROWS, ACTIVE_FILENAME, ACTIVE_WARNINGS
    ACTIVE_ROWS = rows
    ACTIVE_FILENAME = filename
    ACTIVE_WARNINGS = warnings


def csrf(request: Request) -> None:
    origin = request.headers.get("origin", "")
    expected = f'{request.url.scheme}://{request.headers.get("host", "")}'
    if not origin or origin.rstrip("/") != expected:
        raise HTTPException(403, "La solicitud debe originarse desde este panel.")


def ensure_source_loaded() -> None:
    if ACTIVE_ROWS:
        return
    if not DEFAULT_SOURCE.is_file():
        raise HTTPException(404, "Carga el archivo DETALLE FINAL.xlsx para iniciar el análisis.")
    try:
        raw = DEFAULT_SOURCE.read_bytes()
    except OSError as exc:
        raise HTTPException(500, "No se pudo leer el archivo de origen configurado.") from exc
    rows, warnings = parse_workbook(raw)
    set_active(rows, DEFAULT_SOURCE.name, warnings)


def filtered_rows(
    *,
    company: str = "",
    area: str = "",
    advisor: str = "",
    classification: str = "",
    month: str = "",
    date_from: str = "",
    date_to: str = "",
    q: str = "",
) -> list[dict[str, Any]]:
    for value in (date_from, date_to):
        if value:
            try:
                datetime.strptime(value, "%Y-%m-%d")
            except ValueError as exc:
                raise HTTPException(422, "Las fechas deben tener formato AAAA-MM-DD.") from exc
    if date_from and date_to and date_from > date_to:
        raise HTTPException(422, "La fecha inicial no puede ser posterior a la fecha final.")
    if month and not re.fullmatch(r"\d{4}-\d{2}", month):
        raise HTTPException(422, "El mes debe tener formato AAAA-MM.")
    if classification and classification not in CLASSIFICATIONS and classification != "Sin clasificación":
        raise HTTPException(422, "La clasificación seleccionada no es válida.")

    search = q.strip().casefold()
    results = []
    for row in ACTIVE_ROWS:
        row_date = row["FECHA GESTIÓN"]
        if company and row["EMPRESA"] != company:
            continue
        if area and row["ÁREA"] != area:
            continue
        if advisor and row["VENDEDOR"] != advisor:
            continue
        row_classification = row["CLASIFICACIÓN"] if row["CLASIFICACIÓN"] in CLASSIFICATIONS else "Sin clasificación"
        if classification and row_classification != classification:
            continue
        if month and not row_date.startswith(month):
            continue
        if date_from and (not row_date or row_date < date_from):
            continue
        if date_to and (not row_date or row_date > date_to):
            continue
        if search and not any(search in str(value).casefold() for value in row.values()):
            continue
        results.append(row)
    return results


@app.get("/")
def home() -> HTMLResponse:
    return HTMLResponse((BASE / "static" / "index.html").read_text(encoding="utf-8"))


@app.get("/api/data")
def data() -> dict[str, Any]:
    ensure_source_loaded()
    dates = [row["FECHA GESTIÓN"] for row in ACTIVE_ROWS if row["FECHA GESTIÓN"]]
    return {
        "filename": ACTIVE_FILENAME,
        "rows": ACTIVE_ROWS,
        "warnings": ACTIVE_WARNINGS,
        "period": {"from": min(dates, default=""), "to": max(dates, default="")},
    }


@app.post("/api/analyze")
async def upload(request: Request, file: UploadFile = File(...)) -> dict[str, Any]:
    csrf(request)
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Selecciona un archivo .xlsx.")
    raw = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "El archivo supera el límite de 50 MB.")
    rows, warnings = parse_workbook(raw)
    set_active(rows, Path(file.filename).name, warnings)
    dates = [row["FECHA GESTIÓN"] for row in rows if row["FECHA GESTIÓN"]]
    return {
        "filename": ACTIVE_FILENAME,
        "rows": ACTIVE_ROWS,
        "warnings": warnings,
        "period": {"from": min(dates, default=""), "to": max(dates, default="")},
    }


def add_summary_sheet(workbook: Workbook, rows: list[dict[str, Any]]) -> None:
    summary = summarize(rows)
    ws = workbook.active
    ws.title = "Resumen"
    ws.append(["INDICADOR", "RESULTADO"])
    labels = [
        ("Registros del filtro", "total"),
        ("Intentos de anulación (salvadas + caídas)", "intentos_anulacion"),
        ("Ventas salvadas", "salvadas"),
        ("Ventas caídas", "caidas"),
        ("Otras gestiones", "otras"),
        ("Sin clasificación", "sin_clasificacion"),
        ("Ventas salvadas sobre el total de casos (%)", "tasa_salvadas"),
        ("Salvadas dentro de 3 días (columna MOMENTO DÍAS)", "dentro_3_dias"),
        ("Salvadas después de 3 días (columna MOMENTO DÍAS)", "fuera_3_dias"),
        ("Salvadas sin dato de plazo", "plazo_sin_dato"),
        ("Salvadas antes del primer pago (columna MOMENTO PRIMER PAGO)", "antes_primer_pago"),
        ("Salvadas después del primer pago (columna MOMENTO PRIMER PAGO)", "despues_primer_pago"),
        ("Salvadas sin dato de momento del pago", "momento_pago_sin_dato"),
        ("Registros con fecha de primer pago", "registros_con_fecha_primer_pago"),
    ]
    for label, key in labels:
        ws.append([label, summary[key]])
    ws.append([])
    ws.append(["Criterios"])
    ws.append(["Los plazos y momentos de pago se cuentan usando los valores que aparecen en las columnas del Excel."])
    ws.append(["Todos los porcentajes usan como denominador el total de registros de su grupo o filtro."])
    ws.append(["Las fechas sin información no se interpretan ni se completan."])
    style_sheet(ws)


def add_group_sheet(workbook: Workbook, title: str, rows: list[dict[str, Any]], fields: list[str]) -> None:
    ws = workbook.create_sheet(title)
    headers = [
        *fields,
        "Registros",
        "Salvadas",
        "Caídas",
        "Otras gestiones",
        "Tasa salvadas (%)",
        "Salvadas dentro de 3 días",
        "Salvadas después de 3 días",
        "Antes del primer pago",
        "Después del primer pago",
    ]
    ws.append(headers)
    groups: dict[tuple[str, ...], list[dict[str, Any]]] = {}
    for row in rows:
        key = tuple(
            (row.get("FECHA GESTIÓN", "")[:7] if field == "MES" else row.get(field, "")) or "Sin dato"
            for field in fields
        )
        groups.setdefault(key, []).append(row)
    for key in sorted(groups, key=lambda item: tuple(part.casefold() for part in item)):
        values = summarize(groups[key])
        ws.append([
            *key,
            values["total"],
            values["salvadas"],
            values["caidas"],
            values["otras"],
            values["tasa_salvadas"],
            values["dentro_3_dias"],
            values["fuera_3_dias"],
            values["antes_primer_pago"],
            values["despues_primer_pago"],
        ])
    style_sheet(ws)


def style_sheet(ws: Any) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="173B5E")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 30
    for column in ws.columns:
        width = min(max(max(len(str(cell.value or "")) for cell in column) + 2, 12), 42)
        ws.column_dimensions[column[0].column_letter].width = width


@app.get("/api/export")
def export(
    company: str = "",
    area: str = "",
    advisor: str = "",
    classification: str = "",
    month: str = "",
    date_from: str = "",
    date_to: str = "",
    q: str = "",
) -> StreamingResponse:
    ensure_source_loaded()
    rows = filtered_rows(
        company=company,
        area=area,
        advisor=advisor,
        classification=classification,
        month=month,
        date_from=date_from,
        date_to=date_to,
        q=q,
    )
    workbook = Workbook()
    add_summary_sheet(workbook, rows)
    add_group_sheet(workbook, "Empresa y área", rows, ["EMPRESA", "ÁREA"])
    add_group_sheet(workbook, "Asesor", rows, ["VENDEDOR", "EMPRESA", "ÁREA"])
    add_group_sheet(workbook, "Mes", rows, ["MES"])
    detail = workbook.create_sheet("Detalle")
    headers = [column for column in SOURCE_COLUMNS if column != "CONTRATOS"] + ["FILA EXCEL"]
    detail.append(headers)
    for row in rows:
        detail.append([row.get(header, "") for header in headers])
    style_sheet(detail)
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    filename = "reporte_detalle_final.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
