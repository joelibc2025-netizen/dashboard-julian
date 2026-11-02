"""Normalization and reporting rules for the single-source sales workbook."""

from collections import Counter
from datetime import date, datetime, timedelta
import re
from typing import Any, Iterable


SOURCE_COLUMNS = [
    "FECHA GESTIÓN",
    "EMPRESA",
    "ÁREA",
    "VENDEDOR",
    "CLIENTE",
    "DNI",
    "TELÉFONO",
    "REGISTRO 1",
    "REGISTRO 2",
    "REGISTRO 3",
    "ESTADO",
    "CANTIDAD DE COMUNICACIONES DEL MISMO CLIENTE",
    "CONTRATOS",
    "CLIENTE ACEPTA PRODUCTO",
    "FECHA DE ACEPTACIÓN DE PRODUCTO",
    "COMUNICACIÓN DE VENTA",
    "CLASIFICACIÓN",
    "MOMENTO DÍAS",
    "FECHA PRIMER PAGO",
    "MOMENTO PRIMER PAGO",
]

DATE_COLUMNS = {
    "FECHA GESTIÓN",
    "FECHA DE ACEPTACIÓN DE PRODUCTO",
    "FECHA PRIMER PAGO",
}
CLASSIFICATIONS = ("Venta Salvada", "Venta Caída", "Otra Gestión")


def canonical_header(value: Any) -> str:
    text = str(value or "").strip().upper()
    text = re.sub(r"\s+", " ", text)
    return text


def parse_date(value: Any) -> str:
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    if isinstance(value, (int, float)):
        try:
            parsed = date(1899, 12, 30) + timedelta(days=float(value))
            return parsed.isoformat() if date(1900, 1, 1) <= parsed <= date(2200, 12, 31) else ""
        except (OverflowError, ValueError):
            return ""
    text = str(value or "").strip()
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(text[:10], pattern).date().isoformat()
        except ValueError:
            continue
    return ""


def normalize_cell(column: str, value: Any, number_format: str = "General") -> str:
    if value is None:
        return ""
    if column in DATE_COLUMNS:
        return parse_date(value)
    if column in {"DNI", "TELÉFONO"} and isinstance(value, (int, float)):
        try:
            number = int(value)
            if number == value:
                digits = str(number)
                zeros = re.fullmatch(r"0+", number_format)
                if zeros:
                    digits = digits.zfill(len(number_format))
                return digits
        except (OverflowError, ValueError):
            pass
    if isinstance(value, bool):
        return "Sí" if value else "No"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ").strip())


def normalize_record(values: dict[str, Any], row_number: int) -> dict[str, Any]:
    row: dict[str, Any] = {}
    for column in SOURCE_COLUMNS:
        row[column] = normalize_cell(column, values.get(column), values.get(f"__format__{column}", "General"))
    row["FILA EXCEL"] = row_number
    return row


def _classification(row: dict[str, Any]) -> str:
    value = row.get("CLASIFICACIÓN", "").strip()
    return value if value in CLASSIFICATIONS else "Sin clasificación"


def summarize(rows: Iterable[dict[str, Any]]) -> dict[str, Any]:
    records = list(rows)
    classifications = Counter(_classification(row) for row in records)
    saved = [row for row in records if _classification(row) == "Venta Salvada"]
    total_resolved = classifications["Venta Salvada"] + classifications["Venta Caída"]
    timings = Counter(row.get("MOMENTO DÍAS", "") or "Sin dato" for row in saved)
    payments = Counter(row.get("MOMENTO PRIMER PAGO", "") or "Sin dato" for row in saved)
    first_payment_dates = sum(bool(row.get("FECHA PRIMER PAGO")) for row in records)
    return {
        "total": len(records),
        "salvadas": classifications["Venta Salvada"],
        "caidas": classifications["Venta Caída"],
        "intentos_anulacion": total_resolved,
        "otras": classifications["Otra Gestión"],
        "sin_clasificacion": classifications["Sin clasificación"],
        "tasa_salvadas": round(classifications["Venta Salvada"] / len(records) * 100, 1) if records else 0,
        "dentro_3_dias": timings["Dentro de 3 días"],
        "fuera_3_dias": timings["Fuera de 3 días"],
        "plazo_sin_dato": sum(count for label, count in timings.items() if label not in ("Dentro de 3 días", "Fuera de 3 días")),
        "antes_primer_pago": payments["Antes del primer pago"],
        "despues_primer_pago": payments["Después del primer pago"],
        "momento_pago_sin_dato": sum(count for label, count in payments.items() if label not in ("Antes del primer pago", "Después del primer pago")),
        "registros_con_fecha_primer_pago": first_payment_dates,
    }


def group_counts(rows: Iterable[dict[str, Any]], field: str) -> list[dict[str, Any]]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        value = row.get(field, "").strip() or "Sin dato"
        groups.setdefault(value, []).append(row)
    result = []
    for value, group in groups.items():
        summary = summarize(group)
        result.append({"grupo": value, **summary})
    return sorted(result, key=lambda item: (-item["total"], item["grupo"].casefold()))
