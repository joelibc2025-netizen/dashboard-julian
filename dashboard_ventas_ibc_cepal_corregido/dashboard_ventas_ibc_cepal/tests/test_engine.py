import io
import unittest
from datetime import datetime

import app as app_module
from fastapi import HTTPException
from openpyxl import Workbook

from app import filtered_rows, parse_workbook, set_active
from engine import SOURCE_COLUMNS, normalize_record, summarize


class NormalizeRecordTests(unittest.TestCase):
    def test_normalizes_dates_and_preserves_formatted_leading_zeroes(self):
        record = normalize_record(
            {
                "FECHA GESTIÓN": datetime(2026, 7, 1),
                "DNI": 944818,
                "__format__DNI": "00000000",
                "TELÉFONO": 948408249,
                "VENDEDOR": "  ASESOR   UNO ",
            },
            2,
        )

        self.assertEqual(record["FECHA GESTIÓN"], "2026-07-01")
        self.assertEqual(record["DNI"], "00944818")
        self.assertEqual(record["TELÉFONO"], "948408249")
        self.assertEqual(record["VENDEDOR"], "ASESOR UNO")
        self.assertEqual(record["FILA EXCEL"], 2)


class SummaryTests(unittest.TestCase):
    def test_summarizes_classification_and_saved_timing_without_reclassifying(self):
        rows = [
            {
                "CLASIFICACIÓN": "Venta Salvada",
                "MOMENTO DÍAS": "Dentro de 3 días",
                "MOMENTO PRIMER PAGO": "Antes del primer pago",
                "FECHA PRIMER PAGO": "",
            },
            {
                "CLASIFICACIÓN": "Venta Salvada",
                "MOMENTO DÍAS": "Fuera de 3 días",
                "MOMENTO PRIMER PAGO": "Después del primer pago",
                "FECHA PRIMER PAGO": "2026-08-30",
            },
            {
                "CLASIFICACIÓN": "Venta Caída",
                "MOMENTO DÍAS": "Dentro de 3 días",
                "MOMENTO PRIMER PAGO": "Antes del primer pago",
                "FECHA PRIMER PAGO": "",
            },
            {"CLASIFICACIÓN": ""},
        ]

        result = summarize(rows)

        self.assertEqual(result["total"], 4)
        self.assertEqual(result["salvadas"], 2)
        self.assertEqual(result["caidas"], 1)
        self.assertEqual(result["intentos_anulacion"], result["salvadas"] + result["caidas"])
        self.assertEqual(result["intentos_anulacion"], 3)
        self.assertEqual(result["sin_clasificacion"], 1)
        self.assertEqual(result["tasa_salvadas"], 50.0)
        self.assertEqual(result["dentro_3_dias"], 1)
        self.assertEqual(result["fuera_3_dias"], 1)
        self.assertEqual(result["antes_primer_pago"], 1)
        self.assertEqual(result["despues_primer_pago"], 1)
        self.assertEqual(result["registros_con_fecha_primer_pago"], 1)

    def test_does_not_infer_missing_timing_values(self):
        result = summarize([{"CLASIFICACIÓN": "Venta Salvada"}])

        self.assertEqual(result["plazo_sin_dato"], 1)
        self.assertEqual(result["momento_pago_sin_dato"], 1)
        self.assertEqual(result["tasa_salvadas"], 100)


class WorkbookTests(unittest.TestCase):
    def test_reads_single_detail_sheet_and_retains_source_row(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "DETALLE FINAL"
        worksheet.append(SOURCE_COLUMNS)
        values = {column: "" for column in SOURCE_COLUMNS}
        values.update(
            {
                "FECHA GESTIÓN": datetime(2026, 7, 1),
                "EMPRESA": "IBC DIGITAL",
                "ÁREA": "VENTAS",
                "VENDEDOR": "ASESOR UNO",
                "CLIENTE": "CLIENTE DE PRUEBA",
                "DNI": 12345,
                "CLASIFICACIÓN": "Venta Salvada",
                "MOMENTO DÍAS": "Dentro de 3 días",
                "FECHA PRIMER PAGO": "",
                "MOMENTO PRIMER PAGO": "Antes del primer pago",
            }
        )
        worksheet.append([values[column] for column in SOURCE_COLUMNS])
        worksheet["F2"].number_format = "00000000"
        output = io.BytesIO()
        workbook.save(output)

        records, warnings = parse_workbook(output.getvalue())

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["DNI"], "00012345")
        self.assertEqual(records[0]["FECHA GESTIÓN"], "2026-07-01")
        self.assertEqual(records[0]["FILA EXCEL"], 2)
        self.assertEqual(warnings, [])


class FilterTests(unittest.TestCase):
    def setUp(self):
        self.previous_rows = app_module.ACTIVE_ROWS
        self.previous_filename = app_module.ACTIVE_FILENAME
        self.previous_warnings = app_module.ACTIVE_WARNINGS
        set_active(
            [
                {"EMPRESA": "IBC DIGITAL", "ÁREA": "VENTAS", "VENDEDOR": "ASESOR A",
                 "CLASIFICACIÓN": "Venta Salvada", "FECHA GESTIÓN": "2026-07-01", "DNI": "123"},
                {"EMPRESA": "CEPAL", "ÁREA": "ATENCION AL CLIENTE", "VENDEDOR": "ASESOR B",
                 "CLASIFICACIÓN": "Venta Caída", "FECHA GESTIÓN": "2026-07-02", "DNI": "456"},
            ],
            "test.xlsx",
            [],
        )

    def tearDown(self):
        set_active(self.previous_rows, self.previous_filename, self.previous_warnings)

    def test_company_area_and_inclusive_date_filters_combine(self):
        rows = filtered_rows(
            company="IBC DIGITAL",
            area="VENTAS",
            date_from="2026-07-01",
            date_to="2026-07-01",
        )

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["DNI"], "123")

    def test_rejects_reversed_date_range(self):
        with self.assertRaises(HTTPException) as context:
            filtered_rows(date_from="2026-07-02", date_to="2026-07-01")
        self.assertEqual(getattr(context.exception, "status_code", None), 422)


if __name__ == "__main__":
    unittest.main()
