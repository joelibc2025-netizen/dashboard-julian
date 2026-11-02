# Repository Guide

## Project Structure

- `app.py` defines the FastAPI application, single-workbook loading, in-memory analysis, filters, and Excel export.
- `engine.py` contains normalization and calculations for the `DETALLE FINAL.xlsx` detail sheet.
- `static/` contains the HTML, CSS, and browser JavaScript.
- `tests/test_engine.py` covers workbook normalization, filtering, and sales indicators.
- Read `README.md` before changing business rules or deployment behavior.

## Working Conventions

* Keep user-facing text in Spanish and preserve the single-source workbook workflow.
* Preserve the classification values in the workbook. Keep calculations in `engine.py` and cover rule changes with focused tests in `tests/test_engine.py`.
* Do not infer missing business data or silently replace missing dates, IDs, or payment moments.
* Do not add customer data or uploaded workbooks to the repository. The dashboard processes records in memory and exports only when requested.
* The dashboard is local-only and has no authentication. Bind it to `127.0.0.1`; do not expose it to a network or the Internet.
* Treat any legacy `data.sqlite3` and its `-wal`/`-shm` files as live user data; do not modify, delete, or regenerate them.
* The application has no external service dependency. Avoid introducing one without an explicit requirement.

## Verification

Run the test suite from this directory:

```powershell
python -m unittest discover -s tests -v
```

For browser changes, manually check the dashboard at desktop and mobile sizes.
