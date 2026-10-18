"""Read-only Excel adapters. These functions never save or modify source workbooks."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from pdpplanner.models import ReadResult, SapOrder, ValidationIssue
from pdpplanner.order_status import classify_state, normalize_ot

SAP_SHEET = "2024"
SAP_HEADER_ROW = 5
SAP_FIRST_DATA_ROW = SAP_HEADER_ROW + 1

# Column positions are part of the agreed input contract. Column C is
# intentionally absent: it is the supplier-response date and must be ignored.
SAP_COLUMNS = {
    "expected_date": 2,  # B: plazo de entrega
    "actual_date": 4,  # D: fecha de entrega
    "order_id": 5,  # E: número de pedido SAP
    "description": 7,  # G: descripción
    "ot": 8,  # H: O.T.
    "state": 10,  # J: estado
}
EXPECTED_HEADERS = {
    "expected_date": "PLAZO DE ENTREGA",
    "actual_date": "FECHA DE ENTREGA",
    "order_id": "Nº PEDIDO SAP",
    "description": "DESCRIPCION",
    "ot": "O.T.",
    "state": "ESTADO",
}


def _clean_identifier(value: Any) -> str | None:
    if value is None or not str(value).strip():
        return None
    return str(value).strip()


def read_sap_orders(
    path: str | Path,
    *,
    sheet_name: str = SAP_SHEET,
    header_row: int = SAP_HEADER_ROW,
) -> ReadResult[SapOrder]:
    """Read SAP purchase-order rows and return warnings for ambiguous data.

    Rows without an OT or SAP order ID are retained in the result so the
    validation report can explain them; downstream summaries must not attach
    them to an OT by inference. Duplicate order IDs associated with different
    OTs are flagged and left unmerged.
    """
    source = str(path)
    records: list[SapOrder] = []
    issues: list[ValidationIssue] = []
    workbook = load_workbook(path, read_only=True, data_only=True)

    try:
        if sheet_name not in workbook.sheetnames:
            return ReadResult(
                records=[],
                issues=[
                    ValidationIssue(
                        code="SAP_SHEET_MISSING",
                        message=f"No existe la hoja requerida {sheet_name!r}.",
                        source=source,
                        severity="error",
                    )
                ],
            )

        worksheet = workbook[sheet_name]
        for field, column in SAP_COLUMNS.items():
            actual = worksheet.cell(header_row, column).value
            expected = EXPECTED_HEADERS[field]
            actual_text = " ".join(str(actual or "").strip().upper().split())
            expected_text = " ".join(expected.upper().split())
            # Some exported Excel files use Nº/N°/No variants.
            if field == "order_id":
                header_ok = actual_text.replace("N°", "Nº").replace("NO.", "Nº") == expected_text
            else:
                header_ok = actual_text == expected_text
            if not header_ok:
                issues.append(
                    ValidationIssue(
                        code="SAP_HEADER_MISMATCH",
                        message=(
                            f"Encabezado inesperado en {worksheet.cell(header_row, column).coordinate}: "
                            f"{actual!r}; se esperaba {expected!r}."
                        ),
                        source=source,
                        row=header_row,
                    )
                )

        for row in range(header_row + 1, worksheet.max_row + 1):
            raw_values = {
                field: worksheet.cell(row, column).value
                for field, column in SAP_COLUMNS.items()
            }
            if all(value is None or not str(value).strip() for value in raw_values.values()):
                continue

            order_id = _clean_identifier(raw_values["order_id"])
            raw_ot = _clean_identifier(raw_values["ot"])
            ot = normalize_ot(raw_ot) if raw_ot else None
            state = str(raw_values["state"] or "").strip()
            record = SapOrder(
                source_row=row,
                order_id=order_id,
                ot=ot,
                state=state,
                expected_date=raw_values["expected_date"],
                actual_date=raw_values["actual_date"],
                description=str(raw_values["description"] or "").strip(),
            )
            records.append(record)

            if not ot:
                issues.append(
                    ValidationIssue(
                        code="SAP_MISSING_OT",
                        message="El registro SAP no tiene OT; no se asociará por inferencia.",
                        source=source,
                        row=row,
                    )
                )
            if not order_id:
                issues.append(
                    ValidationIssue(
                        code="SAP_MISSING_ORDER_ID",
                        message="El registro SAP no tiene número de pedido.",
                        source=source,
                        row=row,
                    )
                )
            classification = classify_state(state)
            if classification == "unknown":
                issues.append(
                    ValidationIssue(
                        code="SAP_UNKNOWN_STATE",
                        message=f"Estado SAP no reconocido: {state!r}.",
                        source=source,
                        row=row,
                    )
                )

        order_rows: dict[str, list[SapOrder]] = defaultdict(list)
        for record in records:
            if record.order_id:
                order_rows[record.order_id].append(record)
        for order_id, duplicates in order_rows.items():
            ots = {record.ot for record in duplicates if record.ot}
            if len(ots) > 1:
                row_list = ", ".join(str(record.source_row) for record in duplicates)
                issues.append(
                    ValidationIssue(
                        code="SAP_ORDER_ID_MULTIPLE_OTS",
                        message=(
                            f"El pedido SAP {order_id!r} aparece asociado a varias OT "
                            f"({', '.join(sorted(ots))}); filas: {row_list}. No se fusiona."
                        ),
                        source=source,
                        row=duplicates[0].source_row,
                    )
                )
    finally:
        workbook.close()

    return ReadResult(records=records, issues=issues)
