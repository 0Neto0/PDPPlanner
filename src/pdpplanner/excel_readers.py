"""Read-only Excel adapters. These functions never save or modify source workbooks."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from pdpplanner.models import ProductionOperation, ReadResult, SapOrder, ValidationIssue
from pdpplanner.order_status import classify_state, normalize_ot

SAP_SHEET = "2024"
SAP_HEADER_ROW = 5
SAP_FIRST_DATA_ROW = SAP_HEADER_ROW + 1
PRODUCTION_SHEET = "PdP PRODUCCIÓN"
PRODUCTION_HEADER_ROW = 19
PRODUCTION_FIRST_DATA_ROW = 29

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


def _normalize_header(value: Any) -> str:
    return " ".join(str(value or "").strip().upper().split())


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
        header_values = next(
            worksheet.iter_rows(
                min_row=header_row,
                max_row=header_row,
                max_col=max(SAP_COLUMNS.values()),
                values_only=True,
            ),
            (),
        )
        for field, column in SAP_COLUMNS.items():
            actual = header_values[column - 1] if len(header_values) >= column else None
            expected = EXPECTED_HEADERS[field]
            actual_text = _normalize_header(actual)
            expected_text = _normalize_header(expected)
            if field == "order_id":
                actual_text = actual_text.replace("N°", "Nº").replace("NO.", "Nº")
            if actual_text != expected_text:
                issues.append(
                    ValidationIssue(
                        code="SAP_HEADER_MISMATCH",
                        message=(
                            f"Encabezado inesperado en columna {column}, fila {header_row}: "
                            f"{actual!r}; se esperaba {expected!r}."
                        ),
                        source=source,
                        row=header_row,
                    )
                )

        for row_number, values in enumerate(
            worksheet.iter_rows(
                min_row=header_row + 1,
                max_col=max(SAP_COLUMNS.values()),
                values_only=True,
            ),
            start=header_row + 1,
        ):
            raw_values = {
                field: values[column - 1] if len(values) >= column else None
                for field, column in SAP_COLUMNS.items()
            }
            if all(value is None or not str(value).strip() for value in raw_values.values()):
                continue

            order_id = _clean_identifier(raw_values["order_id"])
            raw_ot = _clean_identifier(raw_values["ot"])
            ot = normalize_ot(raw_ot) if raw_ot else None
            state = str(raw_values["state"] or "").strip()
            record = SapOrder(
                source_row=row_number,
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
                        row=row_number,
                    )
                )
            if not order_id:
                issues.append(
                    ValidationIssue(
                        code="SAP_MISSING_ORDER_ID",
                        message="El registro SAP no tiene número de pedido.",
                        source=source,
                        row=row_number,
                    )
                )
            if classify_state(state) == "unknown":
                issues.append(
                    ValidationIssue(
                        code="SAP_UNKNOWN_STATE",
                        message=f"Estado SAP no reconocido: {state!r}.",
                        source=source,
                        row=row_number,
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


def read_production_operations(
    path: str | Path,
    *,
    sheet_name: str = PRODUCTION_SHEET,
    header_row: int = PRODUCTION_HEADER_ROW,
    first_data_row: int = PRODUCTION_FIRST_DATA_ROW,
) -> ReadResult[ProductionOperation]:
    """Read operation rows from ``PdP PRODUCCIÓN`` without assuming 8 rows per OT.

    The worksheet's operation/machine label is in column E. OT values are read
    from column B; when Excel has merged the OT cell across operation rows,
    the last explicit OT is carried only while consecutive operation rows exist.
    Timeline cells and formulas are deliberately not interpreted by this reader.
    """
    source = str(path)
    records: list[ProductionOperation] = []
    issues: list[ValidationIssue] = []
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            return ReadResult(
                records=[],
                issues=[ValidationIssue(
                    code="PRODUCTION_SHEET_MISSING",
                    message=f"No existe la hoja requerida {sheet_name!r}.",
                    source=source,
                    severity="error",
                )],
            )

        worksheet = workbook[sheet_name]
        headers = next(
            worksheet.iter_rows(
                min_row=header_row, max_row=header_row, min_col=2, max_col=6,
                values_only=True,
            ),
            (),
        )
        expected = {0: "Nº OT", 1: "TIPO MÁQUINA", 4: "H. PRESUP."}
        for index, expected_header in expected.items():
            actual = headers[index] if len(headers) > index else None
            if _normalize_header(actual).replace("N°", "Nº") != _normalize_header(expected_header):
                issues.append(ValidationIssue(
                    code="PRODUCTION_HEADER_MISMATCH",
                    message=(f"Encabezado inesperado en columna {index + 2}, fila {header_row}: "
                             f"{actual!r}; se esperaba {expected_header!r}."),
                    source=source,
                    row=header_row,
                ))

        active_ot: str | None = None
        for row_number, values in enumerate(
            worksheet.iter_rows(min_row=first_data_row, min_col=2, max_col=6, values_only=True),
            start=first_data_row,
        ):
            # Values are B:F: OT, description/type, client, machine, budget hours.
            raw_ot, description, client, machine, budget_hours = (
                list(values) + [None] * 5
            )[:5]
            machine_text = str(machine or "").strip()
            if raw_ot is not None and str(raw_ot).strip():
                raw_ot_text = str(raw_ot).strip()
                if raw_ot_text.upper() not in {"OT", "Nº OT", "N° OT"}:
                    active_ot = normalize_ot(raw_ot_text)
            if not machine_text:
                # A blank row breaks a merged-cell carry-forward block.
                if not any(value is not None and str(value).strip() for value in values):
                    active_ot = None
                continue
            if not active_ot:
                issues.append(ValidationIssue(
                    code="PRODUCTION_OPERATION_WITHOUT_OT",
                    message="Hay una operación con máquina pero sin OT identificable.",
                    source=source,
                    row=row_number,
                ))
                continue
            records.append(ProductionOperation(
                source_row=row_number,
                ot=active_ot,
                description=str(description or "").strip(),
                client=str(client or "").strip(),
                machine=machine_text,
                budget_hours=budget_hours,
            ))
    finally:
        workbook.close()
    return ReadResult(records=records, issues=issues)
