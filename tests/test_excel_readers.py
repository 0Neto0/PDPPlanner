from datetime import datetime

from openpyxl import Workbook

from pdpplanner.excel_readers import read_production_operations, read_sap_orders


def _make_sap_workbook(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "2024"
    sheet["B5"] = "PLAZO DE ENTREGA"
    sheet["D5"] = "FECHA DE ENTREGA"
    sheet["E5"] = "Nº PEDIDO SAP"
    sheet["G5"] = "DESCRIPCION"
    sheet["H5"] = "O.T."
    sheet["J5"] = "ESTADO"
    # Column C is intentionally populated to ensure the reader ignores it.
    sheet["C6"] = datetime(2026, 1, 2)
    sheet["B6"] = datetime(2026, 1, 10)
    sheet["D6"] = datetime(2026, 1, 9)
    sheet["E6"] = 4500000001
    sheet["G6"] = "MATERIAL"
    sheet["H6"] = 2300001
    sheet["J6"] = "rECIBIDO"
    sheet["E7"] = 4500000002
    sheet["G7"] = "MATERIAL"
    sheet["H7"] = 2300002
    sheet["J7"] = "ESTADO NUEVO"
    sheet["E8"] = 4500000002
    sheet["H8"] = 2300003
    sheet["J8"] = "EN TRAMITE"
    sheet["E9"] = 4500000003
    sheet["G9"] = "MATERIAL SIN OT"
    sheet["J9"] = "ESTADO NUEVO"
    workbook.save(path)


def test_read_sap_orders_preserves_rows_and_reports_issues(tmp_path):
    path = tmp_path / "sap.xlsx"
    _make_sap_workbook(path)

    result = read_sap_orders(path)

    assert len(result.records) == 4
    assert result.records[0].ot == "2300001"
    assert result.records[0].order_id == "4500000001"
    assert result.records[0].expected_date == datetime(2026, 1, 10)
    assert result.records[0].actual_date == datetime(2026, 1, 9)
    assert {issue.code for issue in result.issues} >= {
        "SAP_MISSING_OT",
        "SAP_UNKNOWN_STATE",
        "SAP_ORDER_ID_MULTIPLE_OTS",
    }


def test_missing_sheet_is_reported_as_error(tmp_path):
    path = tmp_path / "sap.xlsx"
    workbook = Workbook()
    workbook.save(path)

    result = read_sap_orders(path)

    assert result.records == []
    assert result.has_errors
    assert result.issues[0].code == "SAP_SHEET_MISSING"


def test_read_production_operations_handles_variable_ot_block_lengths(tmp_path):
    path = tmp_path / "production.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "PdP PRODUCCIÓN"
    sheet["B19"] = "Nº OT"
    sheet["C19"] = "TIPO MÁQUINA"
    sheet["F19"] = "H. PRESUP."
    # OT 1001 has two operations; OT 1002 has one. The second row of 1001
    # deliberately has a blank OT cell to mimic a merged Excel cell.
    sheet["B29"] = 1001
    sheet["C29"] = "Conjunto A"
    sheet["E29"] = "CALDERERÍA"
    sheet["F29"] = 12
    sheet["C30"] = "Conjunto A"
    sheet["E30"] = "TORNO"
    sheet["F30"] = 4
    sheet["B31"] = 1002
    sheet["C31"] = "Conjunto B"
    sheet["E31"] = "FRESA"
    sheet["F31"] = 7
    workbook.save(path)

    result = read_production_operations(path)

    assert [(item.ot, item.machine, item.budget_hours) for item in result.records] == [
        ("1001", "CALDERERÍA", 12),
        ("1001", "TORNO", 4),
        ("1002", "FRESA", 7),
    ]
    assert not result.has_errors


def test_production_operation_without_ot_is_reported(tmp_path):
    path = tmp_path / "production.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "PdP PRODUCCIÓN"
    sheet["B19"] = "Nº OT"
    sheet["C19"] = "TIPO MÁQUINA"
    sheet["F19"] = "H. PRESUP."
    sheet["E29"] = "TORNO"
    workbook.save(path)

    result = read_production_operations(path)

    assert result.records == []
    assert result.issues[0].code == "PRODUCTION_OPERATION_WITHOUT_OT"


def test_missing_production_sheet_is_reported_as_error(tmp_path):
    path = tmp_path / "production.xlsx"
    Workbook().save(path)

    result = read_production_operations(path)

    assert result.records == []
    assert result.has_errors
    assert result.issues[0].code == "PRODUCTION_SHEET_MISSING"
