"""Load the four workbook sources and consolidate data-quality findings."""

from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

from pdpplanner.excel_readers import (
    read_production_operations,
    read_production_order_list,
    read_sap_orders,
    read_staff_calendar,
)
from pdpplanner.models import ReadResult, SapOrder, ProductionOperation, ProductionOrderInfo, StaffMember, ValidationIssue


@dataclass
class SourceBundle:
    """Parsed source records, kept separate by workbook/sheet responsibility."""

    production_operations: ReadResult[ProductionOperation]
    production_orders: ReadResult[ProductionOrderInfo]
    staff_calendar: ReadResult[StaffMember]
    sap_orders: ReadResult[SapOrder]

    @property
    def issues(self) -> list[ValidationIssue]:
        """Return all validation findings in a stable source order."""
        return [
            *self.production_operations.issues,
            *self.production_orders.issues,
            *self.staff_calendar.issues,
            *self.sap_orders.issues,
        ]

    @property
    def has_errors(self) -> bool:
        return any(issue.severity == "error" for issue in self.issues)

    def summary(self) -> dict[str, Any]:
        """Build a JSON-serializable count summary without exposing row data."""
        by_source = {
            "production_operations": self.production_operations,
            "production_orders": self.production_orders,
            "staff_calendar": self.staff_calendar,
            "sap_orders": self.sap_orders,
        }
        issue_codes = Counter(issue.code for issue in self.issues)
        severities = Counter(issue.severity for issue in self.issues)
        return {
            "record_counts": {name: len(result.records) for name, result in by_source.items()},
            "issue_count": len(self.issues),
            "issues_by_severity": dict(sorted(severities.items())),
            "issues_by_code": dict(sorted(issue_codes.items())),
            "has_errors": self.has_errors,
        }

    def report(self) -> dict[str, Any]:
        """Return a machine-readable summary and all findings for CLI/API use."""
        return {
            "summary": self.summary(),
            "issues": [asdict(issue) for issue in self.issues],
        }


def read_and_validate_sources(
    management_workbook: str | Path,
    sap_workbook: str | Path,
) -> SourceBundle:
    """Read all four source areas without saving or changing either workbook.

    ``management_workbook`` provides the production schedule, order list and
    staff calendar. ``sap_workbook`` provides the SAP sheet named ``2024``.
    Findings are retained rather than repaired silently.
    """
    return SourceBundle(
        production_operations=read_production_operations(management_workbook),
        production_orders=read_production_order_list(management_workbook),
        staff_calendar=read_staff_calendar(management_workbook),
        sap_orders=read_sap_orders(sap_workbook),
    )
