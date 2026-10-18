"""Domain objects shared by workbook readers and validation."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Generic, TypeVar

DateValue = date | datetime | str | None


@dataclass(frozen=True)
class ValidationIssue:
    """A data-quality finding tied to a source workbook location."""

    code: str
    message: str
    source: str
    row: int | None = None
    severity: str = "warning"


@dataclass(frozen=True)
class SapOrder:
    """One source row from the SAP purchase-order tracking sheet."""

    source_row: int
    order_id: str | None
    ot: str | None
    state: str
    expected_date: DateValue = None
    actual_date: DateValue = None
    description: str = ""


@dataclass(frozen=True)
class ProductionOperation:
    """One operation row belonging to a production order (OT)."""

    source_row: int
    ot: str
    description: str = ""
    client: str = ""
    machine: str = ""
    budget_hours: object = None


T = TypeVar("T")


@dataclass
class ReadResult(Generic[T]):
    """Parsed records plus explicit data-quality issues."""

    records: list[T]
    issues: list[ValidationIssue]

    @property
    def has_errors(self) -> bool:
        return any(issue.severity == "error" for issue in self.issues)
