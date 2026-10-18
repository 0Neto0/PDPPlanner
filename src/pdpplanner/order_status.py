"""Helpers for classifying SAP purchase-order states and calculating OT summaries."""

from __future__ import annotations

import unicodedata
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from typing import Any


def normalize_text(value: Any) -> str:
    """Normalize text for comparisons (trim, uppercase, remove accents)."""
    if value is None:
        return ""
    text = " ".join(str(value).strip().split()).upper()
    return "".join(
        char
        for char in unicodedata.normalize("NFD", text)
        if unicodedata.category(char) != "Mn"
    )


PENDING_STATES = frozenset(
    {
        "EN TRAMITE",
        "PARCIALMENTE RECIBIDO",
        "LISTO PARA RECOGER",
        "PENDIENTE COMPROBAR CALIDAD",
        "RECIBIDO ERRONEO",
    }
)
RECEIVED_STATES = frozenset({"RECIBIDO"})
CANCELLED_STATES = frozenset({"ANULADO"})


def classify_state(value: Any) -> str:
    """Return 'pending', 'received', 'cancelled', or 'unknown'."""
    state = normalize_text(value)
    if state in PENDING_STATES:
        return "pending"
    if state in RECEIVED_STATES:
        return "received"
    if state in CANCELLED_STATES:
        return "cancelled"
    return "unknown"


def normalize_ot(value: Any) -> str:
    """Normalize OT identifiers without conflating non-numeric identifiers."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    # Excel can expose integer-valued numeric cells as e.g. 2300156.0.
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except ValueError:
        pass
    return normalize_text(text)


@dataclass(frozen=True)
class OrderRecord:
    ot: Any
    order_id: Any
    state: Any
    expected_date: Any = None
    actual_date: Any = None
    source_row: int | None = None


@dataclass(frozen=True)
class OrderSummary:
    ot: str
    total_orders: int
    pending_orders: int
    unknown_orders: int

    @property
    def pending_percent(self) -> float | None:
        if self.total_orders == 0:
            return None
        return self.pending_orders / self.total_orders * 100


def summarize_orders(records: Iterable[OrderRecord]) -> dict[str, OrderSummary]:
    """Summarize unique SAP orders by OT.

    Cancelled orders are excluded. Unknown states remain in the denominator and
    are surfaced separately so they cannot silently disappear from the metric.
    For a blank SAP order ID, the source row is used as a fallback identity; if
    no row number was supplied, each such record gets a unique in-memory key.
    """
    grouped: dict[str, dict[str, tuple[str, OrderRecord]]] = {}
    anonymous_counter = 0

    for record in records:
        ot = normalize_ot(record.ot)
        if not ot:
            continue

        classification = classify_state(record.state)
        if classification == "cancelled":
            continue

        order_id = str(record.order_id).strip() if record.order_id is not None else ""
        if order_id:
            identity = normalize_ot(order_id)
        elif record.source_row is not None:
            identity = f"__ROW_{record.source_row}"
        else:
            anonymous_counter += 1
            identity = f"__ANONYMOUS_{anonymous_counter}"

        by_order = grouped.setdefault(ot, {})
        previous = by_order.get(identity)
        # If duplicate rows exist for the same OT/order, keep pending over
        # received, and unknown only when no recognized state exists.
        priority = {"unknown": 0, "received": 1, "pending": 2}
        if previous is None or priority[classification] > priority[previous[0]]:
            by_order[identity] = (classification, record)

    summaries: dict[str, OrderSummary] = {}
    for ot, orders in grouped.items():
        classifications = [classification for classification, _ in orders.values()]
        summaries[ot] = OrderSummary(
            ot=ot,
            total_orders=len(orders),
            pending_orders=sum(state == "pending" for state in classifications),
            unknown_orders=sum(state == "unknown" for state in classifications),
        )
    return summaries
