import unittest

from pdpplanner.order_status import (
    OrderRecord,
    classify_state,
    normalize_ot,
    normalize_text,
    summarize_orders,
)


class OrderStatusTests(unittest.TestCase):
    def test_state_normalization_and_classification(self):
        self.assertEqual(normalize_text("  recibido  "), "RECIBIDO")
        self.assertEqual(normalize_text("pendiente comprobar calidad"), "PENDIENTE COMPROBAR CALIDAD")
        self.assertEqual(classify_state("EN TRAMITE"), "pending")
        self.assertEqual(classify_state("PARCIALMENTE RECIBIDO"), "pending")
        self.assertEqual(classify_state("LISTO PARA RECOGER"), "pending")
        self.assertEqual(classify_state("pendiente comprobar calidad"), "pending")
        self.assertEqual(classify_state("RECIBIDO ERRONEO"), "pending")
        self.assertEqual(classify_state("rECIBIDO"), "received")
        self.assertEqual(classify_state("ANULADO"), "cancelled")
        self.assertEqual(classify_state("estado inesperado"), "unknown")

    def test_ot_numeric_values_are_normalized(self):
        self.assertEqual(normalize_ot(2300156), "2300156")
        self.assertEqual(normalize_ot("2300156.0"), "2300156")
        self.assertEqual(normalize_ot(" stock "), "STOCK")

    def test_percentage_counts_distinct_orders_and_excludes_cancelled(self):
        records = [
            OrderRecord(2300001, 4501, "EN TRAMITE", source_row=6),
            OrderRecord("2300001", 4501, "RECIBIDO", source_row=7),
            OrderRecord(2300001, 4502, "RECIBIDO", source_row=8),
            OrderRecord(2300001, 4503, "ANULADO", source_row=9),
        ]
        summary = summarize_orders(records)["2300001"]
        self.assertEqual(summary.total_orders, 2)
        self.assertEqual(summary.pending_orders, 1)
        self.assertEqual(summary.pending_percent, 50.0)

    def test_unknown_states_are_visible_and_ot_without_orders_is_absent(self):
        summaries = summarize_orders(
            [OrderRecord(2300002, 4504, "ESTADO NUEVO", source_row=10)]
        )
        self.assertEqual(summaries["2300002"].total_orders, 1)
        self.assertEqual(summaries["2300002"].unknown_orders, 1)
        self.assertEqual(summaries["2300002"].pending_percent, 0.0)
        self.assertNotIn("2300003", summaries)


if __name__ == "__main__":
    unittest.main()
