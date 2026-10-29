from decimal import Decimal

import pandas as pd
import pytest

from comparison import (
    BOTH, CORRECT, INVOICE_ONLY, ORDER_ONLY, PRICE, QUANTITY,
    DocumentValidationError, compare_documents, format_money,
)
from excel_reader import REQUIRED_COLUMNS


def document(*rows):
    return pd.DataFrame(rows, columns=REQUIRED_COLUMNS)


@pytest.mark.parametrize("quantity,price,state,delta", [
    (2, "10.10", CORRECT, "0"),
    (2, "10.20", PRICE, "0.20"),
    (3, "10.10", QUANTITY, "10.10"),
    (3, "10.20", BOTH, "10.40"),
    (1, "10.00", BOTH, "-10.20"),
])
def test_states_and_economic_difference(quantity, price, state, delta):
    result = compare_documents(document(("A", 2, "10.10")), document(("A", quantity, price)))
    row = result.table.iloc[0]
    assert row["Estado"] == state
    assert row["Importe solicitado"] == Decimal("20.20")
    assert row["Diferencia económica"] == Decimal(delta)
    assert result.net_difference == Decimal(delta)
    assert result.total == 1
    assert result.correct == (state == CORRECT)
    assert result.different == (state != CORRECT)


def test_missing_and_extra_items_use_zero_amounts():
    result = compare_documents(document(("A", 2, 10)), document(("B", 3, 5)))
    left, right = result.table.iloc[0], result.table.iloc[1]
    assert left["Estado"] == ORDER_ONLY
    assert left["Importe facturado"] == Decimal(0)
    assert left["Cantidad facturada"] is None
    assert left["Diferencia económica"] == Decimal(-20)
    assert right["Estado"] == INVOICE_ONLY
    assert right["Importe solicitado"] == Decimal(0)
    assert right["Precio solicitado"] is None
    assert right["Diferencia económica"] == Decimal(15)
    assert result.net_difference == Decimal(-5)
    assert result.different == 2


def test_normalized_names_match_without_mutation():
    order = document(("  Producto   UNO ", 1, 10))
    invoice = document(("producto\tuno", 1, 10))
    original_order, original_invoice = order.copy(deep=True), invoice.copy(deep=True)
    result = compare_documents(order, invoice)
    assert result.correct == result.total == 1
    pd.testing.assert_frame_equal(order, original_order)
    pd.testing.assert_frame_equal(invoice, original_invoice)


@pytest.mark.parametrize("side", ["order", "invoice"])
def test_duplicates_after_normalization_block_comparison(side):
    duplicated = document((" A  B", 1, 1), ("a b ", 2, 2))
    args = (duplicated, document()) if side == "order" else (document(), duplicated)
    with pytest.raises(DocumentValidationError, match="duplicado.*fila 2"):
        compare_documents(*args)


@pytest.mark.parametrize("value", [-1, "-0.01", "texto", "", None, float("nan"),
                                       float("inf"), "-Infinity", True, "1,50"])
@pytest.mark.parametrize("column", ["Cantidad", "Precio unitario"])
def test_invalid_numbers_block_comparison(value, column):
    invalid = document(("A", 1, 1)).astype(object)
    invalid.loc[0, column] = value
    with pytest.raises(DocumentValidationError, match=column):
        compare_documents(document(), invalid)


@pytest.mark.parametrize("name", [None, "", "  ", 123])
def test_invalid_product_names(name):
    with pytest.raises(DocumentValidationError, match="Producto debe ser un texto no vacío"):
        compare_documents(document((name, 1, 1)), document())


def test_errors_from_both_documents_include_rows_and_labels():
    with pytest.raises(DocumentValidationError) as caught:
        compare_documents(document(("A", -1, -2)), document(("B", "x", 1)))
    assert len(caught.value.errors) == 3
    assert "Orden de compra, fila 2" in str(caught.value)
    assert "Factura del proveedor, fila 2" in str(caught.value)


def test_no_intermediate_rounding_and_exact_totals():
    result = compare_documents(
        document(("A", "0.1", "0.2"), ("B", "3", "0.005"), ("C", 1, 5)),
        document(("A", "0.3", "0.2"), ("B", "3", "0.015"), ("C", 1, 5)),
    )
    assert list(result.table["Importe solicitado"]) == [Decimal("0.02"), Decimal("0.015"), Decimal(5)]
    assert result.net_difference == Decimal("0.070")
    assert (result.total, result.correct, result.different) == (3, 1, 2)
    assert format_money(result.net_difference) == "0.07"


def test_precision_greater_than_default_decimal_context():
    price = "123456789012345678901234567890.123456789"
    result = compare_documents(document(), document(("A", "2", price)))
    assert result.net_difference == Decimal("246913578024691357802469135780.246913578")
    assert format_money(result.net_difference) == "246913578024691357802469135780.25"


def test_totals_are_not_sum_of_display_rounded_values():
    result = compare_documents(document(), document(("A", 1, "0.005"), ("B", 1, "0.005")))
    assert result.net_difference == Decimal("0.010")
    assert format_money(result.net_difference) == "0.01"
    assert format_money(Decimal("-0.005")) == "-0.01"


def test_zero_and_empty_documents():
    empty = compare_documents(document(), document())
    assert empty.total == empty.correct == empty.different == 0
    assert empty.net_difference == Decimal(0)
    zero = compare_documents(document(("A", 0, 0)), document(("A", 0, 1)))
    assert zero.table.iloc[0]["Estado"] == PRICE
    assert zero.net_difference == Decimal(0)
