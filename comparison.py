"""Validación y comparación exacta; no depende de Streamlit ni modifica entradas."""

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext

import pandas as pd

from excel_reader import ExcelValidationError, validate_columns


CORRECT = "Correcto"
PRICE = "Diferencia de precio"
QUANTITY = "Diferencia de cantidad"
BOTH = "Diferencia de precio y cantidad"
ORDER_ONLY = "Solo en orden de compra"
INVOICE_ONLY = "Solo en factura"
COLUMNS = [
    "Producto", "Cantidad solicitada", "Cantidad facturada", "Precio solicitado",
    "Precio facturado", "Importe solicitado", "Importe facturado",
    "Diferencia económica", "Estado",
]
MONEY_COLUMNS = COLUMNS[3:8]


class DocumentValidationError(ValueError):
    def __init__(self, errors: list[str]):
        self.errors = errors
        super().__init__("\n".join(errors))


@dataclass(frozen=True)
class Item:
    name: str
    quantity: Decimal
    price: Decimal


@dataclass
class ComparisonResult:
    table: pd.DataFrame
    total: int
    correct: int
    different: int
    net_difference: Decimal
    excess: Decimal
    shortfall: Decimal
    requested_total: Decimal
    billed_total: Decimal


def normalize_product(name: str) -> str:
    return " ".join(name.split()).casefold()


def nonnegative_decimal(value) -> Decimal:
    """Aceptar números finitos o texto decimal con punto, sin separador de miles."""
    try:
        number = Decimal(str(value).strip())
    except (InvalidOperation, ValueError):
        raise ValueError("debe ser un número no negativo (decimal con punto, sin separador de miles)") from None
    if not number.is_finite() or number < 0:
        raise ValueError("debe ser un número finito no negativo")
    return number


def validate_document(data: pd.DataFrame, label: str) -> dict[str, Item]:
    try:
        validate_columns(data)
    except ExcelValidationError as exc:
        raise DocumentValidationError([f"{label}: {exc}"]) from exc
    errors = []
    items = {}
    seen = {}
    for row_number, (_, row) in enumerate(data.iterrows(), start=2):
        prefix = f"{label}, fila {row_number}"
        name = row["Producto"]
        key = None
        if not isinstance(name, str) or not name.strip():
            errors.append(f"{prefix}: Producto debe ser un texto no vacío.")
        else:
            key = normalize_product(name)
            if key in seen:
                errors.append(f"{prefix}: producto duplicado '{name}' (también en fila {seen[key]}).")
            else:
                seen[key] = row_number
        numbers = []
        for column in ("Cantidad", "Precio unitario"):
            try:
                numbers.append(nonnegative_decimal(row[column]))
            except ValueError as exc:
                errors.append(f"{prefix}: {column} {exc}.")
        if key is not None and len(numbers) == 2:
            items[key] = Item(" ".join(name.split()), *numbers)
    if errors:
        raise DocumentValidationError(errors)
    return items


def compare_documents(order: pd.DataFrame, invoice: pd.DataFrame) -> ComparisonResult:
    documents = []
    errors = []
    for data, label in ((order, "Orden de compra"), (invoice, "Factura del proveedor")):
        try:
            documents.append(validate_document(data, label))
        except DocumentValidationError as exc:
            errors.extend(exc.errors)
    if errors:
        raise DocumentValidationError(errors)
    requested, billed = documents
    values = [number for doc in documents for item in doc.values()
              for number in (item.quantity, item.price)]
    # Reservar todos los dígitos de productos, restas y suma, incluso para
    # decimales que excedan la precisión predeterminada de Decimal (28).
    span = max((v.adjusted() for v in values), default=0) - min(
        (v.as_tuple().exponent for v in values), default=0,
    ) + 1
    with localcontext() as context:
        context.prec = max(28, 2 * span + len(str(len(values) + 1)) + 4)
        rows = []
        for key in dict.fromkeys([*requested, *billed]):
            left, right = requested.get(key), billed.get(key)
            if left is None:
                state = INVOICE_ONLY
            elif right is None:
                state = ORDER_ONLY
            else:
                price_diff = left.price != right.price
                quantity_diff = left.quantity != right.quantity
                state = BOTH if price_diff and quantity_diff else (
                    PRICE if price_diff else QUANTITY if quantity_diff else CORRECT
                )
            amount_left = left.quantity * left.price if left else Decimal(0)
            amount_right = right.quantity * right.price if right else Decimal(0)
            rows.append([
                (left or right).name,
                left.quantity if left else None, right.quantity if right else None,
                left.price if left else None, right.price if right else None,
                amount_left, amount_right, amount_right - amount_left, state,
            ])
        table = pd.DataFrame(rows, columns=COLUMNS)
        correct = sum(row[-1] == CORRECT for row in rows)
        net = sum((row[-2] for row in rows), Decimal(0))
        excess = sum((row[-2] for row in rows if row[-2] > 0), Decimal(0))
        shortfall = sum((row[-2] for row in rows if row[-2] < 0), Decimal(0))
        requested_total = sum((row[5] for row in rows), Decimal(0))
        billed_total = sum((row[6] for row in rows), Decimal(0))
    return ComparisonResult(
        table, len(rows), correct, len(rows) - correct, net,
        excess, shortfall, requested_total, billed_total,
    )


def format_money(value: Decimal | None) -> str:
    """Redondear solo para presentación, sin convertir a float."""
    if value is None:
        return "—"
    with localcontext() as context:
        context.prec = max(28, len(value.as_tuple().digits), value.adjusted() + 4)
        rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(abs(rounded) if rounded == 0 else rounded, ".2f")
