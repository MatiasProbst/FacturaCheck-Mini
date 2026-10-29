"""Interfaz Streamlit: carga, resumen, filtros de visualización y descarga."""

from pathlib import Path

import pandas as pd
import streamlit as st

from excel_reader import ExcelValidationError, read_excel
from report import create_report
from comparison import (
    BOTH, CORRECT, INVOICE_ONLY, MONEY_COLUMNS, ORDER_ONLY, PRICE, QUANTITY,
    DocumentValidationError, compare_documents, format_money, nonnegative_decimal, validate_document,
)

EXAMPLES = Path(__file__).resolve().parent / "examples"
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
STATE_LABELS = {
    CORRECT: "✅ Correcto", PRICE: "🟠 Diferencia de precio",
    QUANTITY: "🟡 Diferencia de cantidad", BOTH: "🔴 Diferencia de precio y cantidad",
    ORDER_ONLY: "🔵 Solo en orden de compra", INVOICE_ONLY: "🟣 Solo en factura",
}
ALL_STATES = "Todos los estados"


def format_money_ar(value) -> str:
    """Presentación es-AR sin float ni cambios en la configuración regional del equipo."""
    text = format_money(value)
    if text == "—":
        return text
    integer, decimals = text.split(".")
    sign = "-" if integer.startswith("-") else ""
    digits = integer.lstrip("-")
    groups = []
    while digits:
        groups.insert(0, digits[-3:])
        digits = digits[:-3]
    return sign + ".".join(groups) + "," + decimals


def preview_price(value):
    """Formatear precios válidos sin ocultar el contenido de celdas inválidas."""
    try:
        return format_money_ar(nonnegative_decimal(value))
    except ValueError:
        return str(value)


def filter_results(table: pd.DataFrame, state: str, differences_only: bool) -> pd.DataFrame:
    """Filtrar una copia para presentación; no cambia resultados ni agregados."""
    visible = table
    if state != ALL_STATES:
        visible = visible.loc[visible["Estado"] == state]
    if differences_only:
        visible = visible.loc[visible["Estado"] != CORRECT]
    return visible.copy()


def show_document(label: str, key: str) -> pd.DataFrame | None:
    st.subheader(label)
    uploaded = st.file_uploader(
        f"Cargar {label.lower()}", type=["xlsx"], key=key,
        help="Primera hoja del Excel, con los encabezados en la primera fila.",
    )
    filename = "orden_compra.xlsx" if key == "purchase_order" else "factura_proveedor.xlsx"
    st.download_button(
        f"Descargar ejemplo de {label.lower()}",
        data=(EXAMPLES / filename).read_bytes(), file_name=filename, mime=EXCEL_MIME,
        key=f"example_{key}", icon=":material/download:", width="stretch", on_click="ignore",
    )
    if uploaded is None:
        st.info("Seleccioná un archivo .xlsx o probá con el ejemplo.")
        return None
    try:
        data = read_excel(uploaded)
    except ExcelValidationError as exc:
        st.error(f"{label}: {exc}")
        return None

    errors = []
    try:
        validate_document(data, label)
    except DocumentValidationError as exc:
        errors = exc.errors
    if errors:
        for error in errors:
            st.error(error)
    else:
        st.success(f"{label}: archivo válido · {len(data)} filas.")
        if data.empty:
            st.warning("El archivo tiene encabezados, pero no contiene filas de datos.")
    with st.expander(f"Ver datos de {label.lower()}"):
        preview = data.copy()
        preview["Precio unitario"] = preview["Precio unitario"].map(preview_price)
        st.dataframe(preview, hide_index=True, width="stretch")
    return None if errors else data


def show_comparison(order: pd.DataFrame, invoice: pd.DataFrame) -> None:
    result = compare_documents(order, invoice)
    st.divider()
    st.header("2. Resumen de resultados")
    total, correct, different = st.columns(3)
    total.metric("Productos analizados", result.total, border=True)
    correct.metric("Productos correctos", result.correct, border=True)
    different.metric("Productos con diferencias", result.different, border=True)
    net, excess, shortfall = st.columns(3)
    net.metric("Diferencia económica neta", format_money_ar(result.net_difference), border=True,
               help="Importe facturado menos solicitado. Compensa excesos y faltantes.")
    excess.metric("Importe facturado en exceso (+)", "+" + format_money_ar(result.excess), border=True,
                  help="Suma de diferencias positivas.")
    shortfall.metric("Importe facturado de menos (−)", format_money_ar(result.shortfall), border=True,
                     help="Suma de diferencias negativas, conservando el signo menos.")
    st.caption(
        "Importes en la moneda de los archivos · Dos decimales · "
        "Positivo: mayor importe facturado. Negativo: menor importe facturado."
    )
    with st.container(border=True):
        st.subheader("Informe de la comparación")
        st.write("Resumen, comparación y observaciones en un único Excel.")
        st.download_button(
            "Descargar informe Excel", data=create_report(result),
            file_name="FacturaCheck_Mini_informe.xlsx", mime=EXCEL_MIME,
            key="download_report", type="primary", icon=":material/download:",
            width="stretch", on_click="ignore",
        )
        st.caption("Incluye todos los productos. Los filtros del detalle no modifican el informe ni las tarjetas.")

    st.header("3. Detalle de la comparación")
    state = st.selectbox("Filtrar por estado", [ALL_STATES, *STATE_LABELS], key="state_filter")
    differences_only = st.toggle("Mostrar solamente productos con diferencias", key="differences_only")
    visible = filter_results(result.table, state, differences_only)
    st.caption(f"Mostrando {len(visible)} de {result.total} productos.")
    display = visible.copy()
    for column in MONEY_COLUMNS:
        display[column] = display[column].map(format_money_ar)
    for column in ("Cantidad solicitada", "Cantidad facturada"):
        display[column] = display[column].map(lambda value: "—" if value is None else str(value))
    display["Estado"] = display["Estado"].map(STATE_LABELS)

    def highlight_difference(row):
        value = visible.loc[row.name, "Diferencia económica"]
        color = "#9C3500" if value > 0 else "#1F4E78" if value < 0 else "#374151"
        return [f"color: {color}; font-weight: bold" if column == "Diferencia económica" else ""
                for column in row.index]

    if visible.empty:
        st.info("No hay productos que coincidan con estos filtros. Elegí otro estado o desactivá la opción de diferencias.")
    st.dataframe(
        display.style.apply(highlight_difference, axis=1), hide_index=True, width="stretch",
        column_config={
            "Producto": st.column_config.TextColumn(width="medium"),
            "Estado": st.column_config.TextColumn(width="large"),
            "Diferencia económica": st.column_config.TextColumn(
                help="Facturado menos solicitado. Naranja: positivo; azul: negativo.", width="medium",
            ),
        },
    )
    st.caption("En pantallas pequeñas, deslizá la tabla horizontalmente para ver todas las columnas.")
    with st.expander("Cómo interpretar los resultados"):
        st.write(" · ".join(STATE_LABELS.values()))
        st.write(
            "Los productos ausentes tienen importe cero y cantidad/precio sin dato (—). "
            "Los cálculos no se redondean hasta su presentación: una diferencia menor a un "
            "centavo puede mostrarse como 0,00 y seguir siendo una diferencia."
        )


def main() -> None:
    st.set_page_config(page_title="FacturaCheck Mini", page_icon=":material/fact_check:", layout="wide")
    st.title("FacturaCheck Mini")
    st.write("Compará órdenes de compra y facturas, detectá inconsistencias y generá informes automáticamente")
    st.caption("1. Cargá los dos Excel · 2. Revisá las diferencias · 3. Descargá el informe")
    st.header("1. Cargá tus archivos")
    st.caption(
        "Formato .xlsx · Primera hoja · Encabezados: Producto, Cantidad y Precio unitario. "
        "Podés empezar con los ejemplos ficticios."
    )
    order, invoice = st.columns(2)
    with order, st.container(border=True):
        order_data = show_document("Orden de compra", "purchase_order")
    with invoice, st.container(border=True):
        invoice_data = show_document("Factura del proveedor", "supplier_invoice")
    if order_data is not None and invoice_data is not None:
        show_comparison(order_data, invoice_data)
    else:
        st.caption("Para comparar, cargá ambos documentos y corregí todos los errores de validación.")
    st.divider()
    with st.expander("Acerca del proyecto"):
        st.markdown(
            "**El problema.** Revisar manualmente una orden de compra y una factura puede "
            "ocultar diferencias de precios, cantidades o productos.\n\n"
            "**Cómo funciona.** Lee dos Excel estructurados, valida sus datos y relaciona "
            "los productos ignorando mayúsculas y espacios adicionales. Muestra los resultados "
            "y genera un informe mediante reglas de Python, sin inteligencia artificial.\n\n"
            "**Tecnologías.** Python, Streamlit, Pandas, openpyxl, Decimal, pytest y Git. "
            "Proyecto educativo y de portfolio.\n\n"
            "**Validaciones.** Columnas obligatorias, nombres de productos no vacíos, duplicados "
            "normalizados y cantidades/precios numéricos, finitos y no negativos. "
            "Si hay errores, la comparación queda bloqueada.\n\n"
            "**Limitaciones.** Compara archivos Excel estructurados; no interpreta automáticamente "
            "facturas PDF. Lee solo la primera hoja, supone la misma moneda y no calcula "
            "impuestos, descuentos ni conversiones. Los archivos se procesan en memoria y "
            "los originales no se modifican."
        )


if __name__ == "__main__":
    main()
