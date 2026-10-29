"""Lectura y validación estructural de documentos Excel, sin interfaz."""

from pathlib import Path
from typing import BinaryIO

import pandas as pd


REQUIRED_COLUMNS = ("Producto", "Cantidad", "Precio unitario")


class ExcelValidationError(ValueError):
    """Error de archivo o estructura que puede mostrarse al usuario."""


def validate_columns(data: pd.DataFrame) -> None:
    """Exigir los tres encabezados exactos; permitir columnas adicionales."""
    missing = [column for column in REQUIRED_COLUMNS if column not in data.columns]
    if missing:
        raise ExcelValidationError(
            "Faltan columnas obligatorias: " + ", ".join(missing) + ". "
            "Usá estos encabezados en la primera fila: "
            + ", ".join(REQUIRED_COLUMNS) + "."
        )


def read_excel(source: str | Path | BinaryIO) -> pd.DataFrame:
    """Leer la primera hoja de un .xlsx y validar sus encabezados.

    Acepta rutas o archivos en memoria (incluido UploadedFile de Streamlit).
    Todavía no valida valores ni realiza cálculos monetarios.
    """
    name = str(source) if isinstance(source, (str, Path)) else getattr(source, "name", "")
    if name and Path(name).suffix.lower() != ".xlsx":
        raise ExcelValidationError("El archivo debe tener extensión .xlsx.")

    try:
        if hasattr(source, "seek"):
            source.seek(0)
        # Evitar que Pandas convierta textos decimales a float o nombres como
        # "NA" en valores vacíos. Excel ya limita la precisión de celdas numéricas.
        data = pd.read_excel(
            source, engine="openpyxl", sheet_name=0, dtype=object, keep_default_na=False,
        )
    except Exception as exc:
        # Los errores del parser varían según el daño del archivo; se traducen
        # solo en esta frontera de lectura y se conserva la causa original.
        raise ExcelValidationError(
            "No se pudo procesar el archivo. Verificá que sea un Excel .xlsx "
            "válido, que no esté dañado ni protegido con contraseña."
        ) from exc

    validate_columns(data)
    return data
