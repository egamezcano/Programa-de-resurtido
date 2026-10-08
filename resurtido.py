#!/usr/bin/env python3
"""
resurtido.py — Generador de propuestas de resurtido para proveedores.

Lee las cuatro hojas de inventario_ferreteria_garza.xlsx, valida registros,
calcula necesidades, agrupa por proveedor, retiene pedidos por mínimo de compra,
y genera:
  - pedidos_proveedores.xlsx
  - reporte_excepciones.xlsx
  - borradores_correos/<proveedor>.eml

IMPORTANTE:
- No envía correos. Los archivos .eml son borradores para revisión humana.
- El libro de Excel debe contener una hoja de proveedores con correo y pedido mínimo,
  además de hojas con productos/inventario. Los nombres de columnas se normalizan
  y se aceptan varios alias en español.
- Si tu libro usa nombres diferentes, ajusta ALIASES o la configuración CLI.
"""

from __future__ import annotations

import argparse
import logging
import math
import re
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from email.message import EmailMessage
from pathlib import Path
from typing import Any

import pandas as pd


INPUT_DEFAULT = "inventario_ferreteria_garza.xlsx"
ORDERS_DEFAULT = "pedidos_proveedores.xlsx"
EXCEPTIONS_DEFAULT = "reporte_excepciones.xlsx"
DRAFTS_DEFAULT = "borradores_correos"

# Claves normalizadas: sin espacios, acentos ni signos de puntuación.
ALIASES: dict[str, tuple[str, ...]] = {
    "codigo": ("codigo", "sku", "clave", "codigo producto", "clave producto", "id producto"),
    "descripcion": ("descripcion", "producto", "nombre", "nombre producto", "descripcion producto"),
    "existencia": ("existencia", "existencia actual", "stock", "inventario", "cantidad disponible"),
    "minimo": ("minimo", "stock minimo", "existencia minima", "minimo existencia", "punto reorden"),
    "maximo": ("maximo", "stock maximo", "existencia maxima", "maximo existencia"),
    "activo": ("activo", "estatus", "estado", "producto activo", "status"),
    "empaque_multiplo": ("empaque multiplo", "empaque", "multiplo empaque", "multiplo de empaque", "cantidad empaque"),
    "proveedor": ("proveedor", "codigo proveedor", "id proveedor", "clave proveedor", "proveedor id"),
    "costo_unitario_mxn": ("costo unitario mxn", "costo unitario", "costo", "precio costo", "precio unitario", "costo mxn"),
    "correo": ("correo", "email", "e mail", "correo electronico", "email proveedor"),
    "pedido_minimo_mxn": ("pedido minimo mxn", "pedido minimo", "compra minima", "minimo compra", "monto minimo", "pedido minimo pesos"),
    "contacto": ("contacto", "nombre contacto", "responsable", "nombre proveedor"),
    "nombre_proveedor": ("nombre proveedor", "razon social", "empresa", "proveedor nombre", "nombre"),
}

REQUIRED_PRODUCT = (
    "codigo", "descripcion", "existencia", "minimo", "maximo",
    "activo", "empaque_multiplo", "proveedor", "costo_unitario_mxn",
)
REQUIRED_SUPPLIER = ("proveedor", "correo", "pedido_minimo_mxn")


@dataclass
class Supplier:
    supplier_id: str
    email: str
    minimum_order_mxn: Decimal
    name: str
    contact: str = ""


def normalize_header(value: Any) -> str:
    """Normaliza encabezados para comparar nombres con tolerancia a acentos."""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_code(value: Any) -> str:
    """Convierte claves a texto sin perder ceros iniciales cuando Excel los conserva."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()
    return str(value).strip()


def parse_decimal(value: Any) -> Decimal:
    """Convierte valores numéricos comunes; rechaza valores faltantes o inválidos."""
    if pd.isna(value) or str(value).strip() == "":
        raise ValueError("valor numérico faltante")
    if isinstance(value, Decimal):
        result = value
    else:
        raw = str(value).strip().replace("$", "").replace("MXN", "").replace(" ", "")
        # Admite 1,234.56 y 1234,56. Si hay coma y punto, asume coma de miles.
        if "," in raw and "." in raw:
            raw = raw.replace(",", "")
        elif "," in raw and "." not in raw:
            raw = raw.replace(",", ".")
        result = Decimal(raw)
    if not result.is_finite():
        raise ValueError("valor numérico no finito")
    return result


def parse_active(value: Any) -> bool:
    """Interpreta valores comunes de activo/inactivo."""
    if pd.isna(value):
        raise ValueError("estado activo faltante")
    normalized = normalize_header(value)
    if normalized in {"1", "si", "s", "true", "verdadero", "activo", "active", "yes", "y"}:
        return True
    if normalized in {"0", "no", "n", "false", "falso", "inactivo", "inactive", "disabled"}:
        return False
    raise ValueError(f"estado activo no reconocido: {value!r}")


def canonicalize_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Renombra columnas reconocidas a nombres canónicos."""
    lookup: dict[str, str] = {}
    for canonical, names in ALIASES.items():
        for name in names:
            lookup[normalize_header(name)] = canonical

    renamed: dict[Any, str] = {}
    seen: set[str] = set()
    for col in frame.columns:
        candidate = lookup.get(normalize_header(col))
        if candidate and candidate not in seen:
            renamed[col] = candidate
            seen.add(candidate)
        else:
            renamed[col] = str(col).strip()
    return frame.rename(columns=renamed)


def log_exception(
    exceptions: list[dict[str, Any]],
    sheet: str,
    row_number: Any,
    code: Any,
    reason: str,
    action: str,
    detail: str = "",
) -> None:
    exceptions.append({
        "Hoja": sheet,
        "Fila_Excel": row_number,
        "Codigo": "" if pd.isna(code) else str(code),
        "Motivo": reason,
        "Detalle": detail,
        "Accion_tomada": action,
    })


def load_workbook(path: Path) -> dict[str, pd.DataFrame]:
    """Carga todas las hojas; se espera que el archivo tenga cuatro hojas."""
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {path}")
    sheets = pd.read_excel(path, sheet_name=None, dtype=object, engine="openpyxl")
    if not sheets:
        raise ValueError("El libro no contiene hojas.")
    if len(sheets) != 4:
        logging.warning("Se esperaban cuatro hojas; se encontraron %d.", len(sheets))
    return {name: canonicalize_columns(frame) for name, frame in sheets.items()}


def find_supplier_sheet(sheets: dict[str, pd.DataFrame]) -> str | None:
    """Busca la hoja que contiene las columnas necesarias para proveedores."""
    for name, frame in sheets.items():
        if all(col in frame.columns for col in REQUIRED_SUPPLIER):
            return name
    # Fallback por nombre de hoja para dar un error más claro.
    for name in sheets:
        if "proveedor" in normalize_header(name):
            return name
    return None


def load_suppliers(
    sheet_name: str,
    frame: pd.DataFrame,
    exceptions: list[dict[str, Any]],
) -> dict[str, Supplier]:
    """Valida y construye catálogo de proveedores indexado por clave."""
    missing = [c for c in REQUIRED_SUPPLIER if c not in frame.columns]
    if missing:
        raise ValueError(
            f"La hoja '{sheet_name}' no contiene columnas de proveedor requeridas: {missing}. "
            f"Columnas detectadas: {list(frame.columns)}"
        )

    suppliers: dict[str, Supplier] = {}
    for idx, row in frame.iterrows():
        excel_row = idx + 2
        supplier_id = normalize_code(row.get("proveedor"))
        if not supplier_id:
            log_exception(exceptions, sheet_name, excel_row, "", "Clave de proveedor faltante",
                          "Registro de proveedor omitido")
            continue
        try:
            email = str(row.get("correo", "")).strip()
            if not email or email.lower() == "nan" or not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
                raise ValueError("correo faltante o con formato inválido")
            minimum = parse_decimal(row.get("pedido_minimo_mxn"))
            if minimum < 0:
                raise ValueError("pedido mínimo no puede ser negativo")
            name_value = row.get("nombre_proveedor", supplier_id)
            name = supplier_id if pd.isna(name_value) or not str(name_value).strip() else str(name_value).strip()
            contact_value = row.get("contacto", "")
            contact = "" if pd.isna(contact_value) else str(contact_value).strip()
            if supplier_id in suppliers:
                log_exception(exceptions, sheet_name, excel_row, supplier_id, "Clave de proveedor duplicada",
                              "Se conservó el primer registro", "La clave aparece más de una vez")
                continue
            suppliers[supplier_id] = Supplier(supplier_id, email, minimum, name, contact)
        except (ValueError, InvalidOperation, TypeError) as exc:
            log_exception(exceptions, sheet_name, excel_row, supplier_id, "Datos de proveedor inválidos",
                          "Proveedor omitido", str(exc))
    return suppliers


def load_products(
    sheets: dict[str, pd.DataFrame],
    supplier_sheet: str,
    suppliers: dict[str, Supplier],
    exceptions: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Combina filas de las hojas restantes que contengan columnas de producto."""
    products: list[dict[str, Any]] = []
    seen_codes: dict[str, tuple[str, int]] = {}

    for sheet_name, frame in sheets.items():
        if sheet_name == supplier_sheet:
            continue
        # Ignorar hojas que no sean de productos y reportarlas como no procesables.
        missing = [col for col in REQUIRED_PRODUCT if col not in frame.columns]
        if missing:
            log_exception(
                exceptions, sheet_name, "", "",
                "Hoja sin estructura de inventario reconocida",
                "Hoja omitida",
                f"Faltan columnas: {missing}; columnas detectadas: {list(frame.columns)}",
            )
            continue

        for idx, row in frame.iterrows():
            excel_row = idx + 2
            code = normalize_code(row.get("codigo"))
            try:
                if not code:
                    raise ValueError("código faltante (NaN o vacío)")
                if code in seen_codes:
                    prev_sheet, prev_row = seen_codes[code]
                    log_exception(
                        exceptions, sheet_name, excel_row, code, "Código de producto duplicado",
                        "Registro omitido; se conservó la primera aparición",
                        f"Primera aparición: hoja {prev_sheet}, fila {prev_row}",
                    )
                    continue

                description = str(row.get("descripcion", "")).strip()
                if not description or description.lower() == "nan":
                    raise ValueError("descripción faltante")
                active = parse_active(row.get("activo"))
                existence = parse_decimal(row.get("existencia"))
                minimum = parse_decimal(row.get("minimo"))
                maximum = parse_decimal(row.get("maximo"))
                pack = parse_decimal(row.get("empaque_multiplo"))
                cost = parse_decimal(row.get("costo_unitario_mxn"))
                supplier_id = normalize_code(row.get("proveedor"))

                if existence < 0 or minimum < 0 or maximum < 0:
                    raise ValueError("existencia, mínimo y máximo deben ser no negativos")
                if minimum > maximum:
                    raise ValueError("Mínimo > Máximo")
                if pack <= 0 or pack != pack.to_integral_value():
                    raise ValueError("Empaque_multiplo debe ser un entero positivo")
                if cost < 0:
                    raise ValueError("Costo_unitario_MXN no puede ser negativo")
                if not supplier_id or supplier_id not in suppliers:
                    raise ValueError(f"clave de proveedor huérfana o inexistente: {supplier_id!r}")

                seen_codes[code] = (sheet_name, excel_row)
                if not active:
                    log_exception(exceptions, sheet_name, excel_row, code, "Producto inactivo",
                                  "No se generó pedido")
                    continue
                if existence < minimum:
                    deficit = maximum - existence
                    quantity = int(math.ceil(float(deficit / pack)) * int(pack))
                    if quantity <= 0:
                        log_exception(exceptions, sheet_name, excel_row, code,
                                      "Cantidad sugerida no positiva", "Producto omitido")
                        continue
                    products.append({
                        "Codigo": code,
                        "Descripcion": description,
                        "Proveedor_ID": supplier_id,
                        "Proveedor": suppliers[supplier_id].name,
                        "Correo": suppliers[supplier_id].email,
                        "Cantidad_sugerida": quantity,
                        "Costo_unitario_MXN": cost.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
                        "Total_MXN": (Decimal(quantity) * cost).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
                        "Hoja_origen": sheet_name,
                        "Fila_origen": excel_row,
                    })
            except (ValueError, InvalidOperation, TypeError, OverflowError) as exc:
                log_exception(exceptions, sheet_name, excel_row, code, "Producto no procesado",
                              "Registro omitido", str(exc))
            except Exception as exc:  # Protección por fila ante formatos inesperados.
                log_exception(exceptions, sheet_name, excel_row, code, "Error inesperado",
                              "Registro omitido", f"{type(exc).__name__}: {exc}")
                logging.exception("Error inesperado en hoja %s, fila %s", sheet_name, excel_row)
    return products


def build_orders(products: list[dict[str, Any]], suppliers: dict[str, Supplier]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Añade estado de pedido mínimo y genera resumen por proveedor."""
    columns = [
        "Proveedor_ID", "Proveedor", "Correo", "Codigo", "Descripcion",
        "Cantidad_sugerida", "Costo_unitario_MXN", "Total_MXN",
        "Pedido_minimo_MXN", "Total_proveedor_MXN", "Estado_pedido",
        "Hoja_origen", "Fila_origen",
    ]
    if not products:
        return pd.DataFrame(columns=columns), pd.DataFrame(
            columns=["Proveedor_ID", "Proveedor", "Correo", "Total_proveedor_MXN", "Pedido_minimo_MXN", "Estado_pedido"]
        )

    by_supplier: dict[str, Decimal] = defaultdict(Decimal)
    for item in products:
        by_supplier[item["Proveedor_ID"]] += item["Total_MXN"]

    rows: list[dict[str, Any]] = []
    for item in products:
        supplier = suppliers[item["Proveedor_ID"]]
        total = by_supplier[item["Proveedor_ID"]].quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        state = "RETENIDO_POR_MINIMO" if total < supplier.minimum_order_mxn else "LISTO_PARA_REVISION"
        rows.append({
            **item,
            "Pedido_minimo_MXN": float(supplier.minimum_order_mxn),
            "Total_proveedor_MXN": float(total),
            "Estado_pedido": state,
        })

    summary_rows: list[dict[str, Any]] = []
    for supplier_id, total in by_supplier.items():
        supplier = suppliers[supplier_id]
        total = total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        state = "RETENIDO_POR_MINIMO" if total < supplier.minimum_order_mxn else "LISTO_PARA_REVISION"
        summary_rows.append({
            "Proveedor_ID": supplier_id,
            "Proveedor": supplier.name,
            "Correo": supplier.email,
            "Total_proveedor_MXN": float(total),
            "Pedido_minimo_MXN": float(supplier.minimum_order_mxn),
            "Estado_pedido": state,
        })
    return pd.DataFrame(rows, columns=columns), pd.DataFrame(summary_rows)


def safe_filename(value: str) -> str:
    """Crea un nombre de archivo seguro para el sistema operativo."""
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("._")
    return (value or "proveedor")[:80]


def format_mxn(value: Any) -> str:
    amount = Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"${amount:,.2f} MXN"


def create_email_drafts(orders: pd.DataFrame, suppliers: dict[str, Supplier], output_dir: Path) -> None:
    """Genera borradores .eml para revisión; no envía mensajes."""
    output_dir.mkdir(parents=True, exist_ok=True)
    if orders.empty:
        return

    for supplier_id, group in orders.groupby("Proveedor_ID", sort=True):
        supplier = suppliers[str(supplier_id)]
        total = Decimal(str(group["Total_MXN"].sum())).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        state = str(group["Estado_pedido"].iloc[0])
        lines = [
            f"Estimado/a {supplier.contact or supplier.name}:",
            "",
            "Compartimos la propuesta de resurtido para su revisión.",
            "Este archivo es un BORRADOR y requiere validación y autorización interna antes de enviarse.",
            "",
            f"Proveedor: {supplier.name}",
            f"Estado interno: {state}",
            "",
            f"{'Código':<18} {'Descripción':<34} {'Cantidad':>10} {'Costo unitario':>18} {'Importe':>18}",
            "-" * 104,
        ]
        for _, row in group.iterrows():
            description = str(row["Descripcion"]).replace("\n", " ")[:34]
            lines.append(
                f"{str(row['Codigo'])[:18]:<18} {description:<34} "
                f"{int(row['Cantidad_sugerida']):>10} "
                f"{format_mxn(row['Costo_unitario_MXN']):>18} "
                f"{format_mxn(row['Total_MXN']):>18}"
            )
        lines += [
            "-" * 104,
            f"{'TOTAL PROPUESTO:':>82} {format_mxn(total):>21}",
            "",
            "Agradecemos confirmar disponibilidad, precio vigente y fecha estimada de entrega.",
            "",
            "Saludos cordiales,",
            "Departamento de Compras",
            "",
            "Nota: Este documento es una propuesta preliminar y no constituye una orden de compra autorizada.",
        ]
        msg = EmailMessage()
        msg["To"] = supplier.email
        msg["Subject"] = f"BORRADOR - Propuesta de resurtido - {supplier.name}"
        msg.set_content("\n".join(lines))
        path = output_dir / f"{safe_filename(supplier_id)}_{safe_filename(supplier.name)}.eml"
        try:
            with path.open("wb") as handle:
                handle.write(bytes(msg))
        except Exception as exc:
            logging.exception("No se pudo crear borrador para proveedor %s", supplier_id)
            raise OSError(f"No se pudo escribir el borrador {path}: {exc}") from exc


def write_outputs(
    orders: pd.DataFrame,
    summary: pd.DataFrame,
    exceptions: list[dict[str, Any]],
    orders_path: Path,
    exceptions_path: Path,
) -> None:
    """Escribe los reportes en Excel; genera hojas vacías con encabezados si no hay datos."""
    orders_path.parent.mkdir(parents=True, exist_ok=True)
    exceptions_path.parent.mkdir(parents=True, exist_ok=True)
    exceptions_df = pd.DataFrame(
        exceptions,
        columns=["Hoja", "Fila_Excel", "Codigo", "Motivo", "Detalle", "Accion_tomada"],
    )
    try:
        with pd.ExcelWriter(orders_path, engine="openpyxl") as writer:
            orders.to_excel(writer, sheet_name="Detalle_pedidos", index=False)
            summary.to_excel(writer, sheet_name="Resumen_proveedor", index=False)
    except Exception as exc:
        raise OSError(f"No se pudo generar {orders_path}: {exc}") from exc
    try:
        exceptions_df.to_excel(exceptions_path, index=False, engine="openpyxl")
    except Exception as exc:
        raise OSError(f"No se pudo generar {exceptions_path}: {exc}") from exc


def configure_logging(log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Genera propuestas de resurtido y borradores por proveedor.")
    parser.add_argument("--input", default=INPUT_DEFAULT, help="Ruta del archivo Excel de inventario.")
    parser.add_argument("--orders", default=ORDERS_DEFAULT, help="Ruta del Excel de pedidos consolidados.")
    parser.add_argument("--exceptions", default=EXCEPTIONS_DEFAULT, help="Ruta del reporte de excepciones.")
    parser.add_argument("--drafts", default=DRAFTS_DEFAULT, help="Carpeta de borradores .eml.")
    parser.add_argument("--log", default="logs/resurtido.log", help="Ruta del archivo de log.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = Path(args.input).expanduser().resolve()
    orders_path = Path(args.orders).expanduser().resolve()
    exceptions_path = Path(args.exceptions).expanduser().resolve()
    drafts_path = Path(args.drafts).expanduser().resolve()
    log_path = Path(args.log).expanduser().resolve()
    configure_logging(log_path)

    exceptions: list[dict[str, Any]] = []
    try:
        logging.info("Iniciando proceso de resurtido.")
        sheets = load_workbook(input_path)
        logging.info("Hojas detectadas: %s", ", ".join(sheets.keys()))

        supplier_sheet = find_supplier_sheet(sheets)
        if supplier_sheet is None:
            raise ValueError(
                "No se encontró una hoja de proveedores con columnas equivalentes a "
                "'Proveedor', 'Correo' y 'Pedido_minimo_MXN'."
            )
        logging.info("Hoja de proveedores detectada: %s", supplier_sheet)

        suppliers = load_suppliers(supplier_sheet, sheets[supplier_sheet], exceptions)
        if not suppliers:
            raise ValueError("No se encontró ningún proveedor válido; revise el catálogo de proveedores.")

        products = load_products(sheets, supplier_sheet, suppliers, exceptions)
        orders, summary = build_orders(products, suppliers)

        # Documentar explícitamente los pedidos retenidos en el log.
        if not summary.empty:
            for _, row in summary.iterrows():
                if row["Estado_pedido"] == "RETENIDO_POR_MINIMO":
                    logging.warning(
                        "Pedido retenido por mínimo: proveedor=%s, total=%s, mínimo=%s",
                        row["Proveedor_ID"],
                        format_mxn(row["Total_proveedor_MXN"]),
                        format_mxn(row["Pedido_minimo_MXN"]),
                    )

        write_outputs(orders, summary, exceptions, orders_path, exceptions_path)
        create_email_drafts(orders, suppliers, drafts_path)

        logging.info("Productos candidatos a pedido: %d", len(orders))
        logging.info("Proveedores con propuesta: %d", len(summary))
        logging.info("Excepciones registradas: %d", len(exceptions))
        logging.info("Reporte de pedidos: %s", orders_path)
        logging.info("Reporte de excepciones: %s", exceptions_path)
        logging.info("Borradores .eml: %s", drafts_path)
        logging.info("Proceso terminado. No se enviaron correos.")
        return 0
    except FileNotFoundError as exc:
        logging.error("%s", exc)
        return 2
    except (ValueError, PermissionError, OSError, ImportError) as exc:
        logging.exception("No se pudo completar el proceso: %s", exc)
        return 1
    except Exception as exc:
        logging.exception("Error fatal no controlado: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
