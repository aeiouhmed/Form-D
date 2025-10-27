"""Utilities for transforming source spreadsheets into K1 import format."""

from __future__ import annotations

import io
import logging
import os
import re
from pathlib import Path
from typing import Iterable, List

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "templates" / "K1 Import Template.xls"
if not TEMPLATE_PATH.exists():
    alt_template = TEMPLATE_PATH.with_suffix(".xlsx")
    if alt_template.exists():
        TEMPLATE_PATH = alt_template

RANDOM_SEED = 42
ALWAYS_BLANK_COLUMN_NAMES = [
    "ExciseDutyMethod",
    "ExciseDutyRateExemptedPercentage",
    "ExciseDutyRateExemptedSpecific",
    "VehicleType",
    "VehicleModel",
    "Brand",
    "Engine",
    "Chassis",
    "CC",
    "Year",
]
HS_CANDIDATES = ["Hs Code", "HS Code", "HSCode", "HS-Code"]
NET_WEIGHT_CANDIDATES = [
    "Net Weight(Kg)",
    "Net Weight (Kg)",
    "Net Weight",
    "Weight (Kg)",
    "Weight",
]
AMOUNT_CANDIDATES = ["Amount(USD)", "Amount (USD)", "Amount USD", "Amount"]
PARTS_CANDIDATES = ["Parts Name", "Description", "Item Description"]
QUANTITY_CANDIDATES = ["Quantity", "Qty", "QTY"]
FORM_FLAG_CANDIDATES = ["Form Flag", "FormFlag", "Form_Flag", "Form flag"]


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    return str(value).lstrip("'").strip()


def normalize_for_match(value: object) -> str:
    return normalize_header(value).lower()


ALWAYS_BLANK_NORMALIZED = {
    normalize_for_match(column_name) for column_name in ALWAYS_BLANK_COLUMN_NAMES
}
ALWAYS_BLANK_COLLAPSED = {
    re.sub(r"[^a-z0-9]", "", value) for value in ALWAYS_BLANK_NORMALIZED
}


def convert_to_k1(
    uploaded_bytes: bytes,
    uom_mode: str = "random",
    template_path: str | Path | None = None,
) -> bytes:
    """Convert uploaded spreadsheet bytes into K1 import XLSX bytes."""
    if uom_mode not in {"random", "kgm"}:
        raise ValueError("uom_mode must be 'random' or 'kgm'.")

    np.random.seed(RANDOM_SEED)
    source_df = _load_source_dataframe(uploaded_bytes)
    source_df = _normalize_columns(source_df)

    form_flag_col = _first_matching_col(source_df.columns, FORM_FLAG_CANDIDATES)
    if form_flag_col is None:
        raise ValueError("Missing required column: 'Form Flag'")

    total_rows = len(source_df)
    mask = source_df[form_flag_col].apply(_normalize_flag).eq("form d")
    kept_rows = int(mask.sum())
    logging.info("Form Flag filter: kept %s of %s rows", kept_rows, total_rows)
    if kept_rows == 0:
        raise ValueError("No rows with 'Form D' in 'Form Flag'.")
    source_df = source_df.loc[mask].reset_index(drop=True)

    template_file = _resolve_template_path(template_path)
    logging.info("Using template at: %s", template_file)
    template_columns = _load_template(template_file)

    log_config(list(source_df.columns), template_columns)

    output = pd.DataFrame(index=source_df.index)

    hs_series = _get_series(source_df, HS_CANDIDATES, default="")
    hs_clean = hs_series.apply(_sanitize_hs)

    quantity_series = pd.to_numeric(
        _get_series(source_df, QUANTITY_CANDIDATES, default=0), errors="coerce"
    ).fillna(0)
    net_weight_series = pd.to_numeric(
        _get_series(source_df, NET_WEIGHT_CANDIDATES, default=0),
        errors="coerce",
    ).fillna(0)
    amount_series = pd.to_numeric(
        _get_series(source_df, AMOUNT_CANDIDATES, default=0),
        errors="coerce",
    ).fillna(0)
    parts_name_series = _get_series(
        source_df, PARTS_CANDIDATES, default=""
    ).fillna("")
    parts_name_series = parts_name_series.astype(str)

    if uom_mode == "random":
        uom_values = np.random.choice(["KGM", "UNT"], size=len(source_df))
    else:
        uom_values = np.full(len(source_df), "KGM")
    uom_val = pd.Series(uom_values, index=source_df.index, dtype="object")

    statistical_qty = net_weight_series.copy().astype(float)
    statistical_qty[uom_val == "UNT"] = quantity_series[uom_val == "UNT"]
    declared_qty = statistical_qty.copy()

    output["Country of Origin"] = "ID"
    output["HSCode"] = hs_clean
    output["StatisticalUOM"] = uom_val
    output["DeclaredUOM"] = uom_val
    output["StatisticalQty"] = statistical_qty
    output["DeclaredQty"] = declared_qty
    output["ItemAmount"] = amount_series
    output["ItemDescription"] = parts_name_series
    output["ItemDescription2"] = quantity_series
    output["ItemDescription3"] = ""

    output["ImportDutyMethod"] = "Exemption"
    output["ImportDutyRateExemptedPercentage"] = 100000
    output["ImportDutyRateExemptedSpecific"] = ""

    output["SSTMethod"] = "Exemption"
    output["SSTRateExemptedPercentage"] = 100000
    output["SSTRateExemptedSpecific"] = ""

    output["ExciseDutyMethod"] = ""
    output["ExciseDutyRateExemptedPercentage"] = ""
    output["ExciseDutyRateExemptedSpecific"] = ""

    output["VehicleType"] = ""
    output["VehicleModel"] = ""
    output["Brand"] = ""
    output["Engine"] = ""
    output["Chassis"] = ""
    output["CC"] = ""
    output["Year"] = ""

    if "StatisticalUOM" not in output.columns:
        output["StatisticalUOM"] = pd.Series([""] * len(output), index=output.index, dtype="object")
    if "DeclaredUOM" not in output.columns:
        output["DeclaredUOM"] = output["StatisticalUOM"].copy()

    assert (output["StatisticalUOM"] == output["DeclaredUOM"]).all(), (
        "DeclaredUOM must match StatisticalUOM"
    )

    normalized_output_map = {
        normalize_for_match(column): output[column] for column in output.columns
    }
    collapsed_output_map = {
        re.sub(r"[^a-z0-9]", "", key): value for key, value in normalized_output_map.items()
    }

    method_occurrence = 0
    final_series: list[pd.Series] = []
    for template_column in template_columns:
        normalized_template_column = normalize_header(template_column)
        normalized_template_key = normalize_for_match(template_column)
        collapsed_template_key = re.sub(r"[^a-z0-9]", "", normalized_template_key)

        if normalize_for_match(normalized_template_column) == "method":
            method_occurrence += 1
            if method_occurrence in (1, 2):
                fill_value = "E"
            elif method_occurrence == 3:
                fill_value = ""
            else:
                fill_value = ""
            series = pd.Series(
                [fill_value] * len(output),
                index=output.index,
                dtype="object",
            )
        elif (
            normalized_template_key in ALWAYS_BLANK_NORMALIZED
            or collapsed_template_key in ALWAYS_BLANK_COLLAPSED
        ):
            series = pd.Series([""] * len(output), index=output.index, dtype="object")
        elif normalized_template_key in normalized_output_map:
            series = normalized_output_map[normalized_template_key]
        elif collapsed_template_key in collapsed_output_map:
            series = collapsed_output_map[collapsed_template_key]
        else:
            series = pd.Series([""] * len(output), index=output.index, dtype="object")

        final_series.append(series.rename(template_column))

    final_df = (
        pd.concat(final_series, axis=1) if final_series else pd.DataFrame(index=output.index)
    )

    _maybe_log_debug_samples(final_df)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, header=True)
    try:
        _apply_template_styles(buffer, template_file, len(final_df.columns))
    except Exception as exc:  # noqa: BLE001
        logging.warning("Failed to copy header styles from template: %s", exc)
    buffer.seek(0)
    return buffer.getvalue()


def _load_source_dataframe(uploaded_bytes: bytes) -> pd.DataFrame:
    bio = io.BytesIO(uploaded_bytes)
    bio.seek(0)
    head = bio.read(4)
    bio.seek(0)
    is_xlsx_like = head.startswith(b"PK")
    if is_xlsx_like:
        return pd.read_excel(bio, engine="openpyxl")
    return _df_from_xls_bytes(uploaded_bytes)


def _load_template(template_path: Path) -> List[str]:
    template_df = _load_template_dataframe(template_path)
    return template_df.columns.tolist()


def _load_template_dataframe(template_path: Path) -> pd.DataFrame:
    ext = template_path.suffix.lower()
    if ext == ".xlsx":
        return pd.read_excel(
            template_path,
            engine="openpyxl",
            mangle_dupe_cols=False,
        )
    with open(template_path, "rb") as file:
        return _df_from_xls_bytes(file.read())


def _resolve_template_path(custom_path: str | Path | None) -> Path:
    if custom_path is None:
        path = TEMPLATE_PATH
    else:
        path = Path(custom_path)
        if not path.is_absolute():
            path = path.resolve()
    if not path.exists():
        raise ValueError(f"Template not found at {path.resolve()}")
    return path.resolve()


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    normalized = [normalize_for_match(col) for col in df.columns]
    df.columns = normalized
    return df


def _get_series(
    df: pd.DataFrame,
    candidates: Iterable[str],
    default: object,
) -> pd.Series:
    column_name = _first_matching_col(df.columns, candidates)
    if column_name is None:
        return pd.Series([default] * len(df), index=df.index)
    return df[column_name]


def log_config(source_columns: list[str], template_columns: list[str]) -> None:
    """Log column configuration for debugging purposes."""
    logging.info("Source columns (normalized): %s", source_columns)
    logging.info("Template columns (original): %s", template_columns)


def _first_matching_col(columns: Iterable[str], candidates: Iterable[str]) -> str | None:
    columns_list = list(columns)
    normalized_columns = [normalize_for_match(col) for col in columns_list]
    candidate_keys = [normalize_for_match(candidate) for candidate in candidates]
    for candidate_key in candidate_keys:
        for idx, column_key in enumerate(normalized_columns):
            if column_key == candidate_key:
                return columns_list[idx]
    for candidate_key in candidate_keys:
        if not candidate_key:
            continue
        for idx, column_key in enumerate(normalized_columns):
            if candidate_key in column_key:
                return columns_list[idx]
    collapsed_columns = [re.sub(r"[^a-z0-9]", "", column_key) for column_key in normalized_columns]
    collapsed_candidates = [re.sub(r"[^a-z0-9]", "", candidate_key) for candidate_key in candidate_keys]
    for candidate_key in collapsed_candidates:
        if not candidate_key:
            continue
        for idx, column_key in enumerate(collapsed_columns):
            if column_key == candidate_key or candidate_key in column_key:
                return columns_list[idx]
    return None


def _normalize_flag(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).lower()
    text = re.sub(r"[-_]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _sanitize_hs(value: object) -> str:
    if pd.isna(value):
        return ""
    digits_only = "".join(ch for ch in str(value) if ch.isdigit())
    # TODO: Clamp to required length if business rules change; current spec always appends "00".
    return f"{digits_only}00"


def _maybe_log_debug_samples(df: pd.DataFrame) -> None:
    if os.environ.get("DEBUG_MAPPINGS") == "1":
        records = df.head(3).to_dict(orient="records")
        logging.info("Sample mapped rows: %s", records)


def _apply_template_styles(buffer: io.BytesIO, template_path: Path, column_count: int) -> None:
    buffer.seek(0)
    wb_out = load_workbook(buffer)
    ws_out = wb_out.active

    wb_tpl = load_workbook(template_path, read_only=False, data_only=False)
    ws_tpl = wb_tpl.active

    if ws_tpl.max_column < column_count:
        logging.warning(
            "Template has fewer header columns (%s) than output (%s); skipping header styling.",
            ws_tpl.max_column,
            column_count,
        )
        return

    for col_idx in range(1, column_count + 1):
        src_cell = ws_tpl.cell(row=1, column=col_idx)
        dst_cell = ws_out.cell(row=1, column=col_idx)
        dst_cell.font = src_cell.font
        dst_cell.fill = src_cell.fill
        dst_cell.alignment = src_cell.alignment
        dst_cell.border = src_cell.border

        column_letter = get_column_letter(col_idx)
        tpl_dim = ws_tpl.column_dimensions.get(column_letter)
        if tpl_dim is not None and tpl_dim.width is not None:
            ws_out.column_dimensions[column_letter].width = tpl_dim.width

    ws_out.freeze_panes = "A2"

    wb_out.save(buffer)
    logging.info("Header styles copied from template.")


def _df_from_xls_bytes(xls_bytes: bytes) -> pd.DataFrame:
    import xlrd

    book = xlrd.open_workbook(file_contents=xls_bytes)
    sheet = book.sheet_by_index(0)
    headers = [str(sheet.cell_value(0, c)).lstrip("'").strip() for c in range(sheet.ncols)]
    rows: list[list[object]] = []
    for r in range(1, sheet.nrows):
        row_values = [sheet.cell_value(r, c) for c in range(sheet.ncols)]
        rows.append(row_values)
    if rows:
        return pd.DataFrame(rows, columns=headers)
    return pd.DataFrame(columns=headers)
