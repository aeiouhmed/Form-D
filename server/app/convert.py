"""Utilities for transforming source spreadsheets into K1 import format."""

from __future__ import annotations

from io import BytesIO
import logging
import os
import re
from pathlib import Path
from typing import Iterable, List

import pandas as pd
from openpyxl import load_workbook

TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "templates" / "K1 Import Template.xls"
if not TEMPLATE_PATH.exists():
    alt_template = TEMPLATE_PATH.with_suffix(".xlsx")
    if alt_template.exists():
        TEMPLATE_PATH = alt_template

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
HS_CANDIDATES = ["Hs Code", "HS Code", "Hscode", "HSCode", "HS-Code"]
HS_MAPPING_HEADER_VARIANTS = [
    "Hs Code",
    "HS Code",
    "Hscode",
    "HSCode",
    "HsCode",
    "H S Code",
    "Hs code",
    "HS code",
]
UNIT_HEADER_VARIANTS = ["Unit", "Units", "UOM", "UNTnit"]
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
HS_MAPPING_PATH = Path(__file__).resolve().parent.parent / "templates" / "HSCODE.xlsx"
_CTRL_BAD = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def _normalize_header(value: object) -> str:
    return re.sub(r"[\s_\-]+", "", str(value or "").strip().lower())


def _digits_only(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _hs_out_code(value: object) -> str:
    digits = _digits_only(value)
    return f"{digits}00"


def _sanitize_cell(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, (int, float)):
        return value
    text = str(value)
    cleaned = _CTRL_BAD.sub(" ", text)
    if len(cleaned) > 32767:
        cleaned = cleaned[:32767]
    return cleaned


def _load_template_columns_xlsx(path: Path) -> list[str]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    headers: list[str] = []
    first_row = next(ws.iter_rows(min_row=1, max_row=1), ())
    for cell in first_row:
        value = cell.value
        if isinstance(value, str):
            value = value.strip().lstrip("'")
        headers.append("" if value is None else str(value))
    wb.close()
    return headers


def _load_template_columns_xls(path: Path) -> list[str]:
    import xlrd

    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_index(0)
    headers: list[str] = []
    for col_idx in range(sheet.ncols):
        value = sheet.cell_value(0, col_idx)
        if isinstance(value, str):
            value = value.strip().lstrip("'")
        headers.append("" if value is None else str(value))
    return headers


def _locate_mapping_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
    if df is None or df.empty:
        return None, None
    normalized_cols = {_normalize_header(col): col for col in df.columns}
    code_col = None
    for variant in HS_MAPPING_HEADER_VARIANTS:
        key = _normalize_header(variant)
        if key in normalized_cols:
            code_col = normalized_cols[key]
            break
    unit_col = None
    for variant in UNIT_HEADER_VARIANTS:
        key = _normalize_header(variant)
        if key in normalized_cols:
            unit_col = normalized_cols[key]
            break
    return code_col, unit_col


def _load_hs_mapping() -> dict[str, str]:
    if not HS_MAPPING_PATH.exists():
        raise RuntimeError(f"HSCODE.xlsx not found at {HS_MAPPING_PATH}")
    try:
        sheets = pd.read_excel(
            HS_MAPPING_PATH,
            sheet_name=None,
            dtype=str,
            engine="openpyxl",
        )
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"HSCODE.xlsx: failed to load workbook: {exc}") from exc

    if not sheets:
        raise RuntimeError("HSCODE.xlsx: workbook has no sheets")

    ordered_sheet_names: list[str] = []
    if "Sheet2" in sheets:
        ordered_sheet_names.append("Sheet2")
    ordered_sheet_names.extend(name for name in sheets if name != "Sheet2")

    for sheet_name in ordered_sheet_names:
        df = sheets[sheet_name]
        code_col, unit_col = _locate_mapping_columns(df)
        if not code_col or not unit_col:
            continue
        mapping: dict[str, str] = {}
        for code_raw, unit_raw in zip(df[code_col], df[unit_col]):
            code = _digits_only(code_raw)
            if not code:
                continue
            unit_text = str(unit_raw or "").strip()
            if not unit_text:
                continue
            mapping.setdefault(code, unit_text.upper())
        if mapping:
            logging.info("Loaded HS mapping from sheet '%s' with %s entries.", sheet_name, len(mapping))
            return mapping
    raise RuntimeError("HSCODE.xlsx: could not find a sheet with HS Code and Unit columns")


HS_CODE_TO_UNIT = _load_hs_mapping()
DEBUG_HSLOOKUP = os.getenv("DEBUG_HSLOOKUP") == "1"
_LOGGED_MISSING_CODES: set[tuple[str, str, tuple[str, ...]]] = set()


def normalize_for_match(value: object) -> str:
    return _normalize_header(value)


ALWAYS_BLANK_NORMALIZED = {
    normalize_for_match(column_name) for column_name in ALWAYS_BLANK_COLUMN_NAMES
}
ALWAYS_BLANK_COLLAPSED = {
    re.sub(r"[^a-z0-9]", "", value) for value in ALWAYS_BLANK_NORMALIZED
}


def convert_to_k1(
    uploaded_bytes: bytes,
    country: str = "ID",
    template_path: str | Path | None = None,
) -> bytes:
    """Convert uploaded spreadsheet bytes into K1 import XLSX bytes."""
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

    hs_series = _get_series(source_df, HS_CANDIDATES, default="").fillna("")
    hs_clean = hs_series.apply(_hs_out_code)

    units: list[str] = []
    prefix_lists: list[list[str]] = []
    for raw_hs, hs_out in zip(hs_series, hs_clean):
        hs_out_str = str(hs_out or "")
        max_len = min(len(hs_out_str), 10)
        prefixes: list[str] = []
        for length in range(max_len, 4, -1):
            key = hs_out_str[:length]
            if not key or (prefixes and key == prefixes[-1]):
                continue
            prefixes.append(key)
        if not prefixes:
            prefixes.append(hs_out_str)
        unit = "N/A"
        matched_key = None
        for candidate in prefixes:
            mapping_value = HS_CODE_TO_UNIT.get(candidate)
            if mapping_value is None:
                continue
            normalized_unit = str(mapping_value or "").strip().upper()
            if normalized_unit in {"", "NA", "NAN", "NULL"}:
                continue
            unit = normalized_unit
            matched_key = candidate
            break
        if DEBUG_HSLOOKUP and matched_key and matched_key != hs_out_str:
            logging.info(
                "HS matched on prefix: raw='%s' chosen='%s' unit='%s'",
                raw_hs,
                matched_key,
                unit,
            )
        units.append(unit)
        prefix_lists.append(prefixes)
    unit_series = pd.Series(units, index=source_df.index, dtype="object")

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

    unit_upper = unit_series.str.upper()

    statistical_qty = pd.Series([""] * len(source_df), index=source_df.index, dtype="object")
    declared_qty = pd.Series([""] * len(source_df), index=source_df.index, dtype="object")

    kgm_mask = unit_upper == "KGM"
    if kgm_mask.any():
        net_values = net_weight_series.loc[kgm_mask].astype(float)
        statistical_qty.loc[kgm_mask] = net_values
        declared_qty.loc[kgm_mask] = net_values

    unt_mask = unit_upper == "UNT"
    if unt_mask.any():
        qty_values = quantity_series.loc[unt_mask].astype(float)
        statistical_qty.loc[unt_mask] = qty_values
        declared_qty.loc[unt_mask] = qty_values

    if DEBUG_HSLOOKUP and (unit_upper == "N/A").any():
        for position, (_, value) in enumerate(unit_series.items()):
            if str(value).upper() != "N/A":
                continue
            raw_code = hs_series.iloc[position]
            hs_out_value = hs_clean.iloc[position]
            prefixes = tuple(prefix_lists[position])
            key = (str(raw_code), str(hs_out_value), prefixes)
            if key not in _LOGGED_MISSING_CODES:
                logging.warning(
                    "HS code not in mapping: raw='%s' -> hs_out='%s' candidates=%s",
                    raw_code,
                    hs_out_value,
                    list(prefixes),
                )
                _LOGGED_MISSING_CODES.add(key)

    country_value = (country or "ID").strip().upper() or "ID"

    output["CountryOfOrigin"] = country_value
    output["HSCode"] = hs_clean
    output["StatisticalUOM"] = unit_series
    output["DeclaredUOM"] = unit_series
    output["StatisticalQty"] = statistical_qty
    output["DeclaredQty"] = declared_qty
    output["ItemAmount"] = amount_series
    output["ItemDescription"] = parts_name_series
    output["ItemDescription2"] = quantity_series
    output["ItemDescription3"] = ""

    output["ImportDutyMethod"] = "Exemption"
    output["ImportDutyRateExemptedPercentage"] = 100
    output["ImportDutyRateExemptedSpecific"] = ""

    output["SSTMethod"] = "Exemption"
    output["SSTRateExemptedPercentage"] = 100
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
        normalized_template_column = _normalize_header(template_column)
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

    changed_cells = 0

    def _sanitize_and_count(value: object) -> object:
        nonlocal changed_cells
        sanitized = _sanitize_cell(value)
        if sanitized != value:
            changed_cells += 1
        return sanitized

    try:
        final_df = final_df.map(_sanitize_and_count)  # pandas >= 2.1
    except AttributeError:
        final_df = final_df.apply(lambda col: col.map(_sanitize_and_count))  # older pandas
    if changed_cells and DEBUG_HSLOOKUP:
        logging.debug("Sanitized %s cells containing control or overlength text.", changed_cells)

    return _to_clean_xlsx_bytes(final_df)


def _to_clean_xlsx_bytes(final_df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Sheet1", index=False)
    return buffer.getvalue()


def _load_source_dataframe(uploaded_bytes: bytes) -> pd.DataFrame:
    buffer = BytesIO(uploaded_bytes)
    buffer.seek(0)
    head = buffer.read(4)
    buffer.seek(0)
    is_xlsx_like = head.startswith(b"PK")
    if is_xlsx_like:
        return pd.read_excel(buffer, engine="openpyxl")
    return _df_from_xls_bytes(uploaded_bytes)


def _load_template(template_path: Path) -> List[str]:
    ext = template_path.suffix.lower()
    if ext in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return _load_template_columns_xlsx(template_path)
    if ext == ".xls":
        return _load_template_columns_xls(template_path)
    raise ValueError(f"Unsupported template format: {template_path}")


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
    return _hs_out_code(value)


def _maybe_log_debug_samples(df: pd.DataFrame) -> None:
    if os.environ.get("DEBUG_MAPPINGS") == "1":
        records = df.head(3).to_dict(orient="records")
        logging.info("Sample mapped rows: %s", records)


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
