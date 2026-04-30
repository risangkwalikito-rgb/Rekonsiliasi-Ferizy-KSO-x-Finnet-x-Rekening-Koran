# app_before_recon_pick_date.py
from __future__ import annotations

import io
import re
import unicodedata
from datetime import date
from difflib import get_close_matches
from typing import Any

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Rekonsiliasi Sharing Fee FINNET", layout="wide")

MASTER_FEE_DATA = [
    {"instrumen_pembayaran": "VA BRI", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 777, "pajak": 0.11, "sharing_fee_excl_tax": 700, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BCA", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 278, "pajak": 0.11, "sharing_fee_excl_tax": 250, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Mandiri", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 777, "pajak": 0.11, "sharing_fee_excl_tax": 700, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BM", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 777, "pajak": 0.11, "sharing_fee_excl_tax": 700, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "DANA", "status_pg": "Secondary", "admin_fee": 0.015, "sharing_fee": 0.0012, "pajak": 0.11, "sharing_fee_excl_tax": 0.0011, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "OVO", "status_pg": "Secondary", "admin_fee": 0.0167, "sharing_fee": 0.0007, "pajak": 0.11, "sharing_fee_excl_tax": 0.0006, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BSI", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 777, "pajak": 0.11, "sharing_fee_excl_tax": 700, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Permata", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 666, "pajak": 0.11, "sharing_fee_excl_tax": 600, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "ShopeePay", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 310, "pajak": 0.11, "sharing_fee_excl_tax": 279, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Link Aja", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 765, "pajak": 0.11, "sharing_fee_excl_tax": 689, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Gopay", "status_pg": "Secondary", "admin_fee": 0.0167, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA CIMB", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 666, "pajak": 0.11, "sharing_fee_excl_tax": 600, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BTN", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 888, "pajak": 0.11, "sharing_fee_excl_tax": 800, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BJB", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 888, "pajak": 0.11, "sharing_fee_excl_tax": 800, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Pospay", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Danamon", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 722, "pajak": 0.11, "sharing_fee_excl_tax": 650, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Sumselbabel", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 722, "pajak": 0.11, "sharing_fee_excl_tax": 650, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Maybank", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 722, "pajak": 0.11, "sharing_fee_excl_tax": 650, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "BLU", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 904, "pajak": 0.11, "sharing_fee_excl_tax": 814, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Muamalat", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 722, "pajak": 0.11, "sharing_fee_excl_tax": 650, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Indomaret", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 200, "pajak": 0.11, "sharing_fee_excl_tax": 180, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Alfamart", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 200, "pajak": 0.11, "sharing_fee_excl_tax": 180, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Maspion", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Nagari", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 722, "pajak": 0.11, "sharing_fee_excl_tax": 650, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BTPN", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Neo Commerce", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Sinarmas", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Bank DKI", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 722, "pajak": 0.11, "sharing_fee_excl_tax": 650, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BPD Jatim", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 722, "pajak": 0.11, "sharing_fee_excl_tax": 650, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "PT Pos Indonesia", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "QRIS", "status_pg": "Secondary", "admin_fee": 0.007, "sharing_fee": 0.0013, "pajak": 0.11, "sharing_fee_excl_tax": 0.0011, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Pegadaian", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 1500, "pajak": 0.11, "sharing_fee_excl_tax": 1351, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Yomart", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 1500, "pajak": 0.11, "sharing_fee_excl_tax": 1351, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Delima", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 200, "pajak": 0.11, "sharing_fee_excl_tax": 180, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "OctoClicks", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "E-Pay BRI", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Permata", "status_pg": "Secondary", "admin_fee": 2220, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Credit Card", "status_pg": "Secondary", "admin_fee": 0.015, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Direct Debit", "status_pg": "Secondary", "admin_fee": 0.02, "sharing_fee": 0, "pajak": 0.11, "sharing_fee_excl_tax": 0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Finpay Code", "status_pg": "Main", "admin_fee": 2220, "sharing_fee": 1332, "pajak": 0.11, "sharing_fee_excl_tax": 1200, "settlement_hari_kerja": "H+1"},

]

MANUAL_ALIAS = {
    "bri va": "VA BRI",
    "virtual account bri": "VA BRI",
    "va bank bri": "VA BRI",
    "bca va": "VA BCA",
    "virtual account bca": "VA BCA",
    "mandiri va": "VA Mandiri",
    "virtual account mandiri": "VA Mandiri",
    "bank mega va": "VA BM",
    "mega va": "VA BM",
    "bsi va": "VA BSI",
    "permata va": "VA Permata",
    "va bank permata": "VA Permata",
    "linkaja": "Link Aja",
    "go pay": "Gopay",
    "gopay tokenization": "Gopay",
    "cimb va": "VA CIMB",
    "btn va": "VA BTN",
    "bjb va": "VA BJB",
    "danamon va": "VA Danamon",
    "sumselbabel va": "VA Sumselbabel",
    "maybank va": "VA Maybank",
    "muamalat va": "VA Muamalat",
    "nagari va": "VA Nagari",
    "btpn va": "VA BTPN",
    "neo commerce va": "VA Neo Commerce",
    "sinarmas va": "VA Sinarmas",
    "bank dki va": "VA Bank DKI",
    "bpd jatim va": "VA BPD Jatim",
    "pt pos": "PT Pos Indonesia",
    "pos indonesia": "PT Pos Indonesia",
    "epay bri": "E-Pay BRI",
    "e pay bri": "E-Pay BRI",
    "finpaycode": "Finpay Code",
    "payment code": "Finpay Code",
    "briva": "VA BRI",
    "finpay021": "Finpay Code",
    "tcash": "Link Aja",
    "vabni": "VA BNI",

}

REQUIRED_COLUMN_PATTERNS = {
    "payment_datetime": [r"payment\s*date\s*time"],
    "payment_method": [r"payment\s*method"],
    "merchant_amount": [r"merchant\s*amount"],
    "pg_provider": [r"pg\s*provider"],
}

OPTIONAL_COLUMN_PATTERNS = {
    "trx_id": [r"transaction\s*id", r"trx\s*id", r"order\s*id", r"invoice", r"reference"],
    "customer_name": [r"customer\s*name", r"nama\s*customer", r"payer\s*name"],
}

MERCHANT_AMOUNT_FALLBACK_INDEX = 20
PAYMENT_DATETIME_FALLBACK_INDEX = 2
PREFERRED_SETTLEMENT_SHEET = "detail settlement"
PG_PROVIDER_TARGET = "FINNET"


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.replace("&", " dan ")
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float, np.integer, np.floating)) and not pd.isna(value):
        return float(value)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "-"}:
        return None

    text = text.replace("rp", "").replace("%", "").replace(" ", "")
    text = re.sub(r"[^\d,.-]", "", text)

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(",") == 1 and text.count(".") == 0:
        left, right = text.split(",")
        if len(right) in {1, 2, 3, 4, 5, 6}:
            text = f"{left}.{right}"
        else:
            text = text.replace(",", "")
    else:
        text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def parse_datetime_series(series: pd.Series) -> pd.Series:
    raw = series.copy()

    if pd.api.types.is_datetime64_any_dtype(raw):
        return pd.to_datetime(raw, errors="coerce")

    parsed_direct = pd.to_datetime(raw, errors="coerce")
    numeric_raw = pd.to_numeric(raw, errors="coerce")
    excel_serial = pd.to_datetime(numeric_raw, unit="D", origin="1899-12-30", errors="coerce")
    combined = parsed_direct.fillna(excel_serial)

    if combined.notna().any():
        return combined

    text_raw = raw.astype(str).str.strip().replace({"": np.nan, "nan": np.nan, "NaT": np.nan})
    parsed_text = pd.to_datetime(text_raw, errors="coerce", dayfirst=False)
    if parsed_text.notna().any():
        return parsed_text

    return pd.to_datetime(text_raw, errors="coerce", dayfirst=True)


def detect_column_by_patterns(df: pd.DataFrame, patterns: list[str]) -> str | None:
    normalized = {col: normalize_text(col) for col in df.columns}
    for col, col_norm in normalized.items():
        if any(re.search(pattern, col_norm) for pattern in patterns):
            return col
    return None


def get_column_by_position(df: pd.DataFrame, index: int) -> str | None:
    return df.columns[index] if len(df.columns) > index else None


def select_sheet_name(workbook: pd.ExcelFile) -> str:
    sheet_names = workbook.sheet_names
    normalized_map = {normalize_text(sheet): sheet for sheet in sheet_names}
    preferred_sheet = normalized_map.get(PREFERRED_SETTLEMENT_SHEET)
    if preferred_sheet:
        return preferred_sheet
    return sheet_names[0]


def read_uploaded_file(uploaded_file: Any) -> tuple[pd.DataFrame, str]:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        raw = uploaded_file.getvalue()
        for encoding in ("utf-8", "utf-8-sig", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=encoding), "CSV"
            except Exception:
                continue
        raise ValueError("File CSV tidak bisa dibaca.")

    workbook = pd.ExcelFile(uploaded_file)
    sheet_name = select_sheet_name(workbook)
    return pd.read_excel(workbook, sheet_name=sheet_name), sheet_name


def resolve_required_columns(df: pd.DataFrame) -> dict[str, str]:
    resolved: dict[str, str] = {}
    missing: list[str] = []

    payment_datetime_col = detect_column_by_patterns(df, REQUIRED_COLUMN_PATTERNS["payment_datetime"])
    if not payment_datetime_col:
        payment_datetime_col = get_column_by_position(df, PAYMENT_DATETIME_FALLBACK_INDEX)
    if payment_datetime_col:
        resolved["payment_datetime"] = payment_datetime_col
    else:
        missing.append("payment_datetime")

    payment_method_col = detect_column_by_patterns(df, REQUIRED_COLUMN_PATTERNS["payment_method"])
    if payment_method_col:
        resolved["payment_method"] = payment_method_col
    else:
        missing.append("payment_method")

    merchant_amount_col = detect_column_by_patterns(df, REQUIRED_COLUMN_PATTERNS["merchant_amount"])
    if not merchant_amount_col:
        merchant_amount_col = get_column_by_position(df, MERCHANT_AMOUNT_FALLBACK_INDEX)
    if merchant_amount_col:
        resolved["merchant_amount"] = merchant_amount_col
    else:
        missing.append("merchant_amount")

    pg_provider_col = detect_column_by_patterns(df, REQUIRED_COLUMN_PATTERNS["pg_provider"])
    if pg_provider_col:
        resolved["pg_provider"] = pg_provider_col
    else:
        missing.append("pg_provider")

    if missing:
        labels = {
            "payment_datetime": "Payment Date Time / kolom C",
            "payment_method": "Payment Method",
            "merchant_amount": "Merchant Amount / kolom U",
            "pg_provider": "PG Provider",
        }
        raise ValueError(f"Kolom wajib tidak ditemukan: {', '.join(labels[item] for item in missing)}")

    return resolved


def resolve_optional_columns(df: pd.DataFrame) -> dict[str, str | None]:
    return {role: detect_column_by_patterns(df, patterns) for role, patterns in OPTIONAL_COLUMN_PATTERNS.items()}


def build_master_df() -> pd.DataFrame:
    master = pd.DataFrame(MASTER_FEE_DATA).copy()
    master["instrument_key"] = master["instrumen_pembayaran"].map(normalize_text)
    return master


def build_alias_map(master_df: pd.DataFrame) -> dict[str, str]:
    alias_map = {normalize_text(k): v for k, v in MANUAL_ALIAS.items()}
    for instrument in master_df["instrumen_pembayaran"]:
        key = normalize_text(instrument)
        alias_map[key] = instrument
        alias_map[key.replace(" ", "")] = instrument
    return alias_map


def resolve_instrument(value: Any, master_df: pd.DataFrame, alias_map: dict[str, str]) -> str | None:
    raw = normalize_text(value)
    if not raw:
        return None

    if raw in alias_map:
        return alias_map[raw]

    compact = raw.replace(" ", "")
    if compact in alias_map:
        return alias_map[compact]

    master_keys = dict(zip(master_df["instrument_key"], master_df["instrumen_pembayaran"]))
    if raw in master_keys:
        return master_keys[raw]

    for key, instrument in master_keys.items():
        if raw == key or raw in key or key in raw:
            return instrument

    close = get_close_matches(raw, list(master_keys.keys()), n=1, cutoff=0.75)
    if close:
        return master_keys[close[0]]

    return None


def calc_fee(amount: Any, fee_rule: Any) -> float | None:
    amount_value = parse_number(amount)
    fee_value = parse_number(fee_rule)

    if fee_value is None:
        return None
    if fee_value == 0:
        return 0.0
    if fee_value < 1:
        if amount_value is None:
            return None
        return amount_value * fee_value
    return fee_value


def filter_finnet_rows(df: pd.DataFrame, pg_provider_col: str) -> pd.DataFrame:
    provider_norm = df[pg_provider_col].map(normalize_text)
    return df.loc[provider_norm == normalize_text(PG_PROVIDER_TARGET)].copy()


def filter_by_selected_dates(
    df: pd.DataFrame,
    payment_datetime_col: str,
    start_date: date,
    end_date: date,
) -> pd.DataFrame:
    working = df.copy()
    working["payment_datetime"] = parse_datetime_series(working[payment_datetime_col])
    working = working.loc[working["payment_datetime"].notna()].copy()
    working["payment_date"] = working["payment_datetime"].dt.date
    return working.loc[(working["payment_date"] >= start_date) & (working["payment_date"] <= end_date)].copy()


def round_columns(df: pd.DataFrame, cols: list[str], decimals: int = 2) -> pd.DataFrame:
    formatted = df.copy()
    for col in cols:
        if col in formatted.columns:
            formatted[col] = pd.to_numeric(formatted[col], errors="coerce").round(decimals)
    return formatted


def build_reconciliation(
    df: pd.DataFrame,
    column_map: dict[str, str],
    optional_cols: dict[str, str | None],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    master = build_master_df()
    alias_map = build_alias_map(master)

    working = df.copy()
    working["payment_datetime"] = parse_datetime_series(working[column_map["payment_datetime"]])
    working["payment_date"] = working["payment_datetime"].dt.date
    working["payment_method_raw"] = working[column_map["payment_method"]].astype(str).str.strip()
    working["instrumen_pembayaran"] = working[column_map["payment_method"]].map(
        lambda x: resolve_instrument(x, master, alias_map)
    )
    working["merchant_amount_value"] = working[column_map["merchant_amount"]].map(parse_number)
    working["trx_count"] = 1

    if optional_cols.get("trx_id"):
        working["trx_id"] = working[optional_cols["trx_id"]].astype(str)
    else:
        working["trx_id"] = None

    merged = working.merge(master, how="left", on="instrumen_pembayaran")
    merged["expected_sharing_fee_excl_tax"] = merged.apply(
        lambda row: calc_fee(row["merchant_amount_value"], row["sharing_fee_excl_tax"]),
        axis=1,
    )

    detail = merged[
        [
            "payment_date",
            "payment_datetime",
            "trx_id",
            "payment_method_raw",
            "instrumen_pembayaran",
            "merchant_amount_value",
            "sharing_fee_excl_tax",
            "expected_sharing_fee_excl_tax",
        ]
    ].rename(
        columns={
            "merchant_amount_value": "merchant_amount",
            "sharing_fee_excl_tax": "master_sharing_fee_excl_tax",
        }
    ).copy()

    detail = detail.sort_values(["payment_datetime", "payment_method_raw"], na_position="last")
    detail = round_columns(detail, ["merchant_amount", "expected_sharing_fee_excl_tax"], 2)
    detail = round_columns(detail, ["master_sharing_fee_excl_tax"], 6)

    summary = (
        detail.groupby(["payment_method_raw", "instrumen_pembayaran", "master_sharing_fee_excl_tax"], dropna=False, as_index=False)
        .agg(
            jumlah_transaksi=("payment_method_raw", "size"),
            total_merchant_amount=("merchant_amount", "sum"),
            total_sharing_fee_excl_tax=("expected_sharing_fee_excl_tax", "sum"),
        )
        .sort_values(["payment_method_raw"], na_position="last")
    )
    summary = round_columns(summary, ["total_merchant_amount", "total_sharing_fee_excl_tax"], 2)

    unmatched = (
        detail.loc[detail["instrumen_pembayaran"].isna(), ["payment_method_raw"]]
        .drop_duplicates()
        .sort_values(["payment_method_raw"], na_position="last")
        .rename(columns={"payment_method_raw": "payment_method"})
    )

    return detail, summary, unmatched


def format_id_number(value: Any, force_decimals: int | None = None) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)) or pd.isna(value):
        return ""

    number = float(value)

    if force_decimals is not None:
        text = f"{number:,.{force_decimals}f}"
    else:
        if abs(number) >= 1:
            if abs(number - round(number)) < 1e-9:
                text = f"{number:,.0f}"
            else:
                text = f"{number:,.2f}"
        else:
            if number == 0:
                text = "0"
            else:
                text = f"{number:,.6f}".rstrip("0").rstrip(".")

    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def build_display_summary(summary_df: pd.DataFrame) -> pd.DataFrame:
    display = summary_df.copy()
    display["Payment Method"] = display["payment_method_raw"]
    display["Master Sharing Fee Exclude Tax"] = display["master_sharing_fee_excl_tax"].map(format_id_number)
    display["Jumlah Transaksi"] = display["jumlah_transaksi"].map(lambda x: format_id_number(x, 0))
    display["Total Merchant Amount"] = display["total_merchant_amount"].map(lambda x: format_id_number(x, 2))
    display["Total Sharing Fee Exclude Tax"] = display["total_sharing_fee_excl_tax"].map(lambda x: format_id_number(x, 2))
    return display[
        [
            "Payment Method",
            "Master Sharing Fee Exclude Tax",
            "Jumlah Transaksi",
            "Total Merchant Amount",
            "Total Sharing Fee Exclude Tax",
        ]
    ]


def to_excel_bytes(detail: pd.DataFrame, summary: pd.DataFrame, unmatched: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail.to_excel(writer, index=False, sheet_name="Detail Sharing Fee")
        summary.to_excel(writer, index=False, sheet_name="Rekap Ringkas")
        unmatched.to_excel(writer, index=False, sheet_name="Unmatched Method")

        for sheet_name, frame in {
            "Detail Sharing Fee": detail,
            "Rekap Ringkas": summary,
            "Unmatched Method": unmatched,
        }.items():
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            for idx, col in enumerate(frame.columns, start=1):
                max_len = max([len(str(col))] + [len(str(v)) for v in frame[col].head(200).fillna("")])
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 28)
    return output.getvalue()


st.title("Rekonsiliasi Sharing Fee FINNET")
st.caption("Pilih rentang tanggal dulu, lalu proses rekonsiliasi. Data yang dihitung hanya PG Provider = FINNET.")

with st.expander("Master fee embedded", expanded=False):
    st.dataframe(build_master_df().drop(columns=["instrument_key"]), use_container_width=True, hide_index=True)

left_col, right_col = st.columns([1, 2])

with left_col:
    st.subheader("Parameter")
    date_container = st.container()
    uploader_container = st.container()
    provider_container = st.container()

with right_col:
    st.subheader("Output")
    output_container = st.container()

with uploader_container:
    uploaded_file = st.file_uploader("Upload settlement FINNET", type=["xlsx", "xls", "csv"])

run_recon = False

if not uploaded_file:
    with date_container:
        st.info("Pilih file settlement FINNET dulu. Parameter tanggal akan tampil otomatis di atas uploader.")
    with provider_container:
        st.text_input("PG Provider yang diproses", value=PG_PROVIDER_TARGET, disabled=True)
    with output_container:
        st.info("Upload file, pilih rentang tanggal, lalu klik proses rekonsiliasi.")
else:
    try:
        source_df, sheet_name = read_uploaded_file(uploaded_file)
        column_map = resolve_required_columns(source_df)
        optional_cols = resolve_optional_columns(source_df)
        finnet_df = filter_finnet_rows(source_df, column_map["pg_provider"])
    except Exception as exc:
        with output_container:
            st.error(str(exc))
        st.stop()

    if finnet_df.empty:
        with output_container:
            st.error("Tidak ada data dengan PG Provider = FINNET.")
        st.stop()

    parsed_dates = parse_datetime_series(finnet_df[column_map["payment_datetime"]])
    valid_dates = parsed_dates.dropna().dt.date

    if valid_dates.empty:
        with output_container:
            st.error("Kolom C / Payment Date Time tidak bisa diparse menjadi tanggal.")
        st.stop()

    min_date = valid_dates.min()
    max_date = valid_dates.max()

    with date_container:
        with st.form("filter_form", clear_on_submit=False):
            st.caption(f"Sheet terpakai: {sheet_name}")
            picked = st.date_input(
                "Pilih rentang tanggal Payment Date Time",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
            )
            st.text_input("PG Provider yang diproses", value=PG_PROVIDER_TARGET, disabled=True)
            run_recon = st.form_submit_button("Proses rekonsiliasi", type="primary", use_container_width=True)

    if isinstance(picked, tuple) and len(picked) == 2:
        start_date, end_date = picked
    else:
        start_date = end_date = picked

    with provider_container:
        st.caption(
            "Kolom dipakai: "
            f"Payment Date Time = {column_map['payment_datetime']}, "
            f"Payment Method = {column_map['payment_method']}, "
            f"Merchant Amount = {column_map['merchant_amount']}, "
            f"PG Provider = {column_map['pg_provider']}"
        )

    with output_container:
        if not run_recon:
            st.info("Pilih parameter rentang tanggal dulu, lalu klik **Proses rekonsiliasi**.")
            info1, info2, info3 = st.columns(3)
            info1.metric("Total row upload", format_id_number(len(source_df), 0))
            info2.metric("Row FINNET", format_id_number(len(finnet_df), 0))
            info3.metric("Rentang tanggal tersedia", f"{min_date.strftime('%d-%m-%Y')} s.d. {max_date.strftime('%d-%m-%Y')}")
        else:
            filtered_df = filter_by_selected_dates(
                finnet_df,
                payment_datetime_col=column_map["payment_datetime"],
                start_date=start_date,
                end_date=end_date,
            )

            if filtered_df.empty:
                st.warning("Tidak ada data FINNET pada rentang tanggal yang dipilih.")
            else:
                detail_df, summary_df, unmatched_df = build_reconciliation(filtered_df, column_map, optional_cols)
                display_summary = build_display_summary(summary_df)

                st.markdown(
                    f"### Periode Payment Date Time: {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}"
                )

                top1, top2, top3 = st.columns(3)
                top1.metric("Jumlah Transaksi", format_id_number(summary_df["jumlah_transaksi"].sum(), 0))
                top2.metric("Total Merchant Amount", format_id_number(summary_df["total_merchant_amount"].sum(), 2))
                top3.metric(
                    "Total Sharing Fee Exclude Tax",
                    format_id_number(summary_df["total_sharing_fee_excl_tax"].sum(), 2),
                )

                st.dataframe(display_summary, use_container_width=True, hide_index=True)

                if not unmatched_df.empty:
                    st.warning("Masih ada Payment Method yang belum match ke master fee.")
                    st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

                excel_bytes = to_excel_bytes(detail_df, summary_df, unmatched_df)
                st.download_button(
                    label="Download hasil rekonsiliasi (.xlsx)",
                    data=excel_bytes,
                    file_name="rekonsiliasi_sharing_fee_finnet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.divider()
st.markdown(
    """
    **Logika aktif**
    - Parameter tanggal tampil di panel kiri, di atas uploader.
    - Rentang tanggal dipilih dulu sebelum proses rekonsiliasi.
    - Jika ada beberapa sheet, app akan memilih `Detail Settlement` bila ada; jika tidak ada, pakai sheet pertama.
    - `Payment Date Time` diambil dari header yang cocok atau fallback ke kolom C.
    - `Merchant Amount` diambil dari header yang cocok atau fallback ke kolom U.
    - Rekap tabel dibuat ringkas per `Payment Method` untuk rentang tanggal yang dipilih.
    """
)
