# app.py
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
        if len(right) in {1, 2, 3, 4}:
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

    parsed = pd.to_datetime(raw, errors="coerce", format="mixed")
    if parsed.notna().mean() >= 0.6:
        return parsed

    parsed_dayfirst = pd.to_datetime(raw, errors="coerce", format="mixed", dayfirst=True)
    return parsed_dayfirst


def read_uploaded_file(uploaded_file: Any) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        raw = uploaded_file.getvalue()
        for encoding in ("utf-8", "utf-8-sig", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(raw), encoding=encoding)
            except Exception:
                continue
        raise ValueError("File CSV tidak bisa dibaca.")
    return pd.read_excel(uploaded_file)


def build_master_df() -> pd.DataFrame:
    master = pd.DataFrame(MASTER_FEE_DATA).copy()
    master["instrument_key"] = master["instrumen_pembayaran"].map(normalize_text)
    return master


def detect_column_by_patterns(df: pd.DataFrame, patterns: list[str]) -> str | None:
    normalized = {col: normalize_text(col) for col in df.columns}
    for col, col_norm in normalized.items():
        if any(re.search(pattern, col_norm) for pattern in patterns):
            return col
    return None


def resolve_required_columns(df: pd.DataFrame) -> dict[str, str]:
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for role, patterns in REQUIRED_COLUMN_PATTERNS.items():
        found = detect_column_by_patterns(df, patterns)
        if not found:
            missing.append(role)
        else:
            resolved[role] = found

    if missing:
        labels = {
            "payment_datetime": "Payment Date Time",
            "payment_method": "Payment Method",
            "merchant_amount": "Merchant Amount",
            "pg_provider": "PG Provider",
        }
        missing_labels = ", ".join(labels[item] for item in missing)
        raise ValueError(f"Kolom wajib tidak ditemukan: {missing_labels}")
    return resolved


def resolve_optional_columns(df: pd.DataFrame) -> dict[str, str | None]:
    return {
        role: detect_column_by_patterns(df, patterns)
        for role, patterns in OPTIONAL_COLUMN_PATTERNS.items()
    }


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


def round_columns(df: pd.DataFrame, columns: list[str], decimals: int = 2) -> pd.DataFrame:
    result = df.copy()
    for col in columns:
        if col in result.columns:
            result[col] = pd.to_numeric(result[col], errors="coerce").round(decimals)
    return result


def filter_finnet_rows(df: pd.DataFrame, pg_provider_col: str) -> pd.DataFrame:
    provider_norm = df[pg_provider_col].map(normalize_text)
    finnet_mask = provider_norm.eq("finnet")
    return df.loc[finnet_mask].copy()


def filter_by_selected_dates(df: pd.DataFrame, payment_datetime_col: str, start_date: date, end_date: date) -> pd.DataFrame:
    working = df.copy()
    working["payment_datetime"] = parse_datetime_series(working[payment_datetime_col])
    working["payment_date"] = working["payment_datetime"].dt.date
    working = working.loc[working["payment_datetime"].notna()].copy()
    return working.loc[(working["payment_date"] >= start_date) & (working["payment_date"] <= end_date)].copy()


def build_reconciliation(df: pd.DataFrame, column_map: dict[str, str], optional_cols: dict[str, str | None]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    master = build_master_df()
    alias_map = build_alias_map(master)

    working = df.copy()
    working["payment_datetime"] = parse_datetime_series(working[column_map["payment_datetime"]])
    working["payment_date"] = working["payment_datetime"].dt.date
    working["payment_method_raw"] = working[column_map["payment_method"]]
    working["instrumen_pembayaran"] = working[column_map["payment_method"]].map(lambda x: resolve_instrument(x, master, alias_map))
    working["merchant_amount_value"] = working[column_map["merchant_amount"]].map(parse_number)
    working["pg_provider_value"] = working[column_map["pg_provider"]]

    if optional_cols.get("trx_id"):
        working["trx_id"] = working[optional_cols["trx_id"]].astype(str)
    else:
        working["trx_id"] = None

    if optional_cols.get("customer_name"):
        working["customer_name"] = working[optional_cols["customer_name"]].astype(str)
    else:
        working["customer_name"] = None

    merged = working.merge(master, how="left", on="instrumen_pembayaran")
    merged["expected_admin_fee"] = merged.apply(lambda row: calc_fee(row["merchant_amount_value"], row["admin_fee"]), axis=1)
    merged["expected_sharing_fee"] = merged.apply(lambda row: calc_fee(row["merchant_amount_value"], row["sharing_fee"]), axis=1)
    merged["expected_sharing_fee_excl_tax"] = merged.apply(lambda row: calc_fee(row["merchant_amount_value"], row["sharing_fee_excl_tax"]), axis=1)
    merged["trx_count"] = 1

    detail_columns = [
        "payment_date",
        "payment_datetime",
        "trx_id",
        "customer_name",
        "pg_provider_value",
        "payment_method_raw",
        "instrumen_pembayaran",
        "status_pg",
        "settlement_hari_kerja",
        "merchant_amount_value",
        "admin_fee",
        "sharing_fee",
        "pajak",
        "sharing_fee_excl_tax",
        "expected_admin_fee",
        "expected_sharing_fee",
        "expected_sharing_fee_excl_tax",
    ]
    detail = merged[[col for col in detail_columns if col in merged.columns]].copy()
    detail = detail.rename(columns={
        "pg_provider_value": "pg_provider",
        "merchant_amount_value": "merchant_amount",
    })
    detail = detail.sort_values(["payment_date", "instrumen_pembayaran", "payment_datetime"], na_position="last")
    detail = round_columns(
        detail,
        ["merchant_amount", "expected_admin_fee", "expected_sharing_fee", "expected_sharing_fee_excl_tax"],
        decimals=2,
    )
    detail = round_columns(detail, ["admin_fee", "sharing_fee", "sharing_fee_excl_tax", "pajak"], decimals=6)

    summary = (
        merged.groupby(["payment_date", "instrumen_pembayaran"], dropna=False, as_index=False)
        .agg(
            trx_count=("trx_count", "sum"),
            total_merchant_amount=("merchant_amount_value", "sum"),
            total_expected_admin_fee=("expected_admin_fee", "sum"),
            total_expected_sharing_fee=("expected_sharing_fee", "sum"),
            total_expected_sharing_fee_excl_tax=("expected_sharing_fee_excl_tax", "sum"),
        )
        .sort_values(["payment_date", "instrumen_pembayaran"], na_position="last")
    )
    summary = round_columns(
        summary,
        ["total_merchant_amount", "total_expected_admin_fee", "total_expected_sharing_fee", "total_expected_sharing_fee_excl_tax"],
        decimals=2,
    )

    unmatched = (
        detail.loc[detail["instrumen_pembayaran"].isna(), ["payment_date", "payment_method_raw"]]
        .drop_duplicates()
        .sort_values(["payment_date", "payment_method_raw"], na_position="last")
    )

    return detail, summary, unmatched


def to_excel_bytes(detail: pd.DataFrame, summary: pd.DataFrame, unmatched: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail.to_excel(writer, index=False, sheet_name="Detail Sharing Fee")
        summary.to_excel(writer, index=False, sheet_name="Rekap Sharing Fee")
        unmatched.to_excel(writer, index=False, sheet_name="Unmatched Method")

        for sheet_name, frame in {
            "Detail Sharing Fee": detail,
            "Rekap Sharing Fee": summary,
            "Unmatched Method": unmatched,
        }.items():
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            for idx, col in enumerate(frame.columns, start=1):
                max_len = max([len(str(col))] + [len(str(v)) for v in frame[col].head(200).fillna("")])
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 28)
    return output.getvalue()



st.title("Rekonsiliasi Sharing Fee FINNET")
st.caption(
    "Aplikasi membaca kolom Payment Date Time, Payment Method, Merchant Amount, dan PG Provider. "
    "Data yang diproses hanya PG Provider = FINNET."
)

with st.expander("Master fee embedded", expanded=False):
    st.dataframe(build_master_df().drop(columns=["instrument_key"]), use_container_width=True, hide_index=True)

left_col, right_col = st.columns([1, 2])

with left_col:
    st.subheader("Parameter")
    date_filter_placeholder = st.empty()
    provider_placeholder = st.empty()
    uploaded_file = st.file_uploader(
        "Upload settlement FINNET",
        type=["xlsx", "xls", "csv"],
    )

with right_col:
    st.subheader("Output")
    output_placeholder = st.empty()

if not uploaded_file:
    with date_filter_placeholder.container():
        st.info("Parameter tanggal akan muncul otomatis di atas uploader setelah file settlement FINNET diupload.")
    with provider_placeholder.container():
        st.text_input("PG Provider yang diproses", value="FINNET", disabled=True)
    with output_placeholder.container():
        st.info("Silakan upload file settlement FINNET untuk mulai rekonsiliasi.")

if uploaded_file:
    try:
        source_df = read_uploaded_file(uploaded_file)
        column_map = resolve_required_columns(source_df)
        optional_cols = resolve_optional_columns(source_df)
    except Exception as exc:
        with output_placeholder.container():
            st.error(str(exc))
        st.stop()

    parsed_dates = parse_datetime_series(source_df[column_map["payment_datetime"]])
    valid_dates = parsed_dates.dropna().dt.date

    if valid_dates.empty:
        with output_placeholder.container():
            st.error("Kolom Payment Date Time tidak bisa diparse menjadi tanggal.")
        st.stop()

    min_date = valid_dates.min()
    max_date = valid_dates.max()

    with date_filter_placeholder.container():
        selected_range = st.date_input(
            "Pilih rentang tanggal Payment Date Time",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

    with provider_placeholder.container():
        st.text_input("PG Provider yang diproses", value="FINNET", disabled=True)

    if isinstance(selected_range, tuple) and len(selected_range) == 2:
        start_date, end_date = selected_range
    else:
        start_date = end_date = selected_range

    finnet_df = filter_finnet_rows(source_df, column_map["pg_provider"])
    non_finnet_rows = len(source_df) - len(finnet_df)

    if finnet_df.empty:
        with output_placeholder.container():
            st.error("Tidak ada data dengan PG Provider = FINNET.")
        st.stop()

    filtered_df = filter_by_selected_dates(
        finnet_df,
        payment_datetime_col=column_map["payment_datetime"],
        start_date=start_date,
        end_date=end_date,
    )

    with right_col:
        st.subheader("Preview data")
        st.dataframe(source_df.head(20), use_container_width=True)

        st.subheader("Kolom yang dipakai")
        mapping_preview = pd.DataFrame(
            [
                {"Parameter": "Payment Date Time", "Kolom File": column_map["payment_datetime"]},
                {"Parameter": "Payment Method", "Kolom File": column_map["payment_method"]},
                {"Parameter": "Merchant Amount", "Kolom File": column_map["merchant_amount"]},
                {"Parameter": "PG Provider", "Kolom File": column_map["pg_provider"]},
                {"Parameter": "TRX ID (opsional)", "Kolom File": optional_cols.get("trx_id") or "-"},
                {"Parameter": "Customer Name (opsional)", "Kolom File": optional_cols.get("customer_name") or "-"},
            ]
        )
        st.dataframe(mapping_preview, use_container_width=True, hide_index=True)

        metric1, metric2, metric3 = st.columns(3)
        metric1.metric("Total row upload", f"{len(source_df):,}")
        metric2.metric("Row PG Provider = FINNET", f"{len(finnet_df):,}")
        metric3.metric("Row non-FINNET diabaikan", f"{non_finnet_rows:,}")

        if filtered_df.empty:
            st.warning("Tidak ada data FINNET pada rentang tanggal yang dipilih.")
        else:
            if st.button("Proses rekonsiliasi", type="primary", use_container_width=True):
                detail_df, summary_df, unmatched_df = build_reconciliation(filtered_df, column_map, optional_cols)

                top1, top2, top3, top4 = st.columns(4)
                top1.metric("Transaksi terproses", f"{len(detail_df):,}")
                top2.metric("Total merchant amount", f"{detail_df['merchant_amount'].fillna(0).sum():,.2f}")
                top3.metric("Total expected sharing fee", f"{detail_df['expected_sharing_fee'].fillna(0).sum():,.2f}")
                top4.metric(
                    "Payment method belum match",
                    f"{unmatched_df['payment_method_raw'].nunique() if not unmatched_df.empty else 0:,}",
                )

                tab1, tab2, tab3 = st.tabs(["Detail Sharing Fee", "Rekap Sharing Fee", "Unmatched Method"])

                with tab1:
                    st.dataframe(detail_df, use_container_width=True, hide_index=True)
                with tab2:
                    st.dataframe(summary_df, use_container_width=True, hide_index=True)
                with tab3:
                    if unmatched_df.empty:
                        st.success("Semua Payment Method berhasil dicocokkan ke master fee.")
                    else:
                        st.warning("Masih ada Payment Method yang belum match ke master fee embedded.")
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
    **Logika final**
    - Parameter tanggal tampil otomatis di atas uploader settlement FINNET.
    - `Payment Date Time` diparse menjadi datetime lalu difilter sesuai rentang tanggal yang dipilih.
    - `Payment Method` dipakai untuk join ke master fee embedded.
    - `Merchant Amount` dipakai sebagai amount dasar akumulasi dan perhitungan fee persentase.
    - `PG Provider` wajib FINNET; data provider lain otomatis diabaikan.
    """
)
