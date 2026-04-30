
from __future__ import annotations

from datetime import date
import io
import re
import unicodedata
from typing import Any

import pandas as pd
import streamlit as st
from openpyxl import Workbook


st.set_page_config(page_title="Rekonsiliasi Sharing Fee FINNET", layout="wide")


MASTER_FEE_DATA = [
    {"instrumen_pembayaran": "VA BRI", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 777.0, "pajak": 0.11, "sharing_fee_excl_tax": 700.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BCA", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 278.0, "pajak": 0.11, "sharing_fee_excl_tax": 250.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Mandiri", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 777.0, "pajak": 0.11, "sharing_fee_excl_tax": 700.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BNI", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 777.0, "pajak": 0.11, "sharing_fee_excl_tax": 700.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "DANA", "status_pg": "Secondary", "admin_fee": 0.015, "sharing_fee": 0.0012, "pajak": 0.11, "sharing_fee_excl_tax": 0.0011, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "OVO", "status_pg": "Secondary", "admin_fee": 0.0167, "sharing_fee": 0.0007, "pajak": 0.11, "sharing_fee_excl_tax": 0.0006, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BSI", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 777.0, "pajak": 0.11, "sharing_fee_excl_tax": 700.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Permata", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 666.0, "pajak": 0.11, "sharing_fee_excl_tax": 600.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "ShopeePay", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 310.0, "pajak": 0.11, "sharing_fee_excl_tax": 279.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Link Aja", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 765.0, "pajak": 0.11, "sharing_fee_excl_tax": 689.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Gopay", "status_pg": "Secondary", "admin_fee": 0.0167, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA CIMB", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 666.0, "pajak": 0.11, "sharing_fee_excl_tax": 600.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BTN", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 888.0, "pajak": 0.11, "sharing_fee_excl_tax": 800.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BJB", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 888.0, "pajak": 0.11, "sharing_fee_excl_tax": 800.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Pospay", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Danamon", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 722.0, "pajak": 0.11, "sharing_fee_excl_tax": 650.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Sumselbabel", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 722.0, "pajak": 0.11, "sharing_fee_excl_tax": 650.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Maybank", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 722.0, "pajak": 0.11, "sharing_fee_excl_tax": 650.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "BLU", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 904.0, "pajak": 0.11, "sharing_fee_excl_tax": 814.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Muamalat", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 722.0, "pajak": 0.11, "sharing_fee_excl_tax": 650.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Indomaret", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 200.0, "pajak": 0.11, "sharing_fee_excl_tax": 180.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Alfamart", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 200.0, "pajak": 0.11, "sharing_fee_excl_tax": 180.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Maspion", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Nagari", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 722.0, "pajak": 0.11, "sharing_fee_excl_tax": 650.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BTPN", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Neo Commerce", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Sinarmas", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA Bank DKI", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 722.0, "pajak": 0.11, "sharing_fee_excl_tax": 650.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "VA BPD Jatim", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 722.0, "pajak": 0.11, "sharing_fee_excl_tax": 650.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "PT Pos Indonesia", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "QRIS", "status_pg": "Secondary", "admin_fee": 0.007, "sharing_fee": 0.0013, "pajak": 0.11, "sharing_fee_excl_tax": 0.0011, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Pegadaian", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 1500.0, "pajak": 0.11, "sharing_fee_excl_tax": 1351.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Yomart", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 1500.0, "pajak": 0.11, "sharing_fee_excl_tax": 1351.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Delima", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 200.0, "pajak": 0.11, "sharing_fee_excl_tax": 180.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "OctoClicks", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "E-Pay BRI", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Permata", "status_pg": "Secondary", "admin_fee": 2220.0, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Credit Card", "status_pg": "Secondary", "admin_fee": 0.015, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Direct Debit", "status_pg": "Secondary", "admin_fee": 0.02, "sharing_fee": 0.0, "pajak": 0.11, "sharing_fee_excl_tax": 0.0, "settlement_hari_kerja": "H+1"},
    {"instrumen_pembayaran": "Finpay Code", "status_pg": "Main", "admin_fee": 2220.0, "sharing_fee": 1332.0, "pajak": 0.11, "sharing_fee_excl_tax": 1200.0, "settlement_hari_kerja": "H+1"},
]

ALIAS_MAP = {
    "va bri": "VA BRI",
    "bri va": "VA BRI",
    "briva": "VA BRI",
    "virtual account bri": "VA BRI",
    "va bca": "VA BCA",
    "bca va": "VA BCA",
    "virtual account bca": "VA BCA",
    "va mandiri": "VA Mandiri",
    "mandiri va": "VA Mandiri",
    "virtual account mandiri": "VA Mandiri",
    "va bni": "VA BNI",
    "bni va": "VA BNI",
    "vabni": "VA BNI",
    "virtual account bni": "VA BNI",
    "dana": "DANA",
    "ovo": "OVO",
    "va bsi": "VA BSI",
    "vabsi": "VA BSI",
    "va permata": "VA Permata",
    "permata va": "VA Permata",
    "shopeepay": "ShopeePay",
    "shopee pay": "ShopeePay",
    "link aja": "Link Aja",
    "linkaja": "Link Aja",
    "tcash": "Link Aja",
    "gopay": "Gopay",
    "go pay": "Gopay",
    "va cimb": "VA CIMB",
    "vacimb": "VA CIMB",
    "va btn": "VA BTN",
    "vabtn": "VA BTN",
    "va bjb": "VA BJB",
    "vabjb": "VA BJB",
    "pospay": "Pospay",
    "va danamon": "VA Danamon",
    "vadanamon": "VA Danamon",
    "va sumselbabel": "VA Sumselbabel",
    "vasumselbabel": "VA Sumselbabel",
    "va maybank": "VA Maybank",
    "vamaybank": "VA Maybank",
    "blu": "BLU",
    "va muamalat": "VA Muamalat",
    "vamuamalat": "VA Muamalat",
    "indomaret": "Indomaret",
    "alfamart": "Alfamart",
    "va maspion": "VA Maspion",
    "vamaspion": "VA Maspion",
    "va nagari": "VA Nagari",
    "vanagari": "VA Nagari",
    "va btpn": "VA BTPN",
    "vabtpn": "VA BTPN",
    "va neo commerce": "VA Neo Commerce",
    "vaneocommerce": "VA Neo Commerce",
    "va sinarmas": "VA Sinarmas",
    "vasinarmas": "VA Sinarmas",
    "va bank dki": "VA Bank DKI",
    "vabankdki": "VA Bank DKI",
    "va bpd jatim": "VA BPD Jatim",
    "vabpdjatim": "VA BPD Jatim",
    "pt pos indonesia": "PT Pos Indonesia",
    "pt pos": "PT Pos Indonesia",
    "qris": "QRIS",
    "pegadaian": "Pegadaian",
    "yomart": "Yomart",
    "delima": "Delima",
    "octoclicks": "OctoClicks",
    "e-pay bri": "E-Pay BRI",
    "epay bri": "E-Pay BRI",
    "permata": "Permata",
    "credit card": "Credit Card",
    "direct debit": "Direct Debit",
    "finpay code": "Finpay Code",
    "finpaycode": "Finpay Code",
    "finpay021": "Finpay Code",
}

DISPLAY_COLUMNS = [
    "Payment Method",
    "Jumlah Transaksi",
    "Master Sharing Fee",
    "Total Sharing Fee Exclude Tax",
]


def normalize_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[_\-/]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_number(value: Any) -> float | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = text.replace("Rp", "").replace("rp", "").replace(" ", "")
    text = re.sub(r"[^0-9,.\-]", "", text)

    if not text:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text and "." not in text:
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def format_number_id(value: Any, decimals: int = 2, trim_zero: bool = False) -> str:
    if value is None or pd.isna(value):
        return ""
    formatted = f"{float(value):,.{decimals}f}"
    formatted = formatted.replace(",", "_").replace(".", ",").replace("_", ".")
    if trim_zero and "," in formatted:
        formatted = formatted.rstrip("0").rstrip(",")
    return formatted


def format_integer_id(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return f"{int(round(float(value))):,}".replace(",", ".")


def normalize_date_range(value: Any, fallback_start: date, fallback_end: date) -> tuple[date, date]:
    if isinstance(value, (tuple, list)):
        if len(value) >= 2:
            start_date, end_date = value[0], value[1]
        elif len(value) == 1:
            start_date = end_date = value[0]
        else:
            start_date, end_date = fallback_start, fallback_end
    elif value:
        start_date = end_date = value
    else:
        start_date, end_date = fallback_start, fallback_end

    if start_date > end_date:
        start_date, end_date = end_date, start_date

    return start_date, end_date


def clamp_date_range(start_date: date, end_date: date, min_date: date, max_date: date) -> tuple[date, date]:
    start_date = max(min_date, min(start_date, max_date))
    end_date = max(min_date, min(end_date, max_date))
    if start_date > end_date:
        return min_date, max_date
    return start_date, end_date


def build_master_df() -> pd.DataFrame:
    master = pd.DataFrame(MASTER_FEE_DATA).copy()
    master["method_key"] = master["instrumen_pembayaran"].map(normalize_text)
    return master


def choose_excel_sheet(sheet_names: list[str]) -> str:
    for sheet_name in sheet_names:
        if normalize_text(sheet_name) == "detail settlement":
            return sheet_name
    return sheet_names[0]


@st.cache_data(show_spinner=False)
def read_uploaded_file(file_bytes: bytes, file_name: str) -> tuple[pd.DataFrame, str]:
    lower_name = file_name.lower()

    if lower_name.endswith(".csv"):
        for encoding in ("utf-8", "utf-8-sig", "latin1"):
            try:
                return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding), "CSV"
            except Exception:
                continue
        raise ValueError("File CSV tidak bisa dibaca.")

    workbook = pd.ExcelFile(io.BytesIO(file_bytes))
    selected_sheet = choose_excel_sheet(workbook.sheet_names)
    dataframe = pd.read_excel(workbook, sheet_name=selected_sheet)
    return dataframe, selected_sheet


def detect_column_by_name(columns: list[str], target: str) -> str | None:
    target_key = normalize_text(target)
    for column in columns:
        if normalize_text(column) == target_key:
            return column
    return None


def get_merchant_amount_column(df: pd.DataFrame) -> tuple[str, str]:
    header_column = detect_column_by_name(df.columns.tolist(), "Merchant Amount")
    if header_column is not None:
        return header_column, "Merchant Amount"

    if df.shape[1] >= 21:
        return df.columns[20], "Kolom U"

    raise ValueError("Kolom wajib tidak ditemukan: Merchant Amount / kolom U")


def parse_payment_datetime(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    numeric_series = pd.to_numeric(series, errors="coerce")
    if numeric_series.notna().any():
        parsed_serial = pd.to_datetime(numeric_series, unit="D", origin="1899-12-30", errors="coerce")
        if parsed_serial.notna().sum() >= max(1, int(len(series) * 0.3)):
            return parsed_serial

    text = series.astype(str).str.strip()
    text = text.replace({"nan": "", "NaT": "", "None": ""})

    for fmt in (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y %I:%M %p",
    ):
        parsed = pd.to_datetime(text, format=fmt, errors="coerce")
        if parsed.notna().any():
            return parsed

    return pd.to_datetime(text, errors="coerce", dayfirst=False)


def resolve_payment_method(value: Any, master_df: pd.DataFrame) -> str | None:
    raw = normalize_text(value)
    if not raw:
        return None

    if raw in ALIAS_MAP:
        return ALIAS_MAP[raw]

    master_lookup = dict(zip(master_df["method_key"], master_df["instrumen_pembayaran"]))
    if raw in master_lookup:
        return master_lookup[raw]

    compact = raw.replace(" ", "")
    for key, canonical in master_lookup.items():
        if compact == key.replace(" ", ""):
            return canonical

    for key, canonical in master_lookup.items():
        if raw in key or key in raw:
            return canonical

    return None


def prepare_dataset(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if df.shape[1] < 3:
        raise ValueError("Kolom C / Payment Date Time tidak ditemukan.")

    payment_method_column = detect_column_by_name(df.columns.tolist(), "Payment Method")
    if payment_method_column is None:
        raise ValueError("Kolom wajib tidak ditemukan: Payment Method")

    merchant_name_column = detect_column_by_name(df.columns.tolist(), "Merchant Name")
    if merchant_name_column is None:
        raise ValueError("Kolom wajib tidak ditemukan: Merchant Name")

    pg_provider_column = detect_column_by_name(df.columns.tolist(), "PG Provider")
    if pg_provider_column is None:
        raise ValueError("Kolom wajib tidak ditemukan: PG Provider")

    merchant_amount_column, merchant_amount_source = get_merchant_amount_column(df)

    working = df.copy()
    working["payment_datetime"] = parse_payment_datetime(working.iloc[:, 2])

    if working["payment_datetime"].notna().sum() == 0:
        raise ValueError(
            "Kolom C (Payment Date Time) tidak bisa diparse. Format yang didukung: m/dd/yyyy hh:mm:ss, mm/dd/yyyy hh:mm:ss, atau datetime native Excel."
        )

    working["payment_date"] = working["payment_datetime"].dt.date
    working["payment_method_raw"] = working[payment_method_column].astype(str).str.strip()
    working["merchant_name_raw"] = working[merchant_name_column].astype(str).str.strip()
    working["pg_provider_raw"] = working[pg_provider_column].astype(str).str.strip()
    working["merchant_amount"] = working[merchant_amount_column].map(parse_number)

    working = working.loc[working["payment_datetime"].notna()].copy()
    working = working.loc[working["payment_method_raw"].ne("")].copy()
    working = working.loc[working["merchant_name_raw"].ne("")].copy()
    working = working.loc[working["merchant_amount"].notna()].copy()
    working = working.loc[working["merchant_amount"] != 0].copy()
    working = working.loc[working["pg_provider_raw"].map(normalize_text).eq("finnet")].copy()

    if working.empty:
        raise ValueError("Tidak ada data FINNET yang valid setelah filter PG Provider, tanggal, dan amount.")

    return working, merchant_amount_source


def build_branch_summary(
    prepared_df: pd.DataFrame,
    start_date: date,
    end_date: date,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    master_df = build_master_df()

    filtered = prepared_df.loc[prepared_df["payment_date"].between(start_date, end_date)].copy()
    if filtered.empty:
        raise ValueError("Tidak ada data FINNET pada rentang tanggal yang dipilih.")

    filtered["cabang"] = filtered["merchant_name_raw"].fillna("").astype(str).str.strip()
    filtered["cabang"] = filtered["cabang"].replace("", "(Tanpa Cabang)")
    filtered["payment_method"] = filtered["payment_method_raw"].map(lambda value: resolve_payment_method(value, master_df))
    filtered["trx_count"] = 1
    filtered["payment_method_key"] = filtered["payment_method"].map(normalize_text)

    master_lookup = master_df[["instrumen_pembayaran", "method_key", "sharing_fee"]].drop_duplicates().copy()

    merged = filtered.merge(
        master_lookup,
        how="left",
        left_on="payment_method_key",
        right_on="method_key",
    )

    matched = merged.loc[merged["payment_method"].notna()].copy()
    matched["Master Sharing Fee"] = pd.to_numeric(matched["sharing_fee"], errors="coerce")
    matched["Total Sharing Fee Exclude Tax"] = matched["trx_count"] * matched["Master Sharing Fee"].fillna(0)

    summary = (
        matched.groupby(["cabang", "payment_method"], dropna=False, as_index=False)
        .agg(
            jumlah_transaksi=("trx_count", "sum"),
            master_sharing_fee=("Master Sharing Fee", "first"),
            total_sharing_fee_exclude_tax=("Total Sharing Fee Exclude Tax", "sum"),
        )
        .sort_values(["cabang", "payment_method"], na_position="last")
        .reset_index(drop=True)
    )

    summary = summary.rename(
        columns={
            "cabang": "Cabang",
            "payment_method": "Payment Method",
            "jumlah_transaksi": "Jumlah Transaksi",
            "master_sharing_fee": "Master Sharing Fee",
            "total_sharing_fee_exclude_tax": "Total Sharing Fee Exclude Tax",
        }
    )

    grand_total = pd.DataFrame(
        [
            {
                "Payment Method": "GRAND TOTAL",
                "Jumlah Transaksi": int(summary["Jumlah Transaksi"].sum()) if not summary.empty else 0,
                "Master Sharing Fee": None,
                "Total Sharing Fee Exclude Tax": float(summary["Total Sharing Fee Exclude Tax"].sum()) if not summary.empty else 0.0,
            }
        ]
    )

    unmatched_method = merged["payment_method"].isna()
    missing_master_fee = merged["payment_method"].notna() & merged["sharing_fee"].isna()

    unmatched = (
        merged.loc[unmatched_method | missing_master_fee, ["cabang", "payment_method_raw", "payment_method"]]
        .drop_duplicates()
        .rename(
            columns={
                "cabang": "Cabang",
                "payment_method_raw": "Payment Method Raw",
                "payment_method": "Resolved Payment Method",
            }
        )
        .sort_values(["Cabang", "Payment Method Raw"], na_position="last")
        .reset_index(drop=True)
    )

    return summary, grand_total, unmatched


def format_table_for_display(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    if "Jumlah Transaksi" in display_df.columns:
        display_df["Jumlah Transaksi"] = display_df["Jumlah Transaksi"].map(format_integer_id)
    if "Master Sharing Fee" in display_df.columns:
        display_df["Master Sharing Fee"] = display_df["Master Sharing Fee"].map(
            lambda value: format_number_id(value, decimals=4, trim_zero=True)
        )
    if "Total Sharing Fee Exclude Tax" in display_df.columns:
        display_df["Total Sharing Fee Exclude Tax"] = display_df["Total Sharing Fee Exclude Tax"].map(
            lambda value: format_number_id(value, decimals=2)
        )
    return display_df


def autosize_worksheet(worksheet: Any, dataframe: pd.DataFrame) -> None:
    worksheet.freeze_panes = "A2"
    for column_index, column_name in enumerate(dataframe.columns, start=1):
        values = [len(str(column_name))]
        if not dataframe.empty:
            values.extend(len(str(value)) for value in dataframe[column_name].head(300).fillna(""))
        worksheet.column_dimensions[worksheet.cell(row=1, column=column_index).column_letter].width = min(max(values) + 2, 32)


def to_excel_bytes(
    summary_df: pd.DataFrame,
    grand_total_df: pd.DataFrame,
    unmatched_df: pd.DataFrame,
    file_infos: list[dict[str, Any]],
    start_date: date,
    end_date: date,
) -> bytes:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Rekonsiliasi"

    current_row = 1
    worksheet.cell(row=current_row, column=1, value=f"Periode Payment Date Time: {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}")
    current_row += 2

    for cabang in summary_df["Cabang"].drop_duplicates().tolist():
        cabang_df = summary_df.loc[summary_df["Cabang"] == cabang, DISPLAY_COLUMNS].reset_index(drop=True)
        worksheet.cell(row=current_row, column=1, value=cabang)
        current_row += 1

        for col_idx, column_name in enumerate(DISPLAY_COLUMNS, start=1):
            worksheet.cell(row=current_row, column=col_idx, value=column_name)
        current_row += 1

        for _, row in cabang_df.iterrows():
            worksheet.cell(row=current_row, column=1, value=row["Payment Method"])
            worksheet.cell(row=current_row, column=2, value=int(row["Jumlah Transaksi"]))
            worksheet.cell(row=current_row, column=3, value=None if pd.isna(row["Master Sharing Fee"]) else float(row["Master Sharing Fee"]))
            worksheet.cell(row=current_row, column=4, value=float(row["Total Sharing Fee Exclude Tax"]))
            current_row += 1

        current_row += 1

    worksheet.cell(row=current_row, column=1, value="GRAND TOTAL")
    current_row += 1
    for col_idx, column_name in enumerate(DISPLAY_COLUMNS, start=1):
        worksheet.cell(row=current_row, column=col_idx, value=column_name)
    current_row += 1

    grand_row = grand_total_df.iloc[0]
    worksheet.cell(row=current_row, column=1, value=grand_row["Payment Method"])
    worksheet.cell(row=current_row, column=2, value=int(grand_row["Jumlah Transaksi"]))
    worksheet.cell(row=current_row, column=3, value=None)
    worksheet.cell(row=current_row, column=4, value=float(grand_row["Total Sharing Fee Exclude Tax"]))

    worksheet.freeze_panes = "A3"
    for width_col, width in {"A": 30, "B": 18, "C": 20, "D": 28}.items():
        worksheet.column_dimensions[width_col].width = width

    unmatched_sheet = workbook.create_sheet("Unmatched Method")
    unmatched_sheet.append(unmatched_df.columns.tolist())
    for row in unmatched_df.itertuples(index=False):
        unmatched_sheet.append(list(row))
    autosize_worksheet(unmatched_sheet, unmatched_df)

    info_sheet = workbook.create_sheet("Info File")
    info_df = pd.DataFrame(file_infos)
    if not info_df.empty:
        info_sheet.append(info_df.columns.tolist())
        for row in info_df.itertuples(index=False):
            info_sheet.append(list(row))
    autosize_worksheet(info_sheet, info_df if not info_df.empty else pd.DataFrame(columns=["file_name"]))

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()


def get_file_token(uploaded_files: list[Any] | None) -> str:
    if not uploaded_files:
        return "__no_file__"
    return "||".join(f"{uploaded_file.name}|{uploaded_file.size}" for uploaded_file in uploaded_files)


def load_uploaded_files(uploaded_files: list[Any]) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    prepared_frames: list[pd.DataFrame] = []
    file_infos: list[dict[str, Any]] = []
    errors: list[str] = []

    for uploaded_file in uploaded_files:
        try:
            file_bytes = uploaded_file.getvalue()
            source_df, selected_sheet = read_uploaded_file(file_bytes, uploaded_file.name)
            prepared_df, merchant_amount_source = prepare_dataset(source_df)
        except Exception as exc:
            errors.append(f"{uploaded_file.name}: {exc}")
            continue

        prepared_frames.append(prepared_df.assign(source_file=uploaded_file.name))
        file_infos.append(
            {
                "file_name": uploaded_file.name,
                "selected_sheet": selected_sheet,
                "merchant_amount_source": merchant_amount_source,
                "row_count": int(len(prepared_df)),
                "min_date": prepared_df["payment_date"].min(),
                "max_date": prepared_df["payment_date"].max(),
            }
        )

    if errors:
        raise ValueError("\n".join(errors))

    if not prepared_frames:
        raise ValueError("Tidak ada file valid yang bisa diproses.")

    combined_df = pd.concat(prepared_frames, ignore_index=True)
    return combined_df, file_infos


def render_branch_sections(summary_df: pd.DataFrame) -> None:
    for cabang in summary_df["Cabang"].drop_duplicates().tolist():
        st.markdown(f"### {cabang}")
        cabang_df = summary_df.loc[summary_df["Cabang"] == cabang, DISPLAY_COLUMNS].reset_index(drop=True)
        st.dataframe(format_table_for_display(cabang_df), use_container_width=True, hide_index=True)


def main() -> None:
    st.title("Rekonsiliasi Sharing Fee FINNET")
    st.caption("Pilih rentang tanggal, upload settlement FINNET, lalu proses rekonsiliasi.")

    today = date.today()
    st.session_state.setdefault("date_range_widget", (today, today))
    st.session_state.setdefault("loaded_file_token", "__no_file__")

    left_col, right_col = st.columns([1.2, 3.8], gap="large")

    min_available_date = today
    max_available_date = today
    prepared_df: pd.DataFrame | None = None
    file_infos: list[dict[str, Any]] = []
    load_error: str | None = None

    with left_col:
        uploaded_files = st.file_uploader(
            "Settlement FINNET",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            help="Jika ada sheet 'Detail Settlement', sheet itu yang dipakai. Jika tidak ada, app memakai sheet pertama.",
        )

        if uploaded_files:
            try:
                prepared_df, file_infos = load_uploaded_files(uploaded_files)
                min_available_date = prepared_df["payment_date"].min()
                max_available_date = prepared_df["payment_date"].max()
            except Exception as exc:
                load_error = str(exc)

        current_file_token = get_file_token(uploaded_files)
        if current_file_token != st.session_state["loaded_file_token"]:
            st.session_state["loaded_file_token"] = current_file_token
            st.session_state["date_range_widget"] = (min_available_date, max_available_date)

        current_start, current_end = normalize_date_range(
            st.session_state["date_range_widget"],
            min_available_date,
            max_available_date,
        )
        current_start, current_end = clamp_date_range(
            current_start,
            current_end,
            min_available_date,
            max_available_date,
        )

        st.markdown("**Parameter Tanggal**")
        picked_range = st.date_input(
            "Rentang Payment Date Time",
            value=(current_start, current_end),
            min_value=min_available_date,
            max_value=max_available_date,
            format="DD/MM/YYYY",
            key="date_range_widget",
        )

        start_date, end_date = normalize_date_range(picked_range, current_start, current_end)
        start_date, end_date = clamp_date_range(start_date, end_date, min_available_date, max_available_date)

        process_clicked = st.button("Proses Rekonsiliasi", type="primary", use_container_width=True)

        st.caption(f"Jumlah file terpilih: {len(uploaded_files) if uploaded_files else 0}")
        if file_infos:
            sheet_info = ", ".join(sorted({str(item['selected_sheet']) for item in file_infos}))
            amount_info = ", ".join(sorted({str(item['merchant_amount_source']) for item in file_infos}))
            st.caption(f"Sheet terpilih: {sheet_info}")
            st.caption(f"Sumber Merchant Amount: {amount_info}")

    with right_col:
        if load_error:
            st.error(load_error)
            return

        if not uploaded_files:
            st.info("Upload file settlement FINNET untuk mulai proses rekonsiliasi.")
            return

        if not process_clicked:
            st.info("Pilih rentang tanggal terlebih dahulu, lalu klik **Proses Rekonsiliasi**.")
            return

        try:
            summary_df, grand_total_df, unmatched_df = build_branch_summary(prepared_df, start_date, end_date)
        except Exception as exc:
            st.error(str(exc))
            return

        st.markdown(
            f"**Periode Payment Date Time: {start_date.strftime('%d-%m-%Y')} s.d. {end_date.strftime('%d-%m-%Y')}**"
        )

        render_branch_sections(summary_df)

        st.markdown("### GRAND TOTAL")
        st.dataframe(format_table_for_display(grand_total_df), use_container_width=True, hide_index=True)

        if not unmatched_df.empty:
            st.warning("Ada Payment Method yang belum match ke master fee.")
            st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

        st.download_button(
            "Download Hasil Rekonsiliasi",
            data=to_excel_bytes(summary_df, grand_total_df, unmatched_df, file_infos, start_date, end_date),
            file_name=f"rekonsiliasi_sharing_fee_finnet_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
