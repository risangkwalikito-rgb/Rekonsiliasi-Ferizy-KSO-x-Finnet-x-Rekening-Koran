# app.py
from __future__ import annotations

import io
import re
import unicodedata
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
}

ROLE_PATTERNS = {
    "payment_date": [r"payment\s*date", r"tanggal\s*bayar", r"paid\s*date", r"tgl\s*payment", r"date\s*payment"],
    "instrument": [r"instrumen", r"payment\s*channel", r"payment\s*method", r"channel", r"metode\s*pembayaran", r"issuer", r"bank", r"produk"],
    "amount": [r"transaction\s*amount", r"payment\s*amount", r"gross\s*amount", r"amount", r"nominal", r"bill\s*amount", r"paid\s*amount"],
    "actual_sharing_fee": [r"sharing\s*fee", r"merchant\s*fee", r"fee\s*share", r"revenue\s*share"],
    "status": [r"payment\s*status", r"transaction\s*status", r"status"],
    "trx_id": [r"transaction\s*id", r"trx\s*id", r"reference", r"payment\s*code", r"invoice", r"order\s*id", r"booking\s*code"],
}


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    text = text.replace("&", " dan ")
    text = re.sub(r"[_\-\/]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def clean_column_name(value: Any) -> str:
    return normalize_text(value)


def parse_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float, np.integer, np.floating)) and not pd.isna(value):
        return float(value)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "null", "-"}:
        return None

    text = text.replace("rp", "").replace("%", "").replace(" ", "")
    text = re.sub(r"[^\d,.\-]", "", text)

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
    text_series = series.astype(str)

    try:
        parsed = pd.to_datetime(text_series, format="mixed", errors="coerce", dayfirst=False)
    except TypeError:
        parsed = pd.to_datetime(text_series, errors="coerce")

    if parsed.notna().mean() >= 0.6:
        return parsed

    try:
        return pd.to_datetime(text_series, format="mixed", errors="coerce", dayfirst=True)
    except TypeError:
        return pd.to_datetime(text_series, errors="coerce", dayfirst=True)


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


def detect_columns(df: pd.DataFrame) -> dict[str, str | None]:
    result: dict[str, str | None] = {}
    cleaned = {col: clean_column_name(col) for col in df.columns}

    for role, patterns in ROLE_PATTERNS.items():
        best_col = None
        best_score = -1
        for col, normalized in cleaned.items():
            score = sum(2 for pattern in patterns if re.search(pattern, normalized))
            if role == "amount" and normalized == "amount":
                score += 1
            if role == "status" and normalized == "status":
                score += 1
            if score > best_score:
                best_score = score
                best_col = col
        result[role] = best_col if best_score > 0 else None
    return result


def build_alias_map(master_df: pd.DataFrame) -> dict[str, str]:
    alias_map = {normalize_text(k): v for k, v in MANUAL_ALIAS.items()}
    for instrument in master_df["instrumen_pembayaran"]:
        key = normalize_text(instrument)
        alias_map[key] = instrument
        compact = key.replace(" ", "")
        alias_map[compact] = instrument
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


def round_columns(df: pd.DataFrame, cols: list[str], decimals: int = 2) -> pd.DataFrame:
    formatted = df.copy()
    for col in cols:
        if col in formatted.columns:
            formatted[col] = pd.to_numeric(formatted[col], errors="coerce").round(decimals)
    return formatted


def build_reconciliation(
    df: pd.DataFrame,
    payment_date_col: str,
    instrument_col: str,
    amount_col: str,
    actual_sharing_fee_col: str | None = None,
    status_col: str | None = None,
    trx_id_col: str | None = None,
    allowed_statuses: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    master = build_master_df()
    alias_map = build_alias_map(master)

    working = df.copy()
    working["payment_date_raw"] = working[payment_date_col]
    working["payment_date"] = parse_datetime_series(working[payment_date_col]).dt.date
    working["instrument_raw"] = working[instrument_col]
    working["instrumen_pembayaran"] = working[instrument_col].map(lambda x: resolve_instrument(x, master, alias_map))
    working["amount"] = working[amount_col].map(parse_number)
    working["status_transaksi"] = working[status_col].astype(str) if status_col else None
    working["trx_id"] = working[trx_id_col].astype(str) if trx_id_col else None

    if actual_sharing_fee_col:
        working["actual_sharing_fee"] = working[actual_sharing_fee_col].map(parse_number)
    else:
        working["actual_sharing_fee"] = np.nan

    if status_col and allowed_statuses:
        normalized_allowed = {normalize_text(x) for x in allowed_statuses}
        mask = working[status_col].astype(str).map(normalize_text).isin(normalized_allowed)
        working = working.loc[mask].copy()

    merged = working.merge(master, how="left", on="instrumen_pembayaran")
    merged["expected_admin_fee"] = merged.apply(lambda row: calc_fee(row["amount"], row["admin_fee"]), axis=1)
    merged["expected_sharing_fee"] = merged.apply(lambda row: calc_fee(row["amount"], row["sharing_fee"]), axis=1)
    merged["expected_sharing_fee_excl_tax"] = merged.apply(
        lambda row: calc_fee(row["amount"], row["sharing_fee_excl_tax"]), axis=1
    )
    merged["delta_sharing_fee"] = merged["actual_sharing_fee"] - merged["expected_sharing_fee"]

    detail_columns = [
        "payment_date",
        "trx_id",
        "instrument_raw",
        "instrumen_pembayaran",
        "status_pg",
        "settlement_hari_kerja",
        "amount",
        "admin_fee",
        "sharing_fee",
        "pajak",
        "sharing_fee_excl_tax",
        "expected_admin_fee",
        "expected_sharing_fee",
        "expected_sharing_fee_excl_tax",
        "actual_sharing_fee",
        "delta_sharing_fee",
    ]
    if status_col:
        detail_columns.insert(2, "status_transaksi")

    detail = merged[[col for col in detail_columns if col in merged.columns]].copy()
    detail = detail.sort_values(["payment_date", "instrumen_pembayaran"], na_position="last")
    detail = round_columns(detail, ["amount", "expected_admin_fee", "expected_sharing_fee", "expected_sharing_fee_excl_tax", "actual_sharing_fee", "delta_sharing_fee"], decimals=2)
    detail = round_columns(detail, ["admin_fee", "sharing_fee", "sharing_fee_excl_tax", "pajak"], decimals=6)

    summary_aggs: dict[str, Any] = {
        "amount": "sum",
        "expected_admin_fee": "sum",
        "expected_sharing_fee": "sum",
        "expected_sharing_fee_excl_tax": "sum",
        "trx_count": "sum",
    }
    if actual_sharing_fee_col:
        summary_aggs["actual_sharing_fee"] = "sum"
        summary_aggs["delta_sharing_fee"] = "sum"

    merged["trx_count"] = 1
    summary = (
        merged.groupby(["payment_date", "instrumen_pembayaran"], dropna=False, as_index=False)
        .agg(summary_aggs)
        .sort_values(["payment_date", "instrumen_pembayaran"], na_position="last")
    )

    rename_map = {
        "amount": "total_amount",
        "expected_admin_fee": "total_expected_admin_fee",
        "expected_sharing_fee": "total_expected_sharing_fee",
        "expected_sharing_fee_excl_tax": "total_expected_sharing_fee_excl_tax",
        "actual_sharing_fee": "total_actual_sharing_fee",
        "delta_sharing_fee": "total_delta_sharing_fee",
    }
    summary = summary.rename(columns=rename_map)
    summary = round_columns(
        summary,
        [
            "total_amount",
            "total_expected_admin_fee",
            "total_expected_sharing_fee",
            "total_expected_sharing_fee_excl_tax",
            "total_actual_sharing_fee",
            "total_delta_sharing_fee",
        ],
        decimals=2,
    )

    unmatched = detail.loc[detail["instrumen_pembayaran"].isna(), ["payment_date", "instrument_raw"]].copy()
    unmatched = unmatched.drop_duplicates().sort_values(["payment_date", "instrument_raw"], na_position="last")

    return detail, summary, unmatched


def to_excel_bytes(detail: pd.DataFrame, summary: pd.DataFrame, unmatched: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail.to_excel(writer, index=False, sheet_name="Detail Sharing Fee")
        summary.to_excel(writer, index=False, sheet_name="Rekap Sharing Fee")
        unmatched.to_excel(writer, index=False, sheet_name="Unmatched Channel")

        for sheet_name, frame in {
            "Detail Sharing Fee": detail,
            "Rekap Sharing Fee": summary,
            "Unmatched Channel": unmatched,
        }.items():
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            for idx, col in enumerate(frame.columns, start=1):
                max_len = max([len(str(col))] + [len(str(v)) for v in frame[col].head(200).fillna("")])
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 24)
    return output.getvalue()


st.title("Rekonsiliasi Sharing Fee FINNET")
st.caption("Upload settlement FINNET, cocokkan ke master tarif embedded, lalu bentuk tabel detail dan rekap berdasarkan Payment Date.")

with st.expander("Lihat master tarif sharing fee", expanded=False):
    master_preview = build_master_df().drop(columns=["instrument_key"])
    st.dataframe(master_preview, use_container_width=True, hide_index=True)

uploaded_file = st.file_uploader(
    "Upload file settlement FINNET",
    type=["xlsx", "xls", "csv"],
    help="Gunakan export dashboard FINNET. Tabel hasil akan memakai tanggal dari kolom Payment Date.",
)

if uploaded_file:
    try:
        source_df = read_uploaded_file(uploaded_file)
    except Exception as exc:
        st.error(f"Gagal membaca file: {exc}")
        st.stop()

    if source_df.empty:
        st.warning("File kosong.")
        st.stop()

    st.subheader("Preview data settlement")
    st.dataframe(source_df.head(20), use_container_width=True)

    detected = detect_columns(source_df)
    all_columns = [None] + source_df.columns.tolist()

    st.subheader("Mapping kolom")
    col1, col2, col3 = st.columns(3)
    with col1:
        payment_date_col = st.selectbox("Kolom Payment Date", all_columns, index=all_columns.index(detected["payment_date"]) if detected["payment_date"] in all_columns else 0)
        instrument_col = st.selectbox("Kolom Instrumen/Channel", all_columns, index=all_columns.index(detected["instrument"]) if detected["instrument"] in all_columns else 0)
    with col2:
        amount_col = st.selectbox("Kolom Amount/Nominal", all_columns, index=all_columns.index(detected["amount"]) if detected["amount"] in all_columns else 0)
        actual_sharing_fee_col = st.selectbox("Kolom Actual Sharing Fee (opsional)", all_columns, index=all_columns.index(detected["actual_sharing_fee"]) if detected["actual_sharing_fee"] in all_columns else 0)
    with col3:
        status_col = st.selectbox("Kolom Status (opsional)", all_columns, index=all_columns.index(detected["status"]) if detected["status"] in all_columns else 0)
        trx_id_col = st.selectbox("Kolom TRX ID / Order ID (opsional)", all_columns, index=all_columns.index(detected["trx_id"]) if detected["trx_id"] in all_columns else 0)

    if not payment_date_col or not instrument_col or not amount_col:
        st.info("Pilih minimal kolom Payment Date, Instrumen/Channel, dan Amount.")
        st.stop()

    allowed_statuses: list[str] | None = None
    if status_col:
        status_values = (
            source_df[status_col]
            .dropna()
            .astype(str)
            .str.strip()
            .sort_values()
            .unique()
            .tolist()
        )
        default_values = [value for value in status_values if normalize_text(value) in {"success", "paid", "settled", "completed"}]
        allowed_statuses = st.multiselect(
            "Filter status yang dihitung",
            options=status_values,
            default=default_values or status_values,
            help="Kosongkan pilihan jika semua baris ingin dihitung.",
        )

    process = st.button("Proses rekonsiliasi", type="primary")

    if process:
        detail_df, summary_df, unmatched_df = build_reconciliation(
            df=source_df,
            payment_date_col=payment_date_col,
            instrument_col=instrument_col,
            amount_col=amount_col,
            actual_sharing_fee_col=actual_sharing_fee_col,
            status_col=status_col,
            trx_id_col=trx_id_col,
            allowed_statuses=allowed_statuses,
        )

        metric1, metric2, metric3, metric4 = st.columns(4)
        metric1.metric("Total transaksi", f"{len(detail_df):,}")
        metric2.metric("Total expected sharing fee", f"{detail_df['expected_sharing_fee'].fillna(0).sum():,.2f}")
        metric3.metric("Total expected sharing fee excl tax", f"{detail_df['expected_sharing_fee_excl_tax'].fillna(0).sum():,.2f}")
        metric4.metric("Channel belum match", f"{unmatched_df['instrument_raw'].nunique() if not unmatched_df.empty else 0:,}")

        tab1, tab2, tab3 = st.tabs(["Detail Sharing Fee", "Rekap Sharing Fee", "Unmatched Channel"])

        with tab1:
            st.dataframe(detail_df, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
        with tab3:
            if unmatched_df.empty:
                st.success("Semua channel berhasil dicocokkan ke master tarif.")
            else:
                st.warning("Ada channel yang belum match ke master tarif. Tambahkan alias di kode jika perlu.")
                st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

        export_bytes = to_excel_bytes(detail_df, summary_df, unmatched_df)
        st.download_button(
            label="Download hasil rekonsiliasi (.xlsx)",
            data=export_bytes,
            file_name="rekonsiliasi_sharing_fee_finnet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()
st.markdown(
    """
    **Catatan**
    - Tabel hasil menggunakan tanggal dari kolom **Payment Date**.
    - Master tarif sharing fee sudah ditanam langsung di dalam kode.
    - Saya tidak bisa membaca history chat lain untuk contoh ESPAY, jadi tabel detail di atas dibuat generik agar bisa dipakai untuk format serupa.
    """
)
