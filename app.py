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
}

REQUIRED_COLUMN_PATTERNS = {
    "payment_method": [r"payment\s*method"],
    "pg_provider": [r"pg\s*provider"],
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


def parse_finnet_datetime(series: pd.Series) -> pd.Series:
    text_series = series.astype(str).str.strip()
    extracted = text_series.str.extract(
        r"^(?P<day>\d{1,2})/(?P<month>\d{1,2})/(?P<year>\d{4})\s+(?P<hour>\d{2}):(?P<minute>\d{2}):(?P<second>\d{2})$"
    )

    normalized = pd.Series(pd.NA, index=series.index, dtype="object")
    valid_mask = extracted.notna().all(axis=1)

    if valid_mask.any():
        normalized.loc[valid_mask] = (
            extracted.loc[valid_mask, "day"].str.zfill(2)
            + "/"
            + extracted.loc[valid_mask, "month"].str.zfill(2)
            + "/"
            + extracted.loc[valid_mask, "year"]
            + " "
            + extracted.loc[valid_mask, "hour"]
            + ":"
            + extracted.loc[valid_mask, "minute"]
            + ":"
            + extracted.loc[valid_mask, "second"]
        )

    return pd.to_datetime(normalized, format="%d/%m/%Y %H:%M:%S", errors="coerce")


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

    return pd.read_excel(uploaded_file, sheet_name=0)


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

    if len(df.columns) < 21:
        missing.extend(
            [
                "kolom C untuk Payment Date Time",
                "kolom U untuk Merchant Amount",
            ]
        )
    else:
        resolved["payment_datetime"] = df.columns[2]
        resolved["merchant_amount"] = df.columns[20]

    payment_method_col = detect_column_by_patterns(df, REQUIRED_COLUMN_PATTERNS["payment_method"])
    if payment_method_col:
        resolved["payment_method"] = payment_method_col
    else:
        missing.append("Payment Method")

    pg_provider_col = detect_column_by_patterns(df, REQUIRED_COLUMN_PATTERNS["pg_provider"])
    if pg_provider_col:
        resolved["pg_provider"] = pg_provider_col
    else:
        missing.append("PG Provider")

    if missing:
        raise ValueError(f"Kolom wajib tidak ditemukan: {', '.join(missing)}")

    return resolved


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
    return df.loc[provider_norm.eq("finnet")].copy()


def prepare_finnet_dates(df: pd.DataFrame, payment_datetime_col: str) -> pd.DataFrame:
    working = df.copy()
    working["payment_datetime"] = parse_finnet_datetime(working[payment_datetime_col])
    working["payment_date"] = working["payment_datetime"].dt.date
    return working.loc[working["payment_datetime"].notna()].copy()


def filter_by_selected_date(df: pd.DataFrame, selected_date: date) -> pd.DataFrame:
    return df.loc[df["payment_date"] == selected_date].copy()


def build_reconciliation(df: pd.DataFrame, column_map: dict[str, str]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    master = build_master_df()
    alias_map = build_alias_map(master)

    working = df.copy()
    working["payment_method_raw"] = working[column_map["payment_method"]].astype(str).str.strip()
    working["instrumen_pembayaran"] = working[column_map["payment_method"]].map(
        lambda value: resolve_instrument(value, master, alias_map)
    )
    working["merchant_amount"] = working[column_map["merchant_amount"]].map(parse_number)

    merged = working.merge(master, how="left", on="instrumen_pembayaran")
    merged["expected_sharing_fee"] = merged.apply(
        lambda row: calc_fee(row["merchant_amount"], row["sharing_fee"]),
        axis=1,
    )
    merged["expected_sharing_fee_excl_tax"] = merged.apply(
        lambda row: calc_fee(row["merchant_amount"], row["sharing_fee_excl_tax"]),
        axis=1,
    )
    merged["payment_method_display"] = merged["instrumen_pembayaran"].fillna(merged["payment_method_raw"])
    merged["trx_count"] = 1

    detail = merged[
        [
            "payment_date",
            "payment_datetime",
            "payment_method_raw",
            "payment_method_display",
            "instrumen_pembayaran",
            "merchant_amount",
            "sharing_fee_excl_tax",
            "expected_sharing_fee_excl_tax",
            "sharing_fee",
            "expected_sharing_fee",
        ]
    ].copy()
    detail = detail.sort_values(["payment_method_display", "payment_datetime"], na_position="last")
    detail = round_columns(
        detail,
        [
            "merchant_amount",
            "sharing_fee_excl_tax",
            "expected_sharing_fee_excl_tax",
            "sharing_fee",
            "expected_sharing_fee",
        ],
        decimals=2,
    )

    summary = (
        merged.groupby(["payment_date", "payment_method_display"], dropna=False, as_index=False)
        .agg(
            master_sharing_fee_excl_tax=("sharing_fee_excl_tax", "first"),
            trx_count=("trx_count", "sum"),
            total_merchant_amount=("merchant_amount", "sum"),
            total_expected_sharing_fee_excl_tax=("expected_sharing_fee_excl_tax", "sum"),
        )
        .sort_values(["payment_method_display"], na_position="last")
    )
    summary = round_columns(
        summary,
        [
            "master_sharing_fee_excl_tax",
            "total_merchant_amount",
            "total_expected_sharing_fee_excl_tax",
        ],
        decimals=2,
    )

    unmatched = (
        merged.loc[merged["instrumen_pembayaran"].isna(), ["payment_date", "payment_method_raw"]]
        .drop_duplicates()
        .sort_values(["payment_method_raw"], na_position="last")
        .rename(columns={"payment_method_raw": "payment_method"})
    )

    return detail, summary, unmatched


def build_display_table(summary_df: pd.DataFrame) -> pd.DataFrame:
    display = summary_df.copy()
    display = display.rename(
        columns={
            "payment_method_display": "Payment Method",
            "master_sharing_fee_excl_tax": "Master Sharing Fee Exclude Tax",
            "trx_count": "Jumlah Transaksi",
            "total_merchant_amount": "Total Merchant Amount",
            "total_expected_sharing_fee_excl_tax": "Total Sharing Fee Exclude Tax",
        }
    )
    return display[
        [
            "Payment Method",
            "Master Sharing Fee Exclude Tax",
            "Jumlah Transaksi",
            "Total Merchant Amount",
            "Total Sharing Fee Exclude Tax",
        ]
    ]


def to_excel_bytes(display_df: pd.DataFrame, detail_df: pd.DataFrame, unmatched_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        display_df.to_excel(writer, index=False, sheet_name="Rekonsiliasi Ringkas")
        detail_df.to_excel(writer, index=False, sheet_name="Detail")
        unmatched_df.to_excel(writer, index=False, sheet_name="Unmatched Method")

        for sheet_name, frame in {
            "Rekonsiliasi Ringkas": display_df,
            "Detail": detail_df,
            "Unmatched Method": unmatched_df,
        }.items():
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            for idx, col in enumerate(frame.columns, start=1):
                max_len = max([len(str(col))] + [len(str(v)) for v in frame[col].head(200).fillna("")])
                ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = min(max_len + 2, 30)
    return output.getvalue()


st.title("Rekonsiliasi Sharing Fee FINNET")
st.caption("Pilih 1 tanggal dari Payment Date Time kolom C, proses hanya PG Provider FINNET, dan baca Excel selalu dari sheet pertama.")

with st.expander("Lihat master fee embedded", expanded=False):
    st.dataframe(
        build_master_df().drop(columns=["instrument_key"])[
            ["instrumen_pembayaran", "sharing_fee_excl_tax", "sharing_fee", "admin_fee", "status_pg"]
        ],
        use_container_width=True,
        hide_index=True,
    )

uploaded_file = st.file_uploader("Upload settlement FINNET", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        source_df = read_uploaded_file(uploaded_file)
        column_map = resolve_required_columns(source_df)
    except Exception as exc:
        st.error(str(exc))
        st.stop()

    finnet_df = filter_finnet_rows(source_df, column_map["pg_provider"])
    if finnet_df.empty:
        st.error("Tidak ada data dengan PG Provider = FINNET.")
        st.stop()

    finnet_df = prepare_finnet_dates(finnet_df, column_map["payment_datetime"])
    if finnet_df.empty:
        st.error(
            f"Kolom C (`{column_map['payment_datetime']}`) tidak bisa diparse dengan format d/mm/yyyy hh:mm:ss atau dd/mm/yyyy hh:mm:ss pada data FINNET."
        )
        st.stop()

    available_dates = sorted(finnet_df["payment_date"].dropna().unique())
    selected_date = st.date_input(
        "Pilih tanggal Payment Date Time",
        value=max(available_dates),
        min_value=min(available_dates),
        max_value=max(available_dates),
    )

    filtered_df = filter_by_selected_date(finnet_df, selected_date)

    info1, info2, info3 = st.columns(3)
    info1.metric("Total row upload", f"{len(source_df):,}")
    info2.metric("Row PG Provider = FINNET", f"{len(finnet_df):,}")
    info3.metric("Row tanggal terpilih", f"{len(filtered_df):,}")

    st.caption(
        f"Excel dibaca dari sheet pertama. Tanggal diparse dari kolom C (`{column_map['payment_datetime']}`) "
        f"dengan format `d/mm/yyyy hh:mm:ss` atau `dd/mm/yyyy hh:mm:ss`. Amount diambil dari kolom U (`{column_map['merchant_amount']}`)."
    )

    with st.expander("Preview data upload sheet 1", expanded=False):
        st.dataframe(source_df.head(20), use_container_width=True, hide_index=True)

    if filtered_df.empty:
        st.warning("Tidak ada data FINNET pada tanggal yang dipilih.")
        st.stop()

    detail_df, summary_df, unmatched_df = build_reconciliation(filtered_df, column_map)
    display_df = build_display_table(summary_df)

    st.subheader(f"Tanggal Payment Date Time: {selected_date.strftime('%d-%m-%Y')}")
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Payment Method": st.column_config.TextColumn(width="medium"),
            "Master Sharing Fee Exclude Tax": st.column_config.NumberColumn(format="%.2f"),
            "Jumlah Transaksi": st.column_config.NumberColumn(format="%d"),
            "Total Merchant Amount": st.column_config.NumberColumn(format="%.2f"),
            "Total Sharing Fee Exclude Tax": st.column_config.NumberColumn(format="%.2f"),
        },
    )

    recap1, recap2, recap3 = st.columns(3)
    recap1.metric("Payment Method tampil", f"{len(display_df):,}")
    recap2.metric("Total Merchant Amount", f"{display_df['Total Merchant Amount'].fillna(0).sum():,.2f}")
    recap3.metric("Total Sharing Fee Exclude Tax", f"{display_df['Total Sharing Fee Exclude Tax'].fillna(0).sum():,.2f}")

    if not unmatched_df.empty:
        with st.expander("Payment Method belum match ke master", expanded=False):
            st.dataframe(unmatched_df, use_container_width=True, hide_index=True)

    excel_bytes = to_excel_bytes(display_df, detail_df, unmatched_df)
    st.download_button(
        label="Download hasil rekonsiliasi (.xlsx)",
        data=excel_bytes,
        file_name=f"rekonsiliasi_sharing_fee_finnet_{selected_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.markdown(
    """
    **Aturan yang dikunci**
    - Excel selalu dibaca dari **sheet 1 / sheet pertama**.
    - Tanggal diambil dari **kolom C**.
    - Format `Payment Date Time` bisa **d/mm/yyyy hh:mm:ss** atau **dd/mm/yyyy hh:mm:ss**.
    - Amount diambil otomatis dari **kolom U**.
    - Tabel hanya menampilkan `Payment Method` yang ada pada data `PG Provider = FINNET`.
    """
)
