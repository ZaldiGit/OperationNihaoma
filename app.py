
import base64
import html
import io
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib.parse import urlencode

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

st.set_page_config(page_title="Nihaoma Student Operations", layout="wide")

SCRIPT_URL = (
    st.secrets.get("SCRIPT_URL")
    or st.secrets.get("APPS_SCRIPT_URL")
    or os.getenv("SCRIPT_URL")
    or os.getenv("APPS_SCRIPT_URL")
    or ""
)
WRITE_TOKEN = st.secrets.get("WRITE_TOKEN", os.getenv("WRITE_TOKEN", ""))
TIMEOUT = 90

BASE_DIR = Path(__file__).parent
LOGO_PATH = BASE_DIR / "logo-nihaoma-rounded.png"
SIGNATURE_PATH = BASE_DIR / "signature-ttd.png"
STAMP_PATH = BASE_DIR / "stamp-cap.png"
APPROVAL_PATH = BASE_DIR / "approval-composite.png"

PROFILE_FIXED = {
    "Nama Brand": "Nihaoma Education Center",
    "Alamat ID": "Gedung Wirausaha Lantai 1,\nJalan HR Rasuna Said Kav. C-5,\nKelurahan Karet, Kecamatan Setia Budi,\nJakarta Selatan 12920",
    "Alamat CN": "佛城西路21号楼 1709-C, 江宁区, 南京市, 江苏, China.",
    "Alamat EN": "Fochengxilu 21 building, room 1709-C, Jiangning District,\nNanjing, Jiangsu, China.",
    "Email": "admin@nihaoma-education.com",
    "Telepon / WA": "+62 812-0000-0000",
    "Info Pembayaran": "Bank BCA - 1234567890 a/n Nihaoma Education Center",
    "Catatan Footer": "Terima kasih atas kepercayaan Anda. Invoice ini diterbitkan untuk kebutuhan administrasi program pendidikan ke China.",
}


def ensure_config() -> None:
    if not SCRIPT_URL or not WRITE_TOKEN:
        st.error("SCRIPT_URL atau WRITE_TOKEN belum diisi di secrets / environment.")
        st.stop()


def api_get(action: str, extra_params: Dict[str, Any] | None = None) -> Dict[str, Any]:
    ensure_config()
    params = {"action": action, "token": WRITE_TOKEN}
    if extra_params:
        params.update(extra_params)
    resp = requests.get(SCRIPT_URL, params=params, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


def api_post(action: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    ensure_config()
    body = {"action": action, "token": WRITE_TOKEN}
    body.update(payload)
    resp = requests.post(SCRIPT_URL, json=body, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


@st.cache_data(ttl=60, show_spinner=False)
def load_bootstrap() -> Dict[str, Any]:
    result = api_get("bootstrap")
    if not result.get("ok"):
        raise RuntimeError(result.get("error", "Gagal memuat data awal"))
    return result


def clear_cache_and_rerun() -> None:
    st.cache_data.clear()
    st.rerun()


def as_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value)


def to_number(value: Any) -> float:
    try:
        if value in (None, ""):
            return 0.0
        if isinstance(value, str):
            value = value.replace("Rp", "").replace(".", "").replace(",", ".").strip()
        return float(value)
    except Exception:
        return 0.0


def format_currency(value: Any) -> str:
    return f"Rp {to_number(value):,.0f}".replace(",", ".")


def option_index(options: List[str], value: Any) -> int:
    value = safe_text(value)
    try:
        return options.index(value)
    except ValueError:
        return 0


def ensure_option_list(base_options: List[Any], default_value: Any = "") -> List[str]:
    options = [safe_text(x).strip() for x in (base_options or []) if safe_text(x).strip()]
    default_text = safe_text(default_value).strip()
    if default_text and default_text not in options:
        options = [default_text] + options
    if not options:
        options = [default_text or "-"]
    return options


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].fillna("")
    return out


def maybe_date(value: Any) -> str:
    text = safe_text(value)
    if not text:
        return ""
    try:
        return str(pd.to_datetime(text).date())
    except Exception:
        return text


def format_date_id(value) -> str:
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return "-"
    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return safe_text(value) or "-"
    return ts.strftime("%d/%m/%Y")


def safe_paragraph_text(value) -> str:
    text = "" if value is None else str(value)
    return html.escape(text).replace("\n", "<br/>")


def find_student(students_df: pd.DataFrame, student_id: str) -> Dict[str, Any]:
    if students_df.empty or "student_id" not in students_df.columns:
        return {}
    row_df = students_df[students_df["student_id"].astype(str) == str(student_id)]
    return row_df.iloc[0].to_dict() if not row_df.empty else {}


def normalize_program_name(program: Any) -> str:
    return safe_text(program).strip().lower()


def get_registration_fee(program: Any) -> float:
    name = normalize_program_name(program)
    return 2_000_000.0 if "program bahasa" in name else 3_000_000.0


def get_transport_fee() -> float:
    return 4_000_000.0


def calculate_invoice_package(student: Dict[str, Any]) -> Dict[str, Any]:
    program = safe_text(student.get("program_diminati"))
    base_program_fee = to_number(student.get("estimasi_biaya"))
    registration_fee = get_registration_fee(program)
    transport_fee = get_transport_fee()
    admin_core_fee = max(base_program_fee - registration_fee, 0.0)
    admin_invoice_total = admin_core_fee + transport_fee
    grand_total = registration_fee + admin_invoice_total
    return {
        "program": program,
        "base_program_fee": base_program_fee,
        "registration_fee": registration_fee,
        "admin_core_fee": admin_core_fee,
        "transport_fee": transport_fee,
        "admin_invoice_total": admin_invoice_total,
        "grand_total": grand_total,
    }


def group_student_finance(invoices_df: pd.DataFrame) -> pd.DataFrame:
    if invoices_df.empty:
        return pd.DataFrame()
    inv = invoices_df.copy()
    for col in ["harga_program", "sudah_dibayar", "sisa_tagihan"]:
        if col in inv.columns:
            inv[col] = inv[col].apply(to_number)
        else:
            inv[col] = 0.0
    if "invoice_type" not in inv.columns:
        inv["invoice_type"] = "Manual"
    grouped = (
        inv.groupby(["student_id", "nama_mahasiswa"], dropna=False)
        .agg(
            total_invoice=("invoice_id", "count"),
            total_tagihan=("harga_program", "sum"),
            total_dibayar=("sudah_dibayar", "sum"),
            total_outstanding=("sisa_tagihan", "sum"),
        )
        .reset_index()
    )
    grouped["status_keuangan"] = grouped["total_outstanding"].apply(
        lambda v: "Lunas" if to_number(v) <= 0 else "Outstanding"
    )
    return grouped.sort_values(["total_outstanding", "nama_mahasiswa"], ascending=[False, True])


def register_fonts():
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass


def generate_invoice_pdf(record: dict, profile: dict) -> bytes:
    register_fonts()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=14 * mm,
        leftMargin=14 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    brand_name = ParagraphStyle(
        "BrandName",
        parent=styles["Title"],
        fontSize=20,
        leading=23,
        textColor=colors.HexColor("#1F2937"),
    )
    invoice_title = ParagraphStyle(
        "InvoiceTitle",
        parent=styles["Title"],
        fontSize=23,
        leading=26,
        textColor=colors.HexColor("#D97706"),
        alignment=1,
    )
    body = ParagraphStyle(
        "Body",
        parent=styles["BodyText"],
        fontSize=9.2,
        leading=12,
        textColor=colors.HexColor("#1F2937"),
    )
    chinese = ParagraphStyle("Chinese", parent=body, fontName="STSong-Light")
    label = ParagraphStyle("Label", parent=body, fontSize=9, leading=12)
    value = ParagraphStyle("Value", parent=body, fontSize=9.4, leading=12)

    BLUE = colors.HexColor("#1E88E5")
    BLUE_DARK = colors.HexColor("#0F4C81")
    ORANGE = colors.HexColor("#F59E0B")
    ORANGE_DARK = colors.HexColor("#D97706")
    LIGHT_ORANGE = colors.HexColor("#FFF6E8")

    story = []

    accent = Table([["", "", ""]], colWidths=[112 * mm, 40 * mm, 28 * mm], rowHeights=[4 * mm])
    accent.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), BLUE_DARK),
                ("BACKGROUND", (1, 0), (1, 0), BLUE),
                ("BACKGROUND", (2, 0), (2, 0), ORANGE),
            ]
        )
    )
    story.append(accent)
    story.append(Spacer(1, 6))

    logo_flowable = RLImage(str(LOGO_PATH), width=28 * mm, height=28 * mm) if LOGO_PATH.exists() else Paragraph("", body)
    identity_table = Table(
        [
            [logo_flowable, Paragraph(f"<b>{safe_paragraph_text(profile['Nama Brand'])}</b>", brand_name)],
            ["", Paragraph(safe_paragraph_text(profile["Alamat ID"]), body)],
            ["", Paragraph(safe_paragraph_text(profile["Alamat CN"]), chinese)],
            ["", Paragraph(safe_paragraph_text(profile["Alamat EN"]), body)],
            ["", Paragraph(f"Email: {safe_paragraph_text(profile['Email'])}<br/>WhatsApp: {safe_paragraph_text(profile['Telepon / WA'])}", body)],
        ],
        colWidths=[32 * mm, 78 * mm],
    )
    identity_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("VALIGN", (0, 0), (0, 0), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("SPAN", (0, 1), (0, 4)),
            ]
        )
    )

    invoice_card = Table(
        [
            [Paragraph("INVOICE", invoice_title)],
            [Paragraph(f"<b>Kode</b>: {safe_paragraph_text(record.get('Kode Invoice', '-'))}", body)],
            [Paragraph(f"<b>Tanggal</b>: {format_date_id(record.get('Tanggal Input'))}", body)],
            [Paragraph(f"<b>Status</b>: {safe_paragraph_text(record.get('Status Pelunasan', '-'))}", body)],
        ],
        colWidths=[64 * mm],
    )
    invoice_card.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), LIGHT_ORANGE),
                ("BOX", (0, 0), (-1, -1), 0.9, ORANGE),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )

    top = Table([[identity_table, invoice_card]], colWidths=[112 * mm, 66 * mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(top)
    story.append(Spacer(1, 8))

    def section_title(text, color):
        return Table(
            [[Paragraph(f"<font color='white'><b>{safe_paragraph_text(text)}</b></font>", styles["BodyText"])]],
            colWidths=[178 * mm],
            rowHeights=[8 * mm],
            style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), color), ("LEFTPADDING", (0, 0), (-1, -1), 8)]),
        )

    story.append(section_title("Student & Invoice Details", BLUE))
    detail_rows = [
        [Paragraph("<b>Nama Student</b>", label), Paragraph(safe_paragraph_text(record.get("Nama Student", "-")), value), Paragraph("<b>Program</b>", label), Paragraph(safe_paragraph_text(record.get("Program", "-")), value)],
        [Paragraph("<b>Passport / ID</b>", label), Paragraph(safe_paragraph_text(record.get("No. Paspor / ID", "-")), value), Paragraph("<b>Status Pelunasan</b>", label), Paragraph(safe_paragraph_text(record.get("Status Pelunasan", "-")), value)],
        [Paragraph("<b>Email</b>", label), Paragraph(safe_paragraph_text(record.get("Email Student", "-")), value), Paragraph("<b>Status Pengiriman</b>", label), Paragraph(safe_paragraph_text(record.get("Status Pengiriman", "-")), value)],
        [Paragraph("<b>WhatsApp</b>", label), Paragraph(safe_paragraph_text(record.get("No. WhatsApp", "-")), value), Paragraph("<b>Tanggal Invoice</b>", label), Paragraph(format_date_id(record.get("Tanggal Input")), value)],
        [Paragraph("<b>Intake / Catatan</b>", label), Paragraph(safe_paragraph_text(record.get("Intake / Keterangan", "-") or "-"), value), Paragraph("<b>Kode Invoice</b>", label), Paragraph(safe_paragraph_text(record.get("Kode Invoice", "-")), value)],
    ]
    detail_table = Table(detail_rows, colWidths=[30 * mm, 58 * mm, 33 * mm, 57 * mm])
    detail_table.setStyle(
        TableStyle(
            [
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, colors.HexColor("#EEF6FF")]),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#D6DCE5")),
                ("BOX", (0, 0), (-1, -1), 0.9, BLUE),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(detail_table)
    story.append(Spacer(1, 8))

    story.append(section_title("Program Charge", ORANGE))
    charge_rows = [
        [Paragraph("<font color='white'><b>No</b></font>", label), Paragraph("<font color='white'><b>Deskripsi Program</b></font>", label), Paragraph("<font color='white'><b>Mata Uang</b></font>", label), Paragraph("<font color='white'><b>Total</b></font>", label)],
        [Paragraph("1", value), Paragraph(safe_paragraph_text(f"Biaya program {record.get('Program', '-')}" + ("\nKeterangan: " + str(record.get("Intake / Keterangan", "-") or "-"))), value), Paragraph("IDR", value), Paragraph(format_currency(record.get("Harga Program", 0)), value)],
    ]
    charge_table = Table(charge_rows, colWidths=[14 * mm, 110 * mm, 18 * mm, 36 * mm])
    charge_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), ORANGE),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFF6E8")),
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#D6DCE5")),
                ("BOX", (0, 0), (-1, -1), 0.9, ORANGE_DARK),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(charge_table)
    story.append(Spacer(1, 8))

    story.append(section_title("Pembayaran & Ringkasan", colors.HexColor("#0F4C81")))
    summary_text = (
        f"<b>Total Program</b>: {format_currency(record.get('Harga Program', 0))}<br/>"
        f"<b>Sudah Dibayar</b>: {format_currency(record.get('Sudah Dibayar', 0))}<br/>"
        f"<b>Sisa Tagihan</b>: {format_currency(record.get('Sisa Tagihan', 0))}"
    )
    payment_table = Table(
        [[Paragraph(safe_paragraph_text(profile.get("Info Pembayaran", "-")), value), Paragraph(summary_text, value)]],
        colWidths=[112 * mm, 66 * mm],
    )
    payment_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), colors.HexColor("#EEF6FF")),
                ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#FFF6E8")),
                ("BOX", (0, 0), (-1, -1), 0.9, colors.HexColor("#0F4C81")),
                ("INNERGRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#D6DCE5")),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(payment_table)
    story.append(Spacer(1, 8))

    signature_title = ParagraphStyle("SignatureTitle", parent=body, fontSize=8.8, leading=11, textColor=colors.HexColor("#6B7280"))
    signature_name = ParagraphStyle("SignatureName", parent=body, fontSize=10, leading=12, textColor=colors.HexColor("#1F2937"), alignment=1)
    signature_role = ParagraphStyle("SignatureRole", parent=body, fontSize=9, leading=11, textColor=colors.HexColor("#6B7280"), alignment=1)

    footer_note = Table([[Paragraph(safe_paragraph_text(profile.get("Catatan Footer", "")), body)]], colWidths=[118 * mm])
    footer_note.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF7E6")),
                ("BOX", (0, 0), (-1, -1), 0.9, colors.HexColor("#F59E0B")),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    approval_img = RLImage(str(APPROVAL_PATH), width=56 * mm, height=20 * mm) if APPROVAL_PATH.exists() else Paragraph("", body)

    approval_block = Table(
        [
            [Paragraph("<b>Authorized Signature</b>", signature_title)],
            [approval_img],
            [Paragraph("<b>Yenny Pricila</b>", signature_name)],
            [Paragraph("Management", signature_role)],
        ],
        colWidths=[60 * mm],
    )
    approval_block.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("LINEABOVE", (0, 0), (-1, 0), 0.4, colors.HexColor("#D6DCE5")),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                ("ALIGN", (0, 1), (0, 1), "CENTER"),
            ]
        )
    )

    footer_combo = Table([[footer_note, approval_block]], colWidths=[118 * mm, 60 * mm])
    footer_combo.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(footer_combo)

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


def build_preview_invoice_url(invoice_id: str) -> str:
    ensure_config()
    params = urlencode({"action": "preview_invoice", "token": WRITE_TOKEN, "invoice_id": invoice_id})
    return f"{SCRIPT_URL}?{params}"


def expected_invoice_code(year_value: Any, student_id_value: Any) -> str:
    year = safe_text(year_value)
    year2 = year[2:4] if re.match(r"^\d{4}", year) else ""
    student_text = safe_text(student_id_value)
    match = re.search(r"(\d+)$", student_text)
    seq = match.group(1).zfill(2)[-2:] if match else "00"
    return f"NHEC-{year2}{seq}" if year2 else ""


def render_dashboard(students_df: pd.DataFrame, invoices_df: pd.DataFrame, payments_df: pd.DataFrame) -> None:
    st.subheader("Dashboard")

    active_students = students_df.copy()
    if not active_students.empty and "is_active" in active_students.columns:
        active_students = active_students[
            active_students["is_active"].astype(str).str.upper().isin(["TRUE", "1", "YA", "YES", ""])
        ].copy()

    inv = invoices_df.copy()
    if not inv.empty:
        inv["harga_program"] = inv.get("harga_program", 0).apply(to_number)
        inv["sudah_dibayar"] = inv.get("sudah_dibayar", 0).apply(to_number)
        inv["sisa_tagihan"] = inv.get("sisa_tagihan", 0).apply(to_number)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Calon Mahasiswa", len(active_students))
    c2.metric("Total Invoice", len(inv))
    c3.metric("Nilai Invoice", format_currency(inv["harga_program"].sum() if not inv.empty else 0))
    c4.metric("Sudah Dibayar", format_currency(inv["sudah_dibayar"].sum() if not inv.empty else 0))
    c5.metric("Outstanding", format_currency(inv["sisa_tagihan"].sum() if not inv.empty else 0))

    if not inv.empty:
        st.dataframe(inv, use_container_width=True, hide_index=True)


def render_edit_form(student: Dict[str, Any], refs: Dict[str, Any]) -> None:
    program_options = refs.get("program_diminati", refs.get("program", [])) or [safe_text(student.get("program_diminati"))]
    status_options = refs.get("status_proses", []) or [safe_text(student.get("status_proses"))]
    pic_options = refs.get("pic_admin", []) or [safe_text(student.get("pic_admin"))]
    intake_options = refs.get("intake", []) or [safe_text(student.get("intake"))]
    gender_options = refs.get("jenis_kelamin", []) or [safe_text(student.get("jenis_kelamin"))]
    lead_options = refs.get("sumber_leads", []) or [safe_text(student.get("sumber_leads"))]
    priority_options = refs.get("prioritas", []) or [safe_text(student.get("prioritas"))]

    with st.form("form_edit_student"):
        col1, col2, col3 = st.columns(3)
        nama_lengkap = col1.text_input("Nama Lengkap", value=safe_text(student.get("nama_lengkap")))
        nama_panggilan = col2.text_input("Nama Panggilan", value=safe_text(student.get("nama_panggilan")))
        jenis_kelamin = col3.selectbox("Jenis Kelamin", gender_options, index=option_index(gender_options, student.get("jenis_kelamin")))
        col4, col5, col6 = st.columns(3)
        tanggal_lahir = col4.text_input("Tanggal Lahir", value=safe_text(student.get("tanggal_lahir")))
        kewarganegaraan = col5.text_input("Kewarganegaraan", value=safe_text(student.get("kewarganegaraan")))
        no_whatsapp = col6.text_input("No WhatsApp", value=safe_text(student.get("no_whatsapp")))
        col7, col8, col9 = st.columns(3)
        email = col7.text_input("Email", value=safe_text(student.get("email")))
        no_paspor_atau_nik = col8.text_input("No Paspor / NIK", value=safe_text(student.get("no_paspor_atau_nik")))
        intake = col9.selectbox("Intake", intake_options, index=option_index(intake_options, student.get("intake")))
        col10, col11, col12 = st.columns(3)
        program_diminati = col10.selectbox("Program", program_options, index=option_index(program_options, student.get("program_diminati")))
        kampus_tujuan = col11.text_input("Kampus Tujuan", value=safe_text(student.get("kampus_tujuan")))
        kota_tujuan = col12.text_input("Kota Tujuan", value=safe_text(student.get("kota_tujuan")))
        col13, col14, col15 = st.columns(3)
        negara_tujuan = col13.text_input("Negara Tujuan", value=safe_text(student.get("negara_tujuan")))
        pic_admin = col14.selectbox("PIC", pic_options, index=option_index(pic_options, student.get("pic_admin")))
        status_proses = col15.selectbox("Status Proses", status_options, index=option_index(status_options, student.get("status_proses")))
        col16, col17, col18 = st.columns(3)
        sumber_leads = col16.selectbox("Sumber Leads", lead_options, index=option_index(lead_options, student.get("sumber_leads")))
        prioritas = col17.selectbox("Prioritas", priority_options, index=option_index(priority_options, student.get("prioritas")))
        next_action = col18.text_input("Next Action", value=safe_text(student.get("next_action")))
        alamat = st.text_area("Alamat", value=safe_text(student.get("alamat")))
        catatan_admin = st.text_area("Catatan Admin", value=safe_text(student.get("catatan_admin")))
        catatan_progress = st.text_input("Catatan log progress", value="Update dari form edit")

        if st.form_submit_button("Simpan Perubahan"):
            result = api_post(
                "update_student",
                {
                    "student_id": safe_text(student.get("student_id")),
                    "updated_by": pic_admin or "Admin",
                    "catatan_progress": catatan_progress,
                    "payload": {
                        "nama_lengkap": nama_lengkap,
                        "nama_panggilan": nama_panggilan,
                        "jenis_kelamin": jenis_kelamin,
                        "tanggal_lahir": tanggal_lahir,
                        "kewarganegaraan": kewarganegaraan,
                        "no_whatsapp": no_whatsapp,
                        "email": email,
                        "alamat": alamat,
                        "no_paspor_atau_nik": no_paspor_atau_nik,
                        "program_diminati": program_diminati,
                        "kampus_tujuan": kampus_tujuan,
                        "kota_tujuan": kota_tujuan,
                        "negara_tujuan": negara_tujuan,
                        "intake": intake,
                        "pic_admin": pic_admin,
                        "status_proses": status_proses,
                        "sumber_leads": sumber_leads,
                        "prioritas": prioritas,
                        "next_action": next_action,
                        "catatan_admin": catatan_admin,
                    },
                },
            )
            if result.get("ok"):
                st.success("Data mahasiswa berhasil diperbarui.")
                clear_cache_and_rerun()
            else:
                st.error(result.get("error", "Gagal update mahasiswa"))


def render_add_form(refs: Dict[str, Any]) -> None:
    st.markdown("### Tambah Data Mahasiswa")
    program_options = refs.get("program_diminati", refs.get("program", []))
    status_options = refs.get("status_proses", [])
    pic_options = refs.get("pic_admin", [])
    intake_options = refs.get("intake", [])
    gender_options = refs.get("jenis_kelamin", [])
    lead_options = refs.get("sumber_leads", [])
    priority_options = refs.get("prioritas", [])

    with st.form("form_add_student"):
        col1, col2, col3 = st.columns(3)
        nama_lengkap = col1.text_input("Nama Lengkap")
        nama_panggilan = col2.text_input("Nama Panggilan")
        jenis_kelamin = col3.selectbox("Jenis Kelamin", [""] + gender_options)
        col4, col5, col6 = st.columns(3)
        tanggal_lahir = col4.text_input("Tanggal Lahir")
        kewarganegaraan = col5.text_input("Kewarganegaraan", value="Indonesia")
        no_whatsapp = col6.text_input("No WhatsApp")
        col7, col8, col9 = st.columns(3)
        email = col7.text_input("Email")
        no_paspor_atau_nik = col8.text_input("No Paspor / NIK")
        intake = col9.selectbox("Intake", [""] + intake_options)
        col10, col11, col12 = st.columns(3)
        program_diminati = col10.selectbox("Program", [""] + program_options)
        kampus_tujuan = col11.text_input("Kampus Tujuan")
        kota_tujuan = col12.text_input("Kota Tujuan")
        col13, col14, col15 = st.columns(3)
        negara_tujuan = col13.text_input("Negara Tujuan", value="China")
        pic_admin = col14.selectbox("PIC", [""] + pic_options)
        status_proses = col15.selectbox("Status Proses", status_options, index=0 if status_options else None)
        col16, col17 = st.columns(2)
        sumber_leads = col16.selectbox("Sumber Leads", [""] + lead_options)
        prioritas = col17.selectbox("Prioritas", [""] + priority_options)
        alamat = st.text_area("Alamat")
        catatan_admin = st.text_area("Catatan Admin")

        if st.form_submit_button("Tambah Mahasiswa"):
            result = api_post(
                "add_student",
                {
                    "nama_lengkap": nama_lengkap,
                    "nama_panggilan": nama_panggilan,
                    "jenis_kelamin": jenis_kelamin,
                    "tanggal_lahir": tanggal_lahir,
                    "kewarganegaraan": kewarganegaraan,
                    "no_whatsapp": no_whatsapp,
                    "email": email,
                    "alamat": alamat,
                    "no_paspor_atau_nik": no_paspor_atau_nik,
                    "program_diminati": program_diminati,
                    "kampus_tujuan": kampus_tujuan,
                    "kota_tujuan": kota_tujuan,
                    "negara_tujuan": negara_tujuan,
                    "intake": intake,
                    "sumber_leads": sumber_leads,
                    "pic_admin": pic_admin,
                    "status_proses": status_proses or "New Lead",
                    "prioritas": prioritas or "Sedang",
                    "catatan_admin": catatan_admin,
                    "source": "streamlit",
                },
            )
            if result.get("ok"):
                st.success("Mahasiswa berhasil ditambahkan.")
                clear_cache_and_rerun()
            else:
                st.error(result.get("error", "Gagal menambah mahasiswa"))


def render_student_list(students_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Modul Calon Mahasiswa")
    tabs = st.tabs(["Daftar Mahasiswa", "Tambah Data", "Detail & Progress"])

    with tabs[0]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            st.dataframe(students_df, use_container_width=True, hide_index=True)

    with tabs[1]:
        render_add_form(refs)

    with tabs[2]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            detail_options = students_df["student_id"].astype(str).tolist()
            selected_detail_id = st.selectbox("Pilih mahasiswa", detail_options, key="detail_student_id")
            row_df = students_df[students_df["student_id"].astype(str) == str(selected_detail_id)]
            if not row_df.empty:
                student = row_df.iloc[0].to_dict()
                render_edit_form(student, refs)


def render_documents_module(students_df: pd.DataFrame, documents_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Dokumen")
    if documents_df.empty:
        st.info("Belum ada dokumen.")
    else:
        st.dataframe(documents_df, use_container_width=True, hide_index=True)


def invoice_row_for_pdf(invoice: Dict[str, Any], student: Dict[str, Any]) -> Dict[str, Any]:
    tanggal_input = safe_text(invoice.get("tanggal_invoice") or invoice.get("Tanggal Input"))
    intake_or_note = safe_text(student.get("intake") or invoice.get("catatan_invoice") or student.get("catatan_admin"))
    kode_invoice = safe_text(invoice.get("kode_invoice") or invoice.get("Kode Invoice"))
    expected_code = expected_invoice_code(tanggal_input, invoice.get("student_id"))
    return {
        "Kode Invoice": kode_invoice,
        "Tanggal Input": tanggal_input,
        "Nama Student": safe_text(invoice.get("nama_mahasiswa") or student.get("nama_lengkap")),
        "No. Paspor / ID": safe_text(student.get("no_paspor_atau_nik")),
        "Email Student": safe_text(student.get("email")),
        "No. WhatsApp": safe_text(student.get("no_whatsapp")),
        "Program": safe_text(invoice.get("program") or student.get("program_diminati")),
        "Harga Program": to_number(invoice.get("harga_program")),
        "Sudah Dibayar": to_number(invoice.get("sudah_dibayar")),
        "Sisa Tagihan": to_number(invoice.get("sisa_tagihan")),
        "Status Pelunasan": safe_text(invoice.get("status_pelunasan")),
        "Status Pengiriman": safe_text(invoice.get("status_pengiriman")),
        "Tanggal Kirim": safe_text(invoice.get("tanggal_kirim")),
        "Intake / Keterangan": intake_or_note,
        "_Expected Code": expected_code,
    }


def render_invoice_module(students_df: pd.DataFrame, invoices_df: pd.DataFrame, payments_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Invoice & Pembayaran")
    tabs = st.tabs(["Dashboard Invoice", "Buat Paket Invoice", "Buat Invoice Manual", "Record Pembayaran", "Preview / Save PDF"])

    inv = invoices_df.copy() if not invoices_df.empty else pd.DataFrame()
    if not inv.empty:
        for col in ["harga_program", "sudah_dibayar", "sisa_tagihan", "biaya_pendaftaran", "biaya_admin", "biaya_transport"]:
            if col in inv.columns:
                inv[col] = inv[col].apply(to_number)
            else:
                inv[col] = 0.0
        if "invoice_type" not in inv.columns:
            inv["invoice_type"] = "Manual"
        inv["invoice_type"] = inv["invoice_type"].replace("", "Manual")

    with tabs[0]:
        if inv.empty:
            st.info("Belum ada invoice.")
        else:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Invoice", len(inv))
            c2.metric("Total Nilai", format_currency(inv["harga_program"].sum()))
            c3.metric("Sudah Dibayar", format_currency(inv["sudah_dibayar"].sum()))
            c4.metric("Outstanding", format_currency(inv["sisa_tagihan"].sum()))

            finance_df = group_student_finance(inv)
            if not finance_df.empty:
                show_finance = finance_df.copy()
                for col in ["total_tagihan", "total_dibayar", "total_outstanding"]:
                    show_finance[col] = show_finance[col].apply(format_currency)
                st.dataframe(show_finance, use_container_width=True, hide_index=True)

            st.dataframe(inv, use_container_width=True, hide_index=True)

    with tabs[1]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            student_ids = students_df["student_id"].astype(str).tolist()
            selected_student_id = st.selectbox("Pilih mahasiswa untuk paket invoice", student_ids, key="invoice_package_student_id")
            student = find_student(students_df, selected_student_id)
            package = calculate_invoice_package(student)
            c1, c2, c3 = st.columns(3)
            c1.metric("Biaya pendaftaran", format_currency(package["registration_fee"]))
            c2.metric("Biaya admin", format_currency(package["admin_core_fee"]))
            c3.metric("Biaya transport", format_currency(package["transport_fee"]))

            expected_code = expected_invoice_code(datetime.now().strftime("%Y-%m-%d"), selected_student_id)
            if expected_code:
                st.caption(f"Format kode yang diinginkan: {expected_code}")

            with st.form("form_create_invoice_package"):
                t1, t2 = st.columns(2)
                tanggal_invoice_val = t1.date_input("Tanggal Invoice", value=datetime.now().date(), key="package_tanggal_invoice")
                mata_uang = t2.selectbox("Mata Uang", ["IDR", "USD", "CNY"], key="package_currency")
                status_pengiriman = st.selectbox("Status Pengiriman", refs.get("status_pengiriman", ["Belum Dikirim"]), key="package_status_pengiriman")
                kirim_hari_ini = st.checkbox("Sudah dikirim hari ini", key="package_sent_today")
                tanggal_kirim_val = st.date_input("Tanggal Kirim", value=datetime.now().date(), disabled=not kirim_hari_ini, key="package_tanggal_kirim")
                catatan_invoice = st.text_area("Catatan Invoice Paket", value="Invoice paket otomatis: Pendaftaran + Admin/Transport", key="package_catatan_invoice")
                if st.form_submit_button("Buat 2 invoice otomatis"):
                    result = api_post(
                        "create_invoice_package",
                        {
                            "student_id": selected_student_id,
                            "nama_mahasiswa": safe_text(student.get("nama_lengkap")),
                            "program": package["program"],
                            "tanggal_invoice": str(tanggal_invoice_val),
                            "mata_uang": mata_uang,
                            "status_pengiriman": status_pengiriman,
                            "tanggal_kirim": str(tanggal_kirim_val) if kirim_hari_ini else "",
                            "catatan_invoice": catatan_invoice,
                            "estimated_program_fee": package["base_program_fee"],
                        },
                    )
                    if result.get("ok"):
                        st.success("Paket invoice berhasil dibuat.")
                        st.info("Agar kode invoice mengikuti format NHEC-2603, update Code.gs juga dengan patch yang saya sertakan.")
                        clear_cache_and_rerun()
                    else:
                        st.error(result.get("error", "Gagal membuat paket invoice"))

    with tabs[2]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            student_ids = students_df["student_id"].astype(str).tolist()
            selected_student_id = st.selectbox("Pilih mahasiswa untuk invoice manual", student_ids, key="manual_invoice_student_id")
            student = find_student(students_df, selected_student_id)
            expected_code = expected_invoice_code(datetime.now().strftime("%Y-%m-%d"), selected_student_id)
            if expected_code:
                st.caption(f"Format kode yang diinginkan: {expected_code}")

            with st.form("form_create_invoice"):
                c1, c2, c3 = st.columns(3)
                tanggal_invoice_val = c1.date_input("Tanggal Invoice", value=datetime.now().date(), key="manual_tanggal_invoice")
                invoice_type = c2.selectbox("Jenis Invoice", ["Pendaftaran", "Admin", "Manual"], key="manual_invoice_type")
                mata_uang = c3.selectbox("Mata Uang", ["IDR", "USD", "CNY"], key="manual_currency")
                program_options = ensure_option_list(refs.get("program_diminati", refs.get("program", [])), student.get("program_diminati"))
                program = st.selectbox("Program", program_options, index=option_index(program_options, student.get("program_diminati")), key="manual_program")

                default_nominal = 0.0
                if invoice_type == "Pendaftaran":
                    default_nominal = get_registration_fee(program)
                elif invoice_type == "Admin":
                    package_preview = calculate_invoice_package({"program_diminati": program, "estimasi_biaya": student.get("estimasi_biaya")})
                    default_nominal = package_preview["admin_invoice_total"]

                harga_program = st.number_input("Nominal Invoice", min_value=0.0, value=float(default_nominal), step=100000.0, key="manual_harga_program")
                deskripsi_biaya = st.text_area("Deskripsi Biaya", value="", key="manual_deskripsi_biaya")
                status_pengiriman = st.selectbox("Status Pengiriman", refs.get("status_pengiriman", ["Belum Dikirim"]), key="manual_status_pengiriman")
                kirim_hari_ini = st.checkbox("Sudah dikirim hari ini", key="manual_sent_today")
                tanggal_kirim_val = st.date_input("Tanggal Kirim", value=datetime.now().date(), disabled=not kirim_hari_ini, key="manual_tanggal_kirim")
                catatan_invoice = st.text_area("Catatan Invoice", key="manual_catatan_invoice")
                if st.form_submit_button("Buat Invoice Manual"):
                    result = api_post(
                        "create_invoice",
                        {
                            "student_id": selected_student_id,
                            "nama_mahasiswa": safe_text(student.get("nama_lengkap")),
                            "tanggal_invoice": str(tanggal_invoice_val),
                            "program": program,
                            "deskripsi_biaya": deskripsi_biaya,
                            "mata_uang": mata_uang,
                            "harga_program": harga_program,
                            "status_pengiriman": status_pengiriman,
                            "tanggal_kirim": str(tanggal_kirim_val) if kirim_hari_ini else "",
                            "catatan_invoice": catatan_invoice,
                            "invoice_type": invoice_type,
                        },
                    )
                    if result.get("ok"):
                        st.success("Invoice berhasil dibuat.")
                        st.info("Agar kode invoice mengikuti format NHEC-2603, update Code.gs juga dengan patch yang saya sertakan.")
                        clear_cache_and_rerun()
                    else:
                        st.error(result.get("error", "Gagal membuat invoice"))

    with tabs[3]:
        if inv.empty:
            st.info("Belum ada invoice.")
        else:
            invoice_options = [
                f"{safe_text(row.get('invoice_id'))} | {safe_text(row.get('kode_invoice'))} | {safe_text(row.get('invoice_type'))} | {safe_text(row.get('nama_mahasiswa'))}"
                for _, row in inv.iterrows()
            ]
            selected_label = st.selectbox("Pilih invoice", invoice_options, key="payment_invoice_label")
            selected_invoice_id = selected_label.split("|")[0].strip()
            invoice_row = inv[inv["invoice_id"].astype(str) == selected_invoice_id].iloc[0].to_dict()
            with st.form("form_record_payment"):
                c1, c2, c3 = st.columns(3)
                tanggal_pembayaran = c1.text_input("Tanggal Pembayaran", value=str(datetime.now().date()))
                jumlah_pembayaran = c2.number_input("Jumlah Pembayaran", min_value=0.0, value=float(to_number(invoice_row.get("sisa_tagihan"))), step=100000.0)
                metode_pembayaran = c3.selectbox("Metode Pembayaran", refs.get("metode_pembayaran", ["Transfer"]))
                bukti_pembayaran_link = st.text_input("Link Bukti Pembayaran")
                dicatat_oleh = st.text_input("Dicatat oleh", value="Finance")
                catatan = st.text_area("Catatan Pembayaran")
                if st.form_submit_button("Simpan Pembayaran"):
                    result = api_post(
                        "record_payment",
                        {
                            "invoice_id": selected_invoice_id,
                            "student_id": safe_text(invoice_row.get("student_id")),
                            "tanggal_pembayaran": tanggal_pembayaran,
                            "jumlah_pembayaran": jumlah_pembayaran,
                            "metode_pembayaran": metode_pembayaran,
                            "bukti_pembayaran_link": bukti_pembayaran_link,
                            "dicatat_oleh": dicatat_oleh,
                            "catatan": catatan,
                        },
                    )
                    if result.get("ok"):
                        st.success("Pembayaran berhasil dicatat.")
                        clear_cache_and_rerun()
                    else:
                        st.error(result.get("error", "Gagal mencatat pembayaran"))

    with tabs[4]:
        if inv.empty:
            st.info("Belum ada invoice.")
        else:
            invoice_options = [
                f"{safe_text(row.get('invoice_id'))} | {safe_text(row.get('kode_invoice'))} | {safe_text(row.get('invoice_type'))} | {safe_text(row.get('nama_mahasiswa'))}"
                for _, row in inv.iterrows()
            ]
            selected_label = st.selectbox("Pilih invoice untuk preview", invoice_options, key="styled_invoice_label")
            selected_invoice_id = selected_label.split("|")[0].strip()
            invoice = inv[inv["invoice_id"].astype(str) == selected_invoice_id].iloc[0].to_dict()
            student = find_student(students_df, safe_text(invoice.get("student_id")))
            preview_url = build_preview_invoice_url(selected_invoice_id)
            pdf_record = invoice_row_for_pdf(invoice, student)
            pdf_bytes = generate_invoice_pdf(pdf_record, PROFILE_FIXED)

            left, right = st.columns([1, 1])
            with left:
                st.markdown("### Preview Styled")
                st.link_button("Buka Preview Invoice", preview_url, use_container_width=True)
                st.download_button(
                    "Download PDF Invoice",
                    data=pdf_bytes,
                    file_name=f"{safe_text(invoice.get('kode_invoice') or invoice.get('invoice_id'))}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
                st.info("Untuk hasil paling rapi, Anda tetap bisa buka preview lalu tekan Print > Save as PDF.")
                st.caption("Mac: Cmd+P | Windows: Ctrl+P")
            with right:
                st.markdown("### Ringkasan Invoice")
                st.write(f"**Kode Invoice saat ini:** {safe_text(invoice.get('kode_invoice'))}")
                if pdf_record.get("_Expected Code"):
                    st.write(f"**Format kode yang diinginkan:** {pdf_record['_Expected Code']}")
                st.write(f"**Jenis Invoice:** {safe_text(invoice.get('invoice_type'))}")
                st.write(f"**Nama Mahasiswa:** {safe_text(invoice.get('nama_mahasiswa'))}")
                st.write(f"**Program:** {safe_text(invoice.get('program'))}")
                st.write(f"**Harga Invoice:** {format_currency(invoice.get('harga_program'))}")
                st.write(f"**Sudah Dibayar:** {format_currency(invoice.get('sudah_dibayar'))}")
                st.write(f"**Sisa Tagihan:** {format_currency(invoice.get('sisa_tagihan'))}")
                st.write(f"**Status Pelunasan:** {safe_text(invoice.get('status_pelunasan'))}")


def render_help_module() -> None:
    st.subheader("Bantuan & SOP")
    st.markdown(
        """
        ### Cara pakai invoice yang rapi
        - Buka menu **Invoice & Pembayaran**
        - Pilih tab **Preview / Save PDF**
        - Klik **Buka Preview Invoice** untuk lihat versi browser
        - Klik **Download PDF Invoice** untuk unduh versi PDF dari aplikasi
        - Untuk hasil paling mirip preview, tetap bisa pakai **Print > Save as PDF**

        Catatan:
        - Format kode invoice yang Anda minta (`NHEC-2603`) perlu diterapkan di Apps Script, karena kode invoice dibuat di server.
        - File app ini sudah menampilkan format kode yang diharapkan sebagai panduan.
        """
    )


def main() -> None:
    st.title("Nihaoma Student Operations")
    st.caption("Dashboard operasional calon mahasiswa yang terhubung ke Google Sheet live")

    try:
        data = load_bootstrap()
    except Exception as exc:
        st.error(f"Gagal memuat data awal: {exc}")
        st.stop()

    students_df = normalize_df(as_df(data.get("students", [])))
    documents_df = normalize_df(as_df(data.get("documents", [])))
    invoices_df = normalize_df(as_df(data.get("invoices", [])))
    payments_df = normalize_df(as_df(data.get("payments", [])))
    refs = data.get("references", {}) or {}

    with st.sidebar:
        st.markdown("### Menu")
        page = st.radio("", ["Dashboard", "Calon Mahasiswa", "Dokumen", "Invoice & Pembayaran", "Bantuan & SOP"], label_visibility="collapsed")
        if st.button("Refresh data", use_container_width=True):
            clear_cache_and_rerun()
        st.caption(f"Data terakhir dimuat: {safe_text(data.get('meta', {}).get('generated_at'))}")

    if page == "Dashboard":
        render_dashboard(students_df, invoices_df, payments_df)
    elif page == "Calon Mahasiswa":
        render_student_list(students_df, refs)
    elif page == "Dokumen":
        render_documents_module(students_df, documents_df, refs)
    elif page == "Invoice & Pembayaran":
        render_invoice_module(students_df, invoices_df, payments_df, refs)
    elif page == "Bantuan & SOP":
        render_help_module()


if __name__ == "__main__":
    main()
