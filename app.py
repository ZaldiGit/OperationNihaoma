import base64
import html
import io
import json
import os
import re
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover
    ZoneInfo = None
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
HERO_STUDENT_PATH = BASE_DIR / "hero-student.png"

PROFILE_FIXED = {
    "Nama Brand": "Nihaoma Education Center",
    "Alamat ID": "Gedung Wirausaha Lantai 1,\nJalan HR Rasuna Said Kav. C-5,\nKelurahan Karet, Kecamatan Setia Budi,\nJakarta Selatan 12920",
    "Alamat CN": "佛城西路21号楼 1709-C, 江宁区, 南京市, 江苏, China.",
    "Alamat EN": "Fochengxilu 21 building, room 1709-C, Jiangning District,\nNanjing, Jiangsu, China.",
    "Email": "nihaoma.eduu@gmail.com",
    "Telepon / WA": "083137127808",
    "Info Pembayaran": "Bank BCA - 0068601889 a/n Nihaoma Education Center",
    "Catatan Footer": "Terima kasih atas kepercayaan Anda. Invoice ini diterbitkan untuk kebutuhan administrasi program pendidikan ke China.",
}
def inject_ui_style() -> None:
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap');

        html, body, [class*="css"], .stApp, .stMarkdown, .stText, .stCaption,
        div[data-testid="stMetric"], div[data-testid="stButton"] > button,
        div[data-testid="stSidebar"], input, textarea, select, label {
            font-family: 'Poppins', sans-serif !important;
        }

        h1, h2, h3, h4, h5, h6,
        .hero-title, .hero-subtitle, .section-title {
            font-family: 'Poppins', sans-serif !important;
        }
        .stApp {
            background: linear-gradient(180deg, #fffdf9 0%, #fff7ed 55%, #ffedd5 100%) !important;
        }

        .main .block-container {
            background: transparent !important;
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }

        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #b45309 0%, #f59e0b 100%) !important;
        }

        section[data-testid="stSidebar"] * {
            color: white !important;
        }

        section[data-testid="stSidebar"] img {
            border-radius: 18px;
            margin-bottom: 12px;
        }

        .sidebar-hero-caption {
            text-align: center;
            font-size: 13px;
            font-weight: 700;
            color: white;
            margin-top: -4px;
            margin-bottom: 10px;
        }

        .hero-box {
            background: linear-gradient(135deg, #ffe0a3 0%, #f59e0b 100%);
            border-radius: 28px;
            padding: 30px 34px;
            margin-bottom: 22px;
            box-shadow: 0 12px 28px rgba(245, 158, 11, 0.18);
        }

        .hero-title {
            font-size: 38px;
            font-weight: 800;
            color: #8a3b12;
            line-height: 1.15;
            margin-bottom: 8px;
        }

        .hero-subtitle {
            font-size: 18px;
            color: #9a4b1e;
            margin-bottom: 0;
        }

        .soft-card {
            background: rgba(255,255,255,0.94);
            border-radius: 22px;
            padding: 18px 20px;
            box-shadow: 0 10px 24px rgba(245, 158, 11, 0.08);
            border: 1px solid rgba(217, 119, 6, 0.10);
            margin-bottom: 14px;
        }

        .section-title {
            font-size: 24px;
            font-weight: 700;
            color: #b45309;
            margin: 10px 0 16px 0;
        }

        .notification-badge {
            display: inline-block;
            background: linear-gradient(135deg, #f97316 0%, #ea580c 100%);
            color: white;
            padding: 6px 12px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 700;
            box-shadow: 0 6px 14px rgba(249, 115, 22, 0.18);
        }

        div[data-testid="stMetric"] {
            background: rgba(255,255,255,0.95) !important;
            border: 1px solid rgba(217, 119, 6, 0.10);
            border-radius: 18px;
            padding: 12px;
            box-shadow: 0 8px 20px rgba(0,0,0,0.04);
        }

        div[data-testid="stMetricLabel"] {
            font-size: 14px !important;
            font-weight: 700 !important;
            line-height: 1.2 !important;
        }

        div[data-testid="stMetricValue"] {
            font-size: 28px !important;
            font-weight: 800 !important;
            line-height: 1.05 !important;
        }

        div[data-testid="stMetricValue"] > div {
            font-size: 28px !important;
            font-weight: 800 !important;
            white-space: normal !important;
            overflow: visible !important;
            text-overflow: unset !important;
        }

        div[data-testid="stButton"] > button {
            border-radius: 18px;
            border: 1px solid rgba(217, 119, 6, 0.14);
            background: white !important;
            min-height: 58px;
            font-weight: 700;
            font-size: 18px;
            box-shadow: 0 6px 16px rgba(0,0,0,0.04);
        }

        div[data-testid="stButton"] > button:hover {
            border-color: #d97706;
            color: #d97706;
        }

        div[data-testid="stPlotlyChart"] {
            background: linear-gradient(180deg, rgba(255,255,255,0.94) 0%, rgba(255,248,240,0.99) 100%);
            border: 1px solid rgba(217, 119, 6, 0.10);
            border-radius: 22px;
            padding: 14px 14px 8px 14px;
            box-shadow: 0 10px 24px rgba(0,0,0,0.04);
            margin-bottom: 12px;
            overflow: visible !important;
        }

        div[data-baseweb="select"] > div,
        .stTextInput input,
        .stTextArea textarea,
        .stDateInput input,
        .stNumberInput input {
            background: rgba(255,255,255,0.96) !important;
        }

        [data-testid="stDataFrame"] {
            background: rgba(255,255,255,0.96) !important;
            border-radius: 18px;
        }
    </style>
    """, unsafe_allow_html=True)
def render_top_header() -> None:
    st.markdown("""
    <div class="hero-box">
        <div class="hero-title"> Single Sign-On (SSO) Nihaoma Education Center</div>
        <div class="hero-subtitle">
            Dashboard Operasional Calon Mahasiswa (Klien) Nihaoma
        </div>
    </div>
    """, unsafe_allow_html=True)


# ---------- Core helpers ----------
def ensure_config() -> None:
    if not SCRIPT_URL or not WRITE_TOKEN:
        st.error("SCRIPT_URL atau WRITE_TOKEN belum diisi di secrets / environment.")
        st.stop()


def api_get(action: str, extra_params: Dict[str, Any] | None = None) -> Dict[str, Any]:
    ensure_config()
    params = {"action": action, "token": WRITE_TOKEN}
    if extra_params:
        params.update(extra_params)
    resp = requests.get(
        SCRIPT_URL,
        params=params,
        timeout=TIMEOUT,
    )
    resp.raise_for_status()
    return resp.json()


def api_post(action: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    ensure_config()
    body = {"action": action, "token": WRITE_TOKEN}
    body.update(payload)
    resp = requests.post(SCRIPT_URL, json=body, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()

def upload_invoice_pdf_to_drive(
    invoice_id: str,
    student_id: str,
    nama_mahasiswa: str,
    kode_invoice: str,
    invoice_type: str,
    pdf_bytes: bytes,
) -> Dict[str, Any]:
    file_base64 = base64.b64encode(pdf_bytes).decode("utf-8")
    return api_post(
        "upload_invoice_pdf_file",
        {
            "invoice_id": invoice_id,
            "student_id": student_id,
            "nama_mahasiswa": nama_mahasiswa,
            "kode_invoice": kode_invoice,
            "invoice_type": invoice_type,
            "mime_type": "application/pdf",
            "nama_file": invoice_pdf_filename(kode_invoice, nama_mahasiswa, invoice_type),
            "file_base64": file_base64,
        },
    )

@st.cache_data(ttl=60, show_spinner=False)
def load_bootstrap() -> Dict[str, Any]:
    result = api_get("bootstrap")
    if not result.get("ok"):
        raise RuntimeError(result.get("error", "Gagal memuat data awal"))
    return result


def clear_cache_and_rerun() -> None:
    st.cache_data.clear()
    st.rerun()
def go_to_page(page_name: str) -> None:
    st.session_state["pending_page"] = page_name
    st.rerun()

def detect_new_students(students_df: pd.DataFrame) -> List[Dict[str, Any]]:
    if students_df.empty or "student_id" not in students_df.columns:
        return []

    current_rows = []
    for _, row in students_df.iterrows():
        current_rows.append({
            "student_id": safe_text(row.get("student_id")),
            "nama_lengkap": safe_text(row.get("nama_lengkap")),
            "program_diminati": safe_text(row.get("program_diminati")),
            "tanggal_input": safe_text(row.get("tanggal_input")),
        })

    current_ids = {row["student_id"] for row in current_rows if row["student_id"]}

    if "seen_student_ids" not in st.session_state:
        st.session_state["seen_student_ids"] = current_ids
        st.session_state["latest_new_students"] = []
        return []

    seen_ids = set(st.session_state.get("seen_student_ids", set()))
    new_students = [row for row in current_rows if row["student_id"] and row["student_id"] not in seen_ids]

    st.session_state["seen_student_ids"] = current_ids
    st.session_state["latest_new_students"] = new_students
    return new_students


# ---------- Formatting helpers ----------
def as_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value)

def clean_filename_part(value: Any) -> str:
    text = safe_text(value).strip()
    text = re.sub(r"[^\w\s.-]", "", text, flags=re.UNICODE)
    text = re.sub(r"\s+", "-", text)
    return text.strip("-_.") or "Tanpa-Nama"


def get_student_short_name(student: Dict[str, Any]) -> str:
    return (
        safe_text(student.get("nama_panggilan")).strip()
        or safe_text(student.get("nama_lengkap")).strip()
        or safe_text(student.get("nama_mahasiswa")).strip()
    )


def student_code_name(student_id: Any, nama: Any) -> str:
    sid = safe_text(student_id).strip()
    name = safe_text(nama).strip()
    return f"{sid} - {name}" if sid and name else sid or name


def student_display_label(student: Dict[str, Any]) -> str:
    return student_code_name(
        student.get("student_id"),
        get_student_short_name(student),
    )


def build_student_options(students_df: pd.DataFrame) -> tuple[List[str], Dict[str, str]]:
    labels = []
    mapping = {}

    if students_df.empty:
        return labels, mapping

    for _, row in students_df.iterrows():
        student = row.to_dict()
        sid = safe_text(student.get("student_id"))
        label = student_display_label(student)
        labels.append(label)
        mapping[label] = sid

    return labels, mapping


def invoice_code_name(kode_invoice: Any, nama_mahasiswa: Any) -> str:
    code = safe_text(kode_invoice).strip()
    name = safe_text(nama_mahasiswa).strip()
    return f"{code} - {name}" if code and name else code or name


def invoice_display_label(invoice: Dict[str, Any]) -> str:
    label = invoice_code_name(
        invoice.get("kode_invoice") or invoice.get("invoice_id"),
        invoice.get("nama_mahasiswa"),
    )
    invoice_type = safe_text(invoice.get("invoice_type")).strip()
    return f"{label} ({invoice_type})" if invoice_type else label


def build_invoice_options(inv_df: pd.DataFrame) -> tuple[List[str], Dict[str, str]]:
    labels = []
    mapping = {}

    if inv_df.empty:
        return labels, mapping

    for _, row in inv_df.iterrows():
        invoice = row.to_dict()
        invoice_id = safe_text(invoice.get("invoice_id"))
        label = invoice_display_label(invoice)
        labels.append(label)
        mapping[label] = invoice_id

    return labels, mapping


def invoice_pdf_filename(kode_invoice: Any, nama_mahasiswa: Any, invoice_type: Any = "") -> str:
    parts = []

    code = safe_text(kode_invoice).strip()
    name = safe_text(nama_mahasiswa).strip()
    inv_type = safe_text(invoice_type).strip()

    if code:
        parts.append(clean_filename_part(code))
    if name:
        parts.append(clean_filename_part(name))
    if inv_type:
        parts.append(clean_filename_part(inv_type))

    if not parts:
        return "Invoice.pdf"

    return "-".join(parts) + ".pdf"


def document_filename(student_id: Any, nama_mahasiswa: Any, jenis_dokumen: Any, original_name: str) -> str:
    suffix = Path(original_name).suffix.lower()
    sid = clean_filename_part(student_id)
    name = clean_filename_part(nama_mahasiswa)
    doc_type = clean_filename_part(jenis_dokumen)
    return f"{sid}-{name}-{doc_type}{suffix}"


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


def get_program_total_fee(program: Any, fallback_value: Any = 0) -> float:
    name = normalize_program_name(program).replace(" ", "")

    # Harga khusus Program Bahasa
    if "programbahasa" in name:
        return 17_000_000.0

    # Harga khusus program S1-S3
    if "s1-s3" in name or "s1s3" in name:
        return 28_800_000.0

    return to_number(fallback_value)


def get_transport_fee() -> float:
    return 4_000_000.0


def calculate_invoice_package(student: Dict[str, Any]) -> Dict[str, Any]:
    program = safe_text(student.get("program_diminati"))
    base_program_fee = get_program_total_fee(program, student.get("estimasi_biaya"))
    registration_fee = get_registration_fee(program)

    # Transport sudah dianggap termasuk di estimasi_biaya,
    # jadi TIDAK ditambahkan lagi ke invoice admin.
    admin_core_fee = max(base_program_fee - registration_fee, 0.0)
    admin_invoice_total = admin_core_fee
    grand_total = base_program_fee

    return {
        "program": program,
        "base_program_fee": base_program_fee,
        "registration_fee": registration_fee,
        "admin_core_fee": admin_core_fee,
        "transport_fee": 0.0,
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


# ---------- Legacy local PDF ----------
def build_invoice_pdf(invoice: Dict[str, Any], student: Dict[str, Any]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
    )
    styles = getSampleStyleSheet()
    story = []

    invoice_type = safe_text(invoice.get("invoice_type") or "Invoice")
    story.append(Paragraph("<b>NIHAOMA STUDENT OPERATIONS</b>", styles["Title"]))
    story.append(Paragraph(f"Invoice {invoice_type}", styles["Heading2"]))
    story.append(Spacer(1, 8))

    info_data = [
        ["Kode Invoice", safe_text(invoice.get("kode_invoice"))],
        ["Tanggal Invoice", maybe_date(invoice.get("tanggal_invoice"))],
        ["Jenis Invoice", invoice_type],
        ["Student ID", safe_text(invoice.get("student_id"))],
        ["Nama Mahasiswa", safe_text(invoice.get("nama_mahasiswa") or student.get("nama_lengkap"))],
        ["Program", safe_text(invoice.get("program") or student.get("program_diminati"))],
        ["Intake", safe_text(student.get("intake"))],
    ]
    info_table = Table(info_data, colWidths=[48 * mm, 117 * mm])
    info_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#f6f6f6")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d8d8d8")),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(info_table)
    story.append(Spacer(1, 12))

    harga = to_number(invoice.get("harga_program"))
    dibayar = to_number(invoice.get("sudah_dibayar"))
    sisa = to_number(invoice.get("sisa_tagihan"))
    biaya_pendaftaran = to_number(invoice.get("biaya_pendaftaran"))
    biaya_admin = to_number(invoice.get("biaya_admin"))
    biaya_transport = to_number(invoice.get("biaya_transport"))

    detail_rows = [["Deskripsi", "Nilai"]]
    if invoice_type.lower() == "pendaftaran":
        detail_rows.append(["Biaya pendaftaran", format_currency(biaya_pendaftaran or harga)])
    elif invoice_type.lower() == "admin":
        detail_rows.append(["Biaya admin", format_currency(biaya_admin or harga)])
        detail_rows.append(["Total invoice admin", format_currency(harga)])
    else:
        detail_rows.append([safe_text(invoice.get("deskripsi_biaya") or "Biaya"), format_currency(harga)])

    detail_rows.extend([
        ["Sudah Dibayar", format_currency(dibayar)],
        ["Sisa Tagihan", format_currency(sisa)],
        ["Status Pelunasan", safe_text(invoice.get("status_pelunasan"))],
    ])

    bill_table = Table(detail_rows, colWidths=[110 * mm, 55 * mm])
    bill_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2937")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d8d8d8")),
                ("PADDING", (0, 0), (-1, -1), 6),
                ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ("BACKGROUND", (0, 1), (-1, -1), colors.white),
            ]
        )
    )
    story.append(bill_table)
    story.append(Spacer(1, 10))

    note_text = safe_text(invoice.get("catatan_invoice")) or "Catatan belum diisi."
    story.append(Paragraph(f"<b>Catatan:</b> {note_text}", styles["BodyText"]))
    story.append(Spacer(1, 24))
    story.append(Paragraph("Terima kasih.", styles["BodyText"]))

    doc.build(story)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf


def build_preview_invoice_url(invoice_id: str) -> str:
    ensure_config()
    params = urlencode(
        {
            "action": "preview_invoice",
            "token": WRITE_TOKEN,
            "invoice_id": invoice_id,
        }
    )
    return f"{SCRIPT_URL}?{params}"

def safe_paragraph_text(value) -> str:
    text = "" if value is None else str(value)
    return html.escape(text).replace("\n", "<br/>")


def format_date_id(value) -> str:
    if value is None or value == "" or (isinstance(value, float) and pd.isna(value)):
        return "-"
    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return safe_text(value) or "-"
    return ts.strftime("%d/%m/%Y")


def register_fonts():
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass


def expected_invoice_code(year_value: Any, student_id_value: Any) -> str:
    year_text = safe_text(year_value)
    if not year_text:
        return ""
    ts = pd.to_datetime(year_text, errors="coerce")
    year2 = ts.strftime("%y") if not pd.isna(ts) else ""
    student_text = safe_text(student_id_value)
    m = re.search(r"(\d+)$", student_text)
    seq = m.group(1).zfill(2)[-2:] if m else "00"
    return f"NHEC-{year2}{seq}" if year2 else ""


def invoice_row_for_pdf(invoice: Dict[str, Any], student: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "Kode Invoice": safe_text(invoice.get("kode_invoice")),
        "Tanggal Input": safe_text(invoice.get("tanggal_invoice")),
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
        "Intake / Keterangan": safe_text(
            student.get("intake") or invoice.get("catatan_invoice") or student.get("catatan_admin")
        ),
    }


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
    top.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
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
        [Paragraph("1", value), Paragraph(safe_paragraph_text(f"Biaya program {record.get('Program', '-')}" + ("\nKeterangan: " + str(record.get('Intake / Keterangan', '-') or '-'))), value), Paragraph("IDR", value), Paragraph(format_currency(record.get("Harga Program", 0)), value)],
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
    footer_combo.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(footer_combo)

    doc.build(story)
    buffer.seek(0)
    return buffer.read()

ORANGE_COLORS = ["#C2410C", "#EA580C", "#F97316", "#FB923C", "#FDBA74", "#FED7AA"]

TRACKING_STAGE_COLORS = {
    # Spektrum hijau: proses aktif
    "On Progress": "#166534",
    "Submitted": "#16A34A",
    "Waiting Review": "#86EFAC",

    # Spektrum biru: hasil akademik/beasiswa
    "LOA Issued": "#1D4ED8",
    "LoA Issued": "#1D4ED8",
    "Scholarship Result": "#93C5FD",

    # Kuning: interview
    "Interview": "#FACC15",

    # Abu-abu: belum mulai
    "Belum Mulai": "#9CA3AF",

    # Spektrum merah: gagal/dibatalkan
    "Rejected": "#991B1B",
    "Revoked": "#DC2626",
    "Withdraw": "#FCA5A5",
}

def get_chart_theme() -> dict:
    base = str(st.get_option("theme.base") or "light").lower()
    dark = base == "dark"
    return {
        "font": "#F8FAFC" if dark else "#4B5563",
        "grid": "rgba(255,255,255,0.10)" if dark else "rgba(180, 83, 9, 0.10)",
        "line": "rgba(255,255,255,0.18)" if dark else "rgba(180, 83, 9, 0.18)",
        "paper": "rgba(0,0,0,0)",
        "plot": "rgba(0,0,0,0)",
        "legend_bg": "rgba(255,255,255,0)" if not dark else "rgba(0,0,0,0)",
    }

def style_pie_chart(fig, title: str, hole: float = 0.42):
    theme = get_chart_theme()

    fig.update_traces(
        hole=hole,
        textinfo="percent",
        textfont_size=14,
        marker=dict(
            line=dict(color="rgba(255,255,255,0.85)", width=2)
        ),
        hovertemplate="<b>%{label}</b><br>Jumlah: %{value}<br>Persen: %{percent}<extra></extra>",
        pull=[0.02] * len(fig.data[0]["labels"]) if fig.data else None,
    )

    fig.update_layout(
        title=dict(
            text=title,
            x=0.02,
            xanchor="left",
            font=dict(size=20, color=theme["font"]),
        ),
        font=dict(color=theme["font"], size=13),
        paper_bgcolor=theme["paper"],
        plot_bgcolor=theme["plot"],
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.12,
            xanchor="left",
            x=0,
            font=dict(size=12, color=theme["font"]),
            bgcolor=theme["legend_bg"],
        ),
        margin=dict(l=20, r=20, t=60, b=95),
        height=430,
    )
    return fig

def style_bar_chart(fig, title: str):
    theme = get_chart_theme()

    fig.update_traces(
        marker_line_color=theme["line"],
        marker_line_width=1.2,
        opacity=0.92,
        texttemplate="%{y:,.0f}",
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>Nilai: %{y:,.0f}<extra></extra>",
    )

    fig.update_layout(
        title=dict(
            text=title,
            x=0.02,
            xanchor="left",
            font=dict(size=20, color=theme["font"]),
        ),
        font=dict(color=theme["font"], size=13),
        paper_bgcolor=theme["paper"],
        plot_bgcolor=theme["plot"],
        coloraxis_showscale=False,
        margin=dict(l=20, r=20, t=60, b=70),
        height=430,
        xaxis=dict(
            title="",
            showgrid=False,
            linecolor=theme["line"],
            tickfont=dict(color=theme["font"]),
        ),
        yaxis=dict(
            title="",
            showgrid=True,
            gridcolor=theme["grid"],
            zeroline=False,
            linecolor=theme["line"],
            tickfont=dict(color=theme["font"]),
        ),
    )
    return fig

# ---------- Student Tracking ----------
TRACKING_STATUS_DEFAULT = [
    "Belum Mulai",
    "On Progress",
    "Submitted",
    "Waiting Review",
    "Interview",
    "LOA Issued",
    "Scholarship Result",
    "Visa Process",
    "Ready to Depart",
    "Done",
    "Rejected",
    "Revoked",
    "Withdraw",
]
TRACKING_SUBMIT_DEFAULT = ["", "No", "Yes", "Re-submit", "Pending"]
TRACKING_INTERVIEW_DEFAULT = ["", "Not Required", "Pending", "Scheduled", "Done", "Failed"]
TRACKING_LOA_DEFAULT = ["", "Pending", "Issued", "Rejected", "Not Required"]
TRACKING_SCHOLARSHIP_DEFAULT = ["", "Pending", "Waiting Result", "Approved", "Rejected", "Not Applied", "Not Available"]
TRACKING_VISA_DEFAULT = ["", "Not Started", "Preparing", "Submitted", "Approved", "Rejected"]
TRACKING_PRIORITY_DEFAULT = ["Rendah", "Sedang", "Tinggi", "Urgent"]


def tracking_ref_options(refs: Dict[str, Any], key: str, default_options: List[str], current_value: Any = "") -> List[str]:
    return ensure_option_list(refs.get(key, default_options), current_value or (default_options[0] if default_options else ""))


def normalize_tracking_progress(value: Any) -> float:
    raw = to_number(value)
    if raw <= 0:
        return 0.0
    if raw <= 1:
        raw *= 100
    return max(0.0, min(float(raw), 100.0))


def calculate_tracking_progress(row: Dict[str, Any]) -> float:
    explicit_score = normalize_tracking_progress(row.get("progress_score"))
    if explicit_score > 0:
        return explicit_score

    status = safe_text(row.get("status_pendaftaran")).strip().lower()
    if status in ["done", "ready to depart"]:
        return 100.0
    if status in ["rejected", "revoked", "withdraw", "tidak lolos"]:
        return 0.0

    score = 0.0
    submit = safe_text(row.get("sudah_submit")).strip().lower()
    interview = safe_text(row.get("interview")).strip().lower()
    loa = safe_text(row.get("loa")).strip().lower()
    scholarship = safe_text(row.get("scholarship")).strip().lower()
    visa = safe_text(row.get("visa")).strip().lower()

    if status in ["on progress", "waiting review", "interview", "loa issued", "scholarship result", "visa process"]:
        score = max(score, 10.0)
    if submit in ["yes", "re-submit", "submitted"]:
        score = max(score, 25.0)
    if interview in ["scheduled", "done", "not required"]:
        score = max(score, 40.0)
    if loa in ["issued", "yes", "received"]:
        score = max(score, 65.0)
    if scholarship in ["approved", "rejected", "not applied", "not available"]:
        score = max(score, 78.0)
    if visa in ["submitted", "approved"]:
        score = max(score, 90.0 if visa == "submitted" else 100.0)
    return max(0.0, min(score, 100.0))


def tracking_stage(row: Dict[str, Any]) -> str:
    status = safe_text(row.get("status_pendaftaran")).strip()
    if status:
        return status
    progress = calculate_tracking_progress(row)
    if progress >= 100:
        return "Done"
    if progress >= 90:
        return "Visa Process"
    if progress >= 78:
        return "Scholarship Result"
    if progress >= 65:
        return "LOA Issued"
    if progress >= 40:
        return "Interview"
    if progress >= 25:
        return "Submitted"
    if progress > 0:
        return "On Progress"
    return "Belum Mulai"


def prepare_tracking_df(tracking_df: pd.DataFrame) -> pd.DataFrame:
    if tracking_df.empty:
        return tracking_df.copy()
    df = tracking_df.copy()
    for col in [
        "tracking_id", "student_id", "nama_siswa", "program", "universitas", "negara_kota", "jurusan",
        "status_pendaftaran", "sudah_submit", "interview", "loa", "scholarship", "visa", "web_pendaftaran",
        "portal_username", "portal_password", "deadline", "pic", "prioritas", "next_action", "catatan",
        "tanggal_update", "updated_at", "updated_by", "progress_history",
    ]:
        if col not in df.columns:
            df[col] = ""
    df["progress_percent"] = df.apply(lambda r: calculate_tracking_progress(r.to_dict()), axis=1)
    df["stage"] = df.apply(lambda r: tracking_stage(r.to_dict()), axis=1)
    return df


def build_tracking_options(tracking_df: pd.DataFrame) -> tuple[List[str], Dict[str, str]]:
    labels = []
    mapping = {}
    if tracking_df.empty:
        return labels, mapping

    for _, row in tracking_df.iterrows():
        row_dict = row.to_dict()
        tracking_id = safe_text(row.get("tracking_id"))
        stage = safe_text(row.get("stage") or tracking_stage(row_dict))
        progress = calculate_tracking_progress(row_dict)

        label_parts = [
            tracking_id or "NO-ID",
            safe_text(row.get("nama_siswa")),
            safe_text(row.get("universitas")),
            safe_text(row.get("program")),
            stage,
            f"{progress:.0f}%",
        ]

        label = " - ".join([part for part in label_parts if part])
        labels.append(label)
        mapping[label] = tracking_id

    return labels, mapping


def find_tracking(tracking_df: pd.DataFrame, tracking_id: str) -> Dict[str, Any]:
    if tracking_df.empty or "tracking_id" not in tracking_df.columns:
        return {}
    row_df = tracking_df[tracking_df["tracking_id"].astype(str) == str(tracking_id)]
    return row_df.iloc[0].to_dict() if not row_df.empty else {}


def filter_tracking_df(tracking_df: pd.DataFrame, keyword: str, program: str, universitas: str, status: str, pic: str) -> pd.DataFrame:
    df = prepare_tracking_df(tracking_df)
    if df.empty:
        return df
    if keyword:
        kw = keyword.lower().strip()
        mask = pd.Series(False, index=df.index)
        for col in ["tracking_id", "student_id", "nama_siswa", "universitas", "jurusan", "pic", "portal_username"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.lower().str.contains(kw, na=False)
        df = df[mask]
    if program != "Semua" and "program" in df.columns:
        df = df[df["program"].astype(str) == program]
    if universitas != "Semua" and "universitas" in df.columns:
        df = df[df["universitas"].astype(str) == universitas]
    if status != "Semua" and "stage" in df.columns:
        df = df[df["stage"].astype(str) == status]
    if pic != "Semua" and "pic" in df.columns:
        df = df[df["pic"].astype(str) == pic]
    return df


def render_progress_badge(progress: float, label: str = "") -> None:
    pct = int(max(0, min(progress, 100)))
    caption = f"{pct}%" + (f" • {label}" if label else "")
    st.progress(pct / 100)
    st.caption(caption)


def now_jakarta() -> datetime:
    """Waktu server diseragamkan ke WIB agar tanggal update konsisten untuk tim Indonesia."""
    try:
        if ZoneInfo is not None:
            return datetime.now(ZoneInfo("Asia/Jakarta"))
    except Exception:
        pass
    return datetime.now()


def current_update_date() -> str:
    return now_jakarta().strftime("%Y-%m-%d")


def current_update_timestamp() -> str:
    return now_jakarta().strftime("%Y-%m-%d %H:%M:%S")


def parse_tracking_history(row: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Baca riwayat progress dari kolom JSON di student_tracking."""
    raw = ""
    for key in ["progress_history", "tracking_history", "riwayat_progress", "history_progress"]:
        value = row.get(key)
        if safe_text(value).strip():
            raw = value
            break

    if isinstance(raw, list):
        items = raw
    elif isinstance(raw, dict):
        items = [raw]
    else:
        text = safe_text(raw).strip()
        if not text:
            items = []
        else:
            try:
                loaded = json.loads(text)
                items = loaded if isinstance(loaded, list) else [loaded]
            except Exception:
                # Fallback jika pernah disimpan sebagai baris teks biasa.
                items = []
                for line in text.splitlines():
                    line = line.strip()
                    if not line:
                        continue
                    parts = [part.strip() for part in re.split(r"\s*[|;]\s*", line)]
                    items.append({
                        "tanggal": parts[0] if len(parts) > 0 else "",
                        "stage": parts[1] if len(parts) > 1 else line,
                        "progress_percent": parts[2].replace("%", "") if len(parts) > 2 else "",
                        "catatan": parts[3] if len(parts) > 3 else "",
                    })

    normalized: List[Dict[str, Any]] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        stage = safe_text(item.get("stage") or item.get("status") or item.get("status_pendaftaran")).strip()
        tanggal = safe_text(item.get("tanggal") or item.get("tanggal_update") or item.get("date") or item.get("updated_at")).strip()
        updated_at = safe_text(item.get("updated_at") or item.get("timestamp") or tanggal).strip()
        progress = normalize_tracking_progress(item.get("progress_percent") or item.get("progress_score") or item.get("progress"))
        normalized.append({
            "tanggal": tanggal,
            "updated_at": updated_at,
            "stage": stage,
            "progress_percent": progress,
            "from_stage": safe_text(item.get("from_stage") or item.get("previous_stage")),
            "next_action": safe_text(item.get("next_action")),
            "catatan": safe_text(item.get("catatan") or item.get("note")),
            "updated_by": safe_text(item.get("updated_by") or item.get("pic")),
        })

    def sort_key(item: Dict[str, Any]):
        parsed = pd.to_datetime(item.get("updated_at") or item.get("tanggal"), errors="coerce")
        if pd.isna(parsed):
            return pd.Timestamp.min
        return parsed

    normalized = sorted(normalized, key=sort_key)
    return normalized[-80:]


def build_tracking_history_event(current_row: Dict[str, Any], new_payload: Dict[str, Any]) -> Dict[str, Any]:
    update_date = safe_text(new_payload.get("tanggal_update")) or current_update_date()
    update_ts = safe_text(new_payload.get("updated_at")) or current_update_timestamp()

    after_row = dict(current_row or {})
    after_row.update(new_payload or {})

    before_stage = tracking_stage(current_row or {})
    after_stage = tracking_stage(after_row)
    after_progress = calculate_tracking_progress(after_row)

    return {
        "tanggal": update_date,
        "updated_at": update_ts,
        "from_stage": before_stage,
        "stage": after_stage,
        "progress_percent": round(after_progress, 1),
        "status_pendaftaran": safe_text(after_row.get("status_pendaftaran")),
        "sudah_submit": safe_text(after_row.get("sudah_submit")),
        "interview": safe_text(after_row.get("interview")),
        "loa": safe_text(after_row.get("loa")),
        "scholarship": safe_text(after_row.get("scholarship")),
        "visa": safe_text(after_row.get("visa")),
        "next_action": safe_text(after_row.get("next_action")),
        "catatan": safe_text(after_row.get("catatan")),
        "updated_by": safe_text(after_row.get("updated_by") or after_row.get("pic") or "Admin"),
    }


def with_tracking_update_metadata(current_row: Dict[str, Any], payload: Dict[str, Any]) -> Dict[str, Any]:
    """Tambahkan tanggal otomatis dan append progress_history sebelum dikirim ke Apps Script."""
    enriched = dict(payload or {})
    enriched["tanggal_update"] = current_update_date()
    enriched["updated_at"] = current_update_timestamp()

    history = parse_tracking_history(current_row or {})
    history.append(build_tracking_history_event(current_row or {}, enriched))
    enriched["progress_history"] = json.dumps(history[-80:], ensure_ascii=False)
    return enriched


def update_tracking_progress_with_history(tracking_id: str, current_row: Dict[str, Any], payload: Dict[str, Any]) -> Dict[str, Any]:
    return api_post(
        "update_tracking_progress",
        {
            "tracking_id": tracking_id,
            "payload": with_tracking_update_metadata(current_row, payload),
        },
    )


def progress_history_to_df(row: Dict[str, Any]) -> pd.DataFrame:
    history = parse_tracking_history(row)

    if not history:
        # Fallback agar data existing tetap punya titik awal walaupun belum ada kolom history.
        current_stage = tracking_stage(row)
        current_progress = calculate_tracking_progress(row)
        current_date = safe_text(row.get("tanggal_update") or row.get("updated_at") or row.get("deadline"))
        if current_stage or current_progress > 0:
            history = [{
                "tanggal": current_date,
                "updated_at": safe_text(row.get("updated_at") or current_date),
                "stage": current_stage,
                "progress_percent": current_progress,
                "from_stage": "",
                "next_action": safe_text(row.get("next_action")),
                "catatan": safe_text(row.get("catatan")),
                "updated_by": safe_text(row.get("updated_by") or row.get("pic")),
            }]

    return pd.DataFrame(history)


def tracking_stage_color(stage: Any) -> str:
    stage_text = safe_text(stage).strip()
    return TRACKING_STAGE_COLORS.get(stage_text, TRACKING_STAGE_COLORS.get(stage_text.title(), "#F97316"))

def render_tracking_progress_circle(row: Dict[str, Any], key_suffix: str = "") -> None:
    progress = normalize_tracking_progress(calculate_tracking_progress(row))
    stage = safe_text(tracking_stage(row)) or "On Progress"
    color = tracking_stage_color(stage)
    percent = int(round(progress))
    safe_key = re.sub(r"[^A-Za-z0-9_]+", "_", safe_text(key_suffix) or "progress")

    html_block = f"""
    <style>
    .progress-wrap-{safe_key} {{
        display:flex;
        flex-direction:column;
        align-items:center;
        justify-content:center;
        padding:8px 0 2px 0;
    }}
    .progress-circle-{safe_key} {{
        --size: 250px;
        --thickness: 26px;
        width: var(--size);
        height: var(--size);
        border-radius: 50%;
        background:
            radial-gradient(closest-side, #fffdf9 calc(100% - var(--thickness)), transparent calc(100% - var(--thickness) + 1px) 99.9%, transparent 100%),
            conic-gradient({color} 0 {percent}%, #E5E7EB {percent}% 100%);
        display:flex;
        align-items:center;
        justify-content:center;
        box-shadow: 0 10px 24px rgba(0,0,0,0.04);
        margin: 0 auto;
    }}
    .progress-inner-{safe_key} {{
        text-align:center;
        line-height:1.1;
    }}
    .progress-num-{safe_key} {{
        font-size:56px;
        font-weight:800;
        color:#1f2a44;
        margin-bottom:10px;
    }}
    .progress-label-{safe_key} {{
        font-size:21px;
        font-weight:500;
        color:#8a94a6;
    }}
    .progress-stage-{safe_key} {{
        margin-top:14px;
        display:inline-block;
        padding:8px 14px;
        border-radius:999px;
        background:rgba(255,255,255,0.95);
        border:1px solid rgba(217,119,6,0.12);
        font-size:14px;
        font-weight:700;
        color:#374151;
    }}
    </style>
    <div class="progress-wrap-{safe_key}">
        <div class="progress-circle-{safe_key}">
            <div class="progress-inner-{safe_key}">
                <div class="progress-num-{safe_key}">{percent}%</div>
                <div class="progress-label-{safe_key}">Progress</div>
            </div>
        </div>
        <div class="progress-stage-{safe_key}">{html.escape(stage)}</div>
    </div>
    """

    if hasattr(st, "html"):
        st.html(html_block)
    else:
        st.markdown(html_block, unsafe_allow_html=True)


def render_tracking_history_flow(row: Dict[str, Any], title: str = "Flow Progress Mahasiswa") -> None:
    history_df = progress_history_to_df(row)

    st.markdown(f"### {title}")
    st.caption("Riwayat ini otomatis terbentuk setiap kali progress disimpan. Tanggal update diambil otomatis oleh sistem.")

    if history_df.empty:
        st.info("Belum ada history progress. History akan mulai tampil setelah update progress berikutnya.")
        return

    cards = []
    for idx, item in history_df.iterrows():
        stage = safe_text(item.get("stage")) or "Update Progress"
        color = tracking_stage_color(stage)
        tanggal = format_date_id(item.get("tanggal") or item.get("updated_at"))
        progress = normalize_tracking_progress(item.get("progress_percent"))
        by = safe_text(item.get("updated_by")) or "Admin"
        note = safe_text(item.get("catatan"))
        next_action = safe_text(item.get("next_action"))
        from_stage = safe_text(item.get("from_stage"))
        transition = f"{html.escape(from_stage)} → {html.escape(stage)}" if from_stage and from_stage != stage else html.escape(stage)
        note_html = f"<div class='flow-note'>📝 {html.escape(note)}</div>" if note else ""
        action_html = f"<div class='flow-note'>➡️ {html.escape(next_action)}</div>" if next_action else ""
        cards.append(f"""
<div class="flow-item">
    <div class="flow-dot" style="background:{color};">{idx + 1}</div>
    <div class="flow-card">
        <div class="flow-date">{html.escape(tanggal)} • {html.escape(by)}</div>
        <div class="flow-stage">{transition}</div>
        <div class="flow-progress"><span style="width:{max(2, min(progress, 100)):.0f}%; background:{color};"></span></div>
        <div class="flow-percent">Progress: {progress:.0f}%</div>
        {action_html}
        {note_html}
    </div>
</div>
""")

    html_block = f"""<style>
.flow-wrap {{
    position: relative;
    margin: 10px 0 20px 0;
    padding: 6px 0 4px 0;
}}
.flow-item {{
    display: flex;
    gap: 14px;
    position: relative;
    margin: 0 0 14px 0;
}}
.flow-item:not(:last-child)::after {{
    content: "";
    position: absolute;
    left: 16px;
    top: 34px;
    bottom: -18px;
    width: 2px;
    background: rgba(217, 119, 6, 0.22);
}}
.flow-dot {{
    min-width: 34px;
    height: 34px;
    border-radius: 999px;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    font-size: 13px;
    box-shadow: 0 6px 16px rgba(0,0,0,0.12);
    z-index: 1;
}}
.flow-card {{
    background: rgba(255,255,255,0.96);
    border: 1px solid rgba(217, 119, 6, 0.12);
    border-radius: 18px;
    padding: 12px 14px;
    width: 100%;
    box-shadow: 0 8px 18px rgba(0,0,0,0.04);
}}
.flow-date {{
    font-size: 12px;
    color: #6b7280;
    font-weight: 600;
    margin-bottom: 4px;
}}
.flow-stage {{
    font-size: 16px;
    color: #1f2937;
    font-weight: 800;
    margin-bottom: 8px;
}}
.flow-progress {{
    width: 100%;
    height: 8px;
    background: #ffedd5;
    border-radius: 999px;
    overflow: hidden;
    margin-bottom: 5px;
}}
.flow-progress span {{
    display: block;
    height: 100%;
    border-radius: 999px;
}}
.flow-percent, .flow-note {{
    font-size: 12px;
    color: #4b5563;
    margin-top: 4px;
}}
</style>
<div class="flow-wrap">
{''.join(cards)}
</div>"""

    # Gunakan st.html jika tersedia supaya HTML tidak kebaca sebagai code block.
    # Fallback ke st.markdown untuk versi Streamlit lama.
    if hasattr(st, "html"):
        st.html(html_block)
    else:
        st.markdown(html_block, unsafe_allow_html=True)

def extract_selected_stage_from_plotly_event(event: Any) -> str:
    """Ambil label stage dari klik/selection Plotly pie chart."""
    try:
        if event is None:
            return ""

        if hasattr(event, "selection"):
            selection = event.selection
            points = getattr(selection, "points", []) or []
        elif isinstance(event, dict):
            points = event.get("selection", {}).get("points", []) or []
        else:
            points = []

        if not points:
            return ""

        point = points[0]
        if not isinstance(point, dict):
            point = dict(point)

        for key in ["label", "x", "legendgroup", "customdata"]:
            value = point.get(key)
            if isinstance(value, (list, tuple)) and value:
                value = value[0]
            value = safe_text(value).strip()
            if value:
                return value

    except Exception:
        return ""

    return ""


def render_tracking_quick_update_form(row: Dict[str, Any], refs: Dict[str, Any], form_key_prefix: str = "quick") -> None:
    tracking_id = safe_text(row.get("tracking_id"))
    if not tracking_id:
        st.info("Data tracking tidak valid.")
        return

    st.markdown("#### Update cepat")
    render_progress_badge(calculate_tracking_progress(row), tracking_stage(row))

    safe_key = re.sub(r"[^A-Za-z0-9_]+", "_", f"{form_key_prefix}_{tracking_id}")

    with st.form(f"form_{safe_key}"):
        col1, col2, col3 = st.columns(3)

        status_options = tracking_ref_options(
            refs, "tracking_status", TRACKING_STATUS_DEFAULT, row.get("status_pendaftaran")
        )
        status_pendaftaran = col1.selectbox(
            "Status Pendaftaran",
            status_options,
            index=option_index(status_options, row.get("status_pendaftaran")),
            key=f"status_{safe_key}",
        )

        submit_options = tracking_ref_options(
            refs, "submit_status", TRACKING_SUBMIT_DEFAULT, row.get("sudah_submit")
        )
        sudah_submit = col2.selectbox(
            "Sudah Submit?",
            submit_options,
            index=option_index(submit_options, row.get("sudah_submit")),
            key=f"submit_{safe_key}",
        )

        interview_options = tracking_ref_options(
            refs, "interview_status", TRACKING_INTERVIEW_DEFAULT, row.get("interview")
        )
        interview = col3.selectbox(
            "Interview",
            interview_options,
            index=option_index(interview_options, row.get("interview")),
            key=f"interview_{safe_key}",
        )

        col4, col5, col6 = st.columns(3)

        loa_options = tracking_ref_options(
            refs, "loa_status", TRACKING_LOA_DEFAULT, row.get("loa")
        )
        loa = col4.selectbox(
            "LOA",
            loa_options,
            index=option_index(loa_options, row.get("loa")),
            key=f"loa_{safe_key}",
        )

        scholarship_options = tracking_ref_options(
            refs, "scholarship_status", TRACKING_SCHOLARSHIP_DEFAULT, row.get("scholarship")
        )
        scholarship = col5.selectbox(
            "Scholarship",
            scholarship_options,
            index=option_index(scholarship_options, row.get("scholarship")),
            key=f"scholarship_{safe_key}",
        )

        visa_options = tracking_ref_options(
            refs, "visa_status", TRACKING_VISA_DEFAULT, row.get("visa")
        )
        visa = col6.selectbox(
            "Visa",
            visa_options,
            index=option_index(visa_options, row.get("visa")),
            key=f"visa_{safe_key}",
        )

        progress_score = st.slider(
            "Progress Manual (%)",
            0,
            100,
            int(calculate_tracking_progress(row)),
            5,
            key=f"progress_{safe_key}",
        )

        next_action = st.text_input(
            "Next Action",
            value=safe_text(row.get("next_action")),
            key=f"next_{safe_key}",
        )

        catatan = st.text_area(
            "Catatan update",
            value=safe_text(row.get("catatan")),
            key=f"note_{safe_key}",
        )

        updated_by = st.text_input(
            "Updated by",
            value=safe_text(row.get("updated_by") or row.get("pic") or "Admin"),
            key=f"by_{safe_key}",
        )

        st.caption(f"Tanggal update otomatis: {format_date_id(current_update_date())}")

        if st.form_submit_button("Simpan Update", type="primary"):
            update_payload = {
                "status_pendaftaran": status_pendaftaran,
                "sudah_submit": sudah_submit,
                "interview": interview,
                "loa": loa,
                "scholarship": scholarship,
                "visa": visa,
                "progress_score": progress_score,
                "next_action": next_action,
                "catatan": catatan,
                "updated_by": updated_by,
            }
            result = update_tracking_progress_with_history(tracking_id, row, update_payload)

            if result.get("ok"):
                st.success("Progress tracking berhasil diperbarui.")
                st.session_state.pop("tracking_selected_stage", None)
                clear_cache_and_rerun()
            else:
                st.error(result.get("error", "Gagal update progress tracking"))


def render_tracking_stage_panel(stage: str, stage_rows: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.markdown(f"### Status: {stage}")
    st.caption(f"Total data: {len(stage_rows)}")

    if stage_rows.empty:
        st.info("Tidak ada data pada status ini.")
        return

    show_cols = [
        "tracking_id",
        "nama_siswa",
        "program",
        "universitas",
        "jurusan",
        "stage",
        "progress_percent",
        "tanggal_update",
        "deadline",
        "pic",
        "prioritas",
        "next_action",
    ]

    display = stage_rows[[c for c in show_cols if c in stage_rows.columns]].copy()

    if "progress_percent" in display.columns:
        display["progress_percent"] = display["progress_percent"].apply(lambda x: f"{to_number(x):.0f}%")

    st.dataframe(display, use_container_width=True, hide_index=True)

    options, mapping = build_tracking_options(stage_rows)

    if options:
        selected_label = st.selectbox(
            "Pilih data untuk update langsung",
            options,
            key=f"dialog_select_{stage}",
        )

        selected_tracking_id = mapping[selected_label]
        selected_row = find_tracking(stage_rows, selected_tracking_id)

        st.divider()
        st.write(
            f"**{safe_text(selected_row.get('nama_siswa'))}** — "
            f"{safe_text(selected_row.get('universitas'))} — "
            f"{safe_text(selected_row.get('program'))}"
        )

        with st.expander("Lihat flow progress mahasiswa", expanded=True):
            render_tracking_history_flow(selected_row)

        render_tracking_portal_account_panel(
            selected_row,
            tracking_id=selected_tracking_id,
            allow_update=False,
            form_key_prefix="dialog",
        )

        render_tracking_quick_update_form(selected_row, refs, form_key_prefix="dialog")

    if st.button("Tutup", key=f"close_stage_dialog_{stage}"):
        st.session_state.pop("tracking_selected_stage", None)
        st.session_state["tracking_stage_chart_nonce"] = st.session_state.get("tracking_stage_chart_nonce", 0) + 1
        st.rerun()


if hasattr(st, "dialog"):
    @st.dialog("Detail Student Tracking")
    def render_tracking_stage_dialog(stage: str, stage_rows: pd.DataFrame, refs: Dict[str, Any]) -> None:
        render_tracking_stage_panel(stage, stage_rows, refs)
else:
    def render_tracking_stage_dialog(stage: str, stage_rows: pd.DataFrame, refs: Dict[str, Any]) -> None:
        st.warning("Versi Streamlit ini belum mendukung pop-up dialog. Detail ditampilkan di halaman.")
        render_tracking_stage_panel(stage, stage_rows, refs)

def tracking_status_icon(stage: Any) -> str:
    stage_text = safe_text(stage).strip().lower()
    if stage_text in ["on progress", "submitted", "waiting review"]:
        return "🟢"
    if stage_text in ["loa issued", "loa issued", "scholarship result"]:
        return "🔵"
    if stage_text == "interview":
        return "🟡"
    if stage_text == "belum mulai":
        return "⚪"
    if stage_text in ["rejected", "revoked", "withdraw"]:
        return "🔴"
    if stage_text in ["visa process", "ready to depart", "done"]:
        return "✅"
    return "📌"


def ordered_tracking_stages(stages: List[Any]) -> List[str]:
    available = []
    for stage in stages:
        stage_text = safe_text(stage).strip()
        if stage_text and stage_text not in available:
            available.append(stage_text)

    default_order = []
    for stage in TRACKING_STATUS_DEFAULT:
        stage_text = safe_text(stage).strip()
        if stage_text and stage_text not in default_order:
            default_order.append(stage_text)

    ordered = [stage for stage in default_order if stage in available]
    ordered.extend(sorted([stage for stage in available if stage not in ordered]))
    return ordered


def render_tracking_status_filter_buttons(tracking: pd.DataFrame) -> None:
    if tracking.empty or "stage" not in tracking.columns:
        return

    status_df = (
        tracking.groupby("stage", dropna=False)
        .size()
        .reset_index(name="jumlah")
    )
    status_df["stage"] = status_df["stage"].apply(lambda x: safe_text(x).strip() or "Belum Diisi")
    count_map = dict(zip(status_df["stage"], status_df["jumlah"]))
    ordered_stages = ordered_tracking_stages(count_map.keys())

    st.markdown(
        """
        <div class="soft-card" style="padding:18px 20px 12px 20px; margin:2px 0 18px 0;">
            <div style="font-size:18px; font-weight:800; color:#92400e; margin-bottom:4px;">
                Filter Berdasarkan Status Progress
            </div>
            <div style="font-size:13px; color:#6b7280; margin-bottom:12px;">
                Klik salah satu status untuk membuka pop-up daftar mahasiswa, lalu update progress langsung dari nama yang dipilih.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    buttons_per_row = 4
    for start in range(0, len(ordered_stages), buttons_per_row):
        cols = st.columns(buttons_per_row)
        row_stages = ordered_stages[start:start + buttons_per_row]
        for idx, stage in enumerate(row_stages):
            stage_count = int(count_map.get(stage, 0))
            button_label = f"{tracking_status_icon(stage)} {stage}\n{stage_count} data"
            button_key = re.sub(r"[^A-Za-z0-9_]+", "_", f"btn_filter_stage_{stage}")
            if cols[idx].button(
                button_label,
                key=button_key,
                use_container_width=True,
                help=f"Lihat dan update data dengan status {stage}",
            ):
                st.session_state["tracking_selected_stage"] = stage
                st.session_state["tracking_stage_source"] = "button"

    if st.session_state.get("tracking_selected_stage"):
        if st.button("↩️ Tutup filter aktif", key="clear_tracking_stage_filter", use_container_width=True):
            st.session_state.pop("tracking_selected_stage", None)
            st.session_state.pop("tracking_stage_source", None)
            st.session_state["tracking_stage_chart_nonce"] = st.session_state.get("tracking_stage_chart_nonce", 0) + 1
            st.rerun()



def tracking_payload_from_form(
    current: Dict[str, Any],
    selected_student: Dict[str, Any],
    form_values: Dict[str, Any],
) -> Dict[str, Any]:
    payload = {
        "tracking_id": safe_text(current.get("tracking_id")),
        "student_id": safe_text(form_values.get("student_id") or selected_student.get("student_id") or current.get("student_id")),
        "nama_siswa": safe_text(form_values.get("nama_siswa") or selected_student.get("nama_lengkap") or current.get("nama_siswa")),
        "program": safe_text(form_values.get("program") or selected_student.get("program_diminati") or current.get("program")),
        "universitas": safe_text(form_values.get("universitas")),
        "negara_kota": safe_text(form_values.get("negara_kota")),
        "jurusan": safe_text(form_values.get("jurusan")),
        "status_pendaftaran": safe_text(form_values.get("status_pendaftaran")),
        "sudah_submit": safe_text(form_values.get("sudah_submit")),
        "interview": safe_text(form_values.get("interview")),
        "loa": safe_text(form_values.get("loa")),
        "scholarship": safe_text(form_values.get("scholarship")),
        "visa": safe_text(form_values.get("visa")),
        "web_pendaftaran": safe_text(form_values.get("web_pendaftaran")),
        "portal_username": safe_text(form_values.get("portal_username")),
        "portal_password": safe_text(form_values.get("portal_password")),
        "deadline": safe_text(form_values.get("deadline")),
        "pic": safe_text(form_values.get("pic")),
        "prioritas": safe_text(form_values.get("prioritas")),
        "next_action": safe_text(form_values.get("next_action")),
        "catatan": safe_text(form_values.get("catatan")),
        "progress_score": float(form_values.get("progress_score") or 0),
        "updated_by": safe_text(form_values.get("updated_by") or form_values.get("pic") or "Admin"),
    }
    return with_tracking_update_metadata(current, payload)


def render_tracking_portal_account_panel(
    row: Dict[str, Any],
    tracking_id: str = "",
    allow_update: bool = False,
    form_key_prefix: str = "portal",
) -> None:
    """Tampilkan informasi portal langsung di area Update Progress."""
    tracking_id = safe_text(tracking_id or row.get("tracking_id"))
    safe_key = re.sub(r"[^A-Za-z0-9_]+", "_", f"{form_key_prefix}_{tracking_id or 'no_id'}")

    with st.expander("Portal & Akun Pendaftaran", expanded=True):
        st.caption("Cek link, username, dan password portal langsung sebelum update progress.")

        website = safe_text(row.get("web_pendaftaran")).strip()
        username = safe_text(row.get("portal_username")).strip()
        password = safe_text(row.get("portal_password")).strip()

        left, right = st.columns([0.95, 1.55], gap="large")

        with left:
            render_tracking_progress_circle(row, key_suffix=f"portal_{safe_key}")

        with right:
            info_left, info_right = st.columns([1.15, 1], gap="large")

            with info_left:
                st.markdown("#### Detail Portal")
                st.write(f"**Nama:** {safe_text(row.get('nama_siswa')) or '-'}")
                st.write(f"**Universitas:** {safe_text(row.get('universitas')) or '-'}")
                st.write(f"**Program:** {safe_text(row.get('program')) or '-'}")
                st.write(f"**Website:** {website or '-'}")
                if website:
                    st.link_button("Buka Website Pendaftaran", website, use_container_width=True)

            with info_right:
                st.markdown("#### Akun Login")
                st.write(f"**Username:** `{username or '-'}`")
                show_password = st.toggle(
                    "Tampilkan password akun ini",
                    value=False,
                    key=f"show_password_{safe_key}",
                )
                if show_password:
                    st.write(f"**Password:** `{password or '-'}`")
                else:
                    st.write("**Password:** `••••••`")

        st.warning(
            "Catatan keamanan: password yang disimpan di Google Sheet tidak terenkripsi. "
            "Batasi akses file dan Apps Script hanya untuk admin yang berwenang."
        )

        if allow_update and tracking_id:
            st.divider()
            with st.expander("✏️ Edit Portal / Akun", expanded=False):
                st.caption("Klik untuk membuka form edit portal. Secara default form disembunyikan agar halaman lebih rapi.")
                with st.form(f"form_update_portal_{safe_key}"):
                    c1, c2, c3 = st.columns([1.4, 1, 1])
                    web_pendaftaran = c1.text_input(
                        "Link Website Pendaftaran",
                        value=website,
                        key=f"portal_web_{safe_key}",
                    )
                    portal_username = c2.text_input(
                        "Username Portal",
                        value=username,
                        key=f"portal_user_{safe_key}",
                    )
                    portal_password = c3.text_input(
                        "Password Portal",
                        value=password,
                        type="password",
                        key=f"portal_pass_{safe_key}",
                    )
                    updated_by = st.text_input(
                        "Updated by",
                        value=safe_text(row.get("updated_by") or row.get("pic") or "Admin"),
                        key=f"portal_by_{safe_key}",
                    )

                    if st.form_submit_button("Simpan Akun Portal", type="primary"):
                        result = api_post(
                            "update_tracking_progress",
                            {
                                "tracking_id": tracking_id,
                                "payload": {
                                    "web_pendaftaran": web_pendaftaran,
                                    "portal_username": portal_username,
                                    "portal_password": portal_password,
                                    "tanggal_update": current_update_date(),
                                    "updated_at": current_update_timestamp(),
                                    "updated_by": updated_by,
                                },
                            },
                        )
                        if result.get("ok"):
                            st.success("Akun portal berhasil diperbarui.")
                            clear_cache_and_rerun()
                        else:
                            st.error(result.get("error", "Gagal memperbarui akun portal"))


# ---------- AI Assistant ----------
def ai_tracking_context(tracking: pd.DataFrame, max_rows: int = 120) -> Dict[str, Any]:
    """Buat konteks aman untuk AI. Password portal tidak pernah dikirim ke model."""
    df = prepare_tracking_df(tracking)

    if df.empty:
        return {
            "summary": "Belum ada data tracking.",
            "stage_counts": {},
            "priority_follow_up": [],
            "records": [],
        }

    safe_cols = [
        "tracking_id", "student_id", "nama_siswa", "program", "universitas", "jurusan",
        "stage", "progress_percent", "tanggal_update", "updated_at", "deadline", "pic",
        "prioritas", "next_action", "catatan", "sudah_submit", "interview", "loa",
        "scholarship", "visa", "updated_by",
    ]
    existing_cols = [c for c in safe_cols if c in df.columns]
    data = df[existing_cols].copy()

    if "progress_percent" in data.columns:
        data["progress_percent"] = data["progress_percent"].apply(lambda x: round(to_number(x), 0))

    for col in data.columns:
        data[col] = data[col].apply(lambda x: safe_text(x))

    stage_counts = (
        df.groupby("stage", dropna=False)
        .size()
        .reset_index(name="jumlah")
        .sort_values("jumlah", ascending=False)
    )
    stage_count_map = {
        safe_text(row.get("stage") or "Belum Diisi"): int(row.get("jumlah") or 0)
        for _, row in stage_counts.iterrows()
    }

    priority_cols = [c for c in [
        "tracking_id", "nama_siswa", "program", "universitas", "stage", "progress_percent",
        "tanggal_update", "deadline", "pic", "prioritas", "next_action", "catatan"
    ] if c in df.columns]
    priority = df[priority_cols].copy()
    if not priority.empty:
        if "progress_percent" in priority.columns:
            priority["progress_percent"] = priority["progress_percent"].apply(lambda x: round(to_number(x), 0))
        sort_cols = [c for c in ["prioritas", "deadline", "progress_percent"] if c in priority.columns]
        if sort_cols:
            priority = priority.sort_values(sort_cols, ascending=[False, True, True][:len(sort_cols)])
        priority = priority.head(15)
        for col in priority.columns:
            priority[col] = priority[col].apply(lambda x: safe_text(x))

    return {
        "summary": f"Total tracking: {len(df)}. Rata-rata progress: {df['progress_percent'].mean():.0f}%.",
        "stage_counts": stage_count_map,
        "priority_follow_up": priority.to_dict(orient="records") if not priority.empty else [],
        "records": data.head(max_rows).to_dict(orient="records"),
        "security_note": "Data portal password tidak dikirim ke AI assistant.",
    }


def call_openai_chat(messages: List[Dict[str, str]], model: str, api_key: str) -> str:
    response = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json={
            "model": model,
            "messages": messages,
            "max_completion_tokens": 900,
        },
        timeout=90,
    )

    if response.status_code >= 400:
        try:
            detail = response.json().get("error", {}).get("message", response.text)
        except Exception:
            detail = response.text
        raise RuntimeError(detail)

    data = response.json()
    return safe_text(data.get("choices", [{}])[0].get("message", {}).get("content")).strip()


def call_gemini_chat(messages: List[Dict[str, str]], model: str, api_key: str) -> str:
    """Panggil Gemini API langsung via REST. Tidak perlu dependency tambahan."""
    system_parts: List[str] = []
    contents: List[Dict[str, Any]] = []

    for message in messages:
        role = safe_text(message.get("role")).strip().lower()
        content = safe_text(message.get("content")).strip()
        if not content:
            continue

        if role == "system":
            system_parts.append(content)
        elif role == "assistant":
            contents.append({"role": "model", "parts": [{"text": content}]})
        else:
            contents.append({"role": "user", "parts": [{"text": content}]})

    if not contents:
        contents.append({"role": "user", "parts": [{"text": "Bantu ringkas data student tracking."}]})

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload: Dict[str, Any] = {
        "contents": contents,
        "generationConfig": {
            "temperature": 0.25,
            "maxOutputTokens": 900,
        },
    }

    if system_parts:
        payload["systemInstruction"] = {"parts": [{"text": "\n\n".join(system_parts)}]}

    response = requests.post(
        url,
        params={"key": api_key},
        headers={"Content-Type": "application/json"},
        json=payload,
        timeout=90,
    )

    if response.status_code >= 400:
        try:
            detail_json = response.json()
            detail = detail_json.get("error", {}).get("message", response.text)
        except Exception:
            detail = response.text
        raise RuntimeError(detail)

    data = response.json()
    candidates = data.get("candidates", [])
    if not candidates:
        return ""

    parts = candidates[0].get("content", {}).get("parts", [])
    answer_parts = [safe_text(part.get("text")) for part in parts if safe_text(part.get("text"))]
    return "\n".join(answer_parts).strip()


def get_ai_provider_config() -> Dict[str, str]:
    """Ambil konfigurasi AI dari Streamlit Secrets / environment.

    Default diarahkan ke Gemini agar tidak lagi memakai OpenAI ketika user memakai Gemini API key.
    Jika AI_PROVIDER tidak diisi, app otomatis memilih Gemini bila GEMINI_API_KEY ada.
    Jika Gemini key tidak sengaja ditaruh di OPENAI_API_KEY, app tetap mencoba memakai Gemini
    asalkan key tersebut tidak diawali 'sk-'.
    """
    provider = safe_text(st.secrets.get("AI_PROVIDER", os.getenv("AI_PROVIDER", ""))).strip().lower()

    openai_key = safe_text(st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", ""))).strip()
    openai_model = safe_text(st.secrets.get("OPENAI_MODEL", os.getenv("OPENAI_MODEL", "gpt-4o-mini"))).strip() or "gpt-4o-mini"

    gemini_key = safe_text(st.secrets.get("GEMINI_API_KEY", os.getenv("GEMINI_API_KEY", ""))).strip()
    gemini_model = safe_text(st.secrets.get("GEMINI_MODEL", os.getenv("GEMINI_MODEL", "gemini-2.5-flash"))).strip() or "gemini-2.5-flash"

    if not provider:
        if gemini_key:
            provider = "gemini"
        elif openai_key and not openai_key.startswith("sk-"):
            # Fallback khusus: key Gemini kadang tidak sengaja ditempel ke OPENAI_API_KEY.
            provider = "gemini"
            gemini_key = openai_key
        else:
            provider = "openai"

    if provider not in ["gemini", "openai"]:
        provider = "gemini"

    if provider == "gemini":
        return {"provider": "gemini", "api_key": gemini_key, "model": gemini_model}

    return {"provider": "openai", "api_key": openai_key, "model": openai_model}


def call_ai_chat(messages: List[Dict[str, str]], provider: str, model: str, api_key: str) -> str:
    provider = safe_text(provider).lower().strip()
    if provider == "gemini":
        return call_gemini_chat(messages, model=model, api_key=api_key)
    return call_openai_chat(messages, model=model, api_key=api_key)


def build_ai_messages(tracking: pd.DataFrame, chat_history: List[Dict[str, str]], user_prompt: str) -> List[Dict[str, str]]:
    context = ai_tracking_context(tracking)
    system_prompt = """
Anda adalah AI Assistant Operasional Nihaoma Education Center.
Tugas utama: membantu admin membaca data student tracking, menentukan prioritas follow-up, membuat ringkasan progress, dan membuat draft pesan WhatsApp/email yang sopan.

Aturan penting:
1. Jawab hanya berdasarkan data tracking yang diberikan di konteks, kecuali pengguna meminta template umum.
2. Jangan menebak data yang tidak tersedia. Kalau data tidak ada, katakan data belum tersedia.
3. Jangan meminta, menampilkan, atau menebak password portal. Password memang tidak diberikan ke AI.
4. Jangan mengklaim sudah mengubah data. Anda hanya memberi saran atau draft teks; perubahan data tetap dilakukan melalui form Streamlit.
5. Gunakan bahasa Indonesia yang ringkas, praktis, dan siap dipakai admin.
6. Untuk daftar follow-up, prioritaskan status yang mandek, progress rendah, deadline dekat, prioritas tinggi/urgent, dan next_action kosong.
""".strip()

    context_prompt = (
        "Berikut konteks data student tracking terbaru dalam format JSON. "
        "Gunakan ini sebagai sumber utama jawaban:\n\n"
        + json.dumps(context, ensure_ascii=False, indent=2)
    )

    messages: List[Dict[str, str]] = [
        {"role": "system", "content": system_prompt},
        {"role": "system", "content": context_prompt},
    ]

    for item in chat_history[-10:]:
        role = item.get("role") if item.get("role") in ["user", "assistant"] else "user"
        content = safe_text(item.get("content"))
        if content:
            messages.append({"role": role, "content": content})

    messages.append({"role": "user", "content": user_prompt})
    return messages


def render_ai_assistant_tab(tracking: pd.DataFrame) -> None:
    st.markdown("### 🤖 AI Assistant Operasional Nihaoma")
    st.caption(
        "Asisten ini membaca data student tracking terbaru untuk membantu follow-up, ringkasan progress, dan draft pesan. "
        "Password portal tidak dikirim ke AI."
    )

    ai_config = get_ai_provider_config()
    ai_provider = ai_config.get("provider", "gemini")
    api_key = ai_config.get("api_key", "")
    model = ai_config.get("model", "gemini-2.5-flash")

    context = ai_tracking_context(tracking)
    m1, m2, m3 = st.columns(3)
    m1.metric("Data Tracking", len(context.get("records", [])))
    m2.metric("Avg Progress", f"{tracking['progress_percent'].mean():.0f}%" if not tracking.empty else "0%")
    m3.metric("Status Aktif", len(context.get("stage_counts", {})))

    st.caption(f"Provider aktif: **{ai_provider.upper()}** • Model: `{model}`")

    if not api_key:
        if ai_provider == "gemini":
            st.warning("GEMINI_API_KEY belum diisi di Streamlit Secrets, jadi AI Assistant belum bisa dipakai.")
            st.code(
                "AI_PROVIDER = \"gemini\"\nGEMINI_API_KEY = \"isi_api_key_gemini_anda\"\nGEMINI_MODEL = \"gemini-2.5-flash\"",
                language="toml",
            )
        else:
            st.warning("OPENAI_API_KEY belum diisi di Streamlit Secrets, jadi AI Assistant belum bisa dipakai.")
            st.code(
                "AI_PROVIDER = \"openai\"\nOPENAI_API_KEY = \"isi_api_key_openai_anda\"\nOPENAI_MODEL = \"gpt-4o-mini\"",
                language="toml",
            )
        st.info("Setelah secrets disimpan, redeploy/restart aplikasi Streamlit agar asisten aktif.")
        return

    if "nihaoma_ai_messages" not in st.session_state:
        st.session_state["nihaoma_ai_messages"] = [
            {
                "role": "assistant",
                "content": "Halo, saya Asisten AI Nihaoma. Saya bisa bantu cek prioritas follow-up, ringkas progress, atau buat draft pesan untuk mahasiswa/kampus.",
            }
        ]

    col_a, col_b, col_c, col_d = st.columns(4)
    if col_a.button("Follow-up prioritas", use_container_width=True):
        st.session_state["nihaoma_ai_pending_prompt"] = "Tolong buatkan daftar mahasiswa yang paling perlu difollow-up hari ini beserta alasannya."
    if col_b.button("Progress mandek", use_container_width=True):
        st.session_state["nihaoma_ai_pending_prompt"] = "Mahasiswa mana yang progress-nya terlihat mandek atau rendah? Berikan prioritas tindak lanjut."
    if col_c.button("Ringkasan dashboard", use_container_width=True):
        st.session_state["nihaoma_ai_pending_prompt"] = "Buatkan ringkasan kondisi student tracking saat ini untuk laporan singkat ke manajemen."
    if col_d.button("Reset chat", use_container_width=True):
        st.session_state.pop("nihaoma_ai_messages", None)
        st.session_state.pop("nihaoma_ai_pending_prompt", None)
        st.rerun()

    for message in st.session_state.get("nihaoma_ai_messages", []):
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

    pending_prompt = st.session_state.pop("nihaoma_ai_pending_prompt", "")
    user_prompt = pending_prompt or st.chat_input("Tanya asisten, misalnya: buatkan pesan WA follow-up untuk Poppy")

    if user_prompt:
        st.session_state["nihaoma_ai_messages"].append({"role": "user", "content": user_prompt})
        with st.chat_message("user"):
            st.markdown(user_prompt)

        with st.chat_message("assistant"):
            with st.spinner("Asisten sedang membaca data tracking..."):
                try:
                    messages = build_ai_messages(
                        tracking,
                        st.session_state.get("nihaoma_ai_messages", [])[:-1],
                        user_prompt,
                    )
                    answer = call_ai_chat(messages, provider=ai_provider, model=model, api_key=api_key)
                    if not answer:
                        answer = "Maaf, respons AI kosong. Coba ulangi pertanyaan dengan lebih spesifik."
                except Exception as exc:
                    answer = f"Maaf, AI Assistant belum bisa menjawab via {ai_provider.upper()}. Detail error: {safe_text(exc)}"
                st.markdown(answer)

        st.session_state["nihaoma_ai_messages"].append({"role": "assistant", "content": answer})


def render_student_tracking_module(students_df: pd.DataFrame, tracking_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Student Tracking")
    st.caption("Pantau progres pendaftaran per mahasiswa, per universitas, termasuk portal pendaftaran, username, dan password.")

    tracking = prepare_tracking_df(tracking_df)
    tabs = st.tabs(["Overview", "Daftar Tracking", "Tambah / Edit", "Update Progress", "AI Assistant"])

    with tabs[0]:
        if tracking.empty:
            st.info("Belum ada data student tracking. Tambahkan data pertama di tab Tambah / Edit.")
        else:
            render_tracking_status_filter_buttons(tracking)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Total Aplikasi", len(tracking))
            c2.metric("Unique Student", tracking["student_id"].replace("", pd.NA).dropna().nunique() or tracking["nama_siswa"].nunique())
            c3.metric("Avg Progress", f"{tracking['progress_percent'].mean():.0f}%")
            c4.metric("Submitted", int(tracking["sudah_submit"].astype(str).str.lower().isin(["yes", "re-submit", "submitted"]).sum()))
            c5.metric("LOA Issued", int(tracking["loa"].astype(str).str.lower().isin(["issued", "yes", "received"]).sum()))

            left, right = st.columns(2)
            with left:
                status_df = tracking.groupby("stage", dropna=False).size().reset_index(name="jumlah").sort_values("jumlah", ascending=False)
                fig_stage = px.pie(
                    status_df,
                    names="stage",
                    values="jumlah",
                    color="stage",
                    color_discrete_map=TRACKING_STAGE_COLORS,
                    color_discrete_sequence=ORANGE_COLORS,
                )
                fig_stage = style_pie_chart(fig_stage, "Distribusi Progress Student", hole=0.46)
                fig_stage.update_traces(customdata=status_df["stage"].astype(str))

                try:
                    chart_nonce = st.session_state.get("tracking_stage_chart_nonce", 0)
                    pie_event = st.plotly_chart(
                        fig_stage,
                        use_container_width=True,
                        config={"displayModeBar": False},
                        on_select="rerun",
                        selection_mode="points",
                        key=f"tracking_stage_pie_chart_{chart_nonce}",
                    )
                except TypeError:
                    # Fallback untuk versi Streamlit lama yang belum support klik chart.
                    st.plotly_chart(fig_stage, use_container_width=True, config={"displayModeBar": False})
                    pie_event = None

                selected_stage_from_chart = extract_selected_stage_from_plotly_event(pie_event)

                if selected_stage_from_chart:
                    st.session_state["tracking_selected_stage"] = selected_stage_from_chart

            with right:
                uni_df = (
                    tracking.groupby("universitas", dropna=False)["progress_percent"]
                    .mean()
                    .reset_index()
                    .replace("", "Belum Diisi")
                    .sort_values("progress_percent", ascending=False)
                    .head(12)
                )
                fig_uni = px.bar(
                    uni_df,
                    x="universitas",
                    y="progress_percent",
                    color="progress_percent",
                    color_continuous_scale=[
                        [0.00, "#FED7AA"],
                        [0.35, "#FDBA74"],
                        [0.70, "#F97316"],
                        [1.00, "#C2410C"],
                    ],
                )
                fig_uni = style_bar_chart(fig_uni, "Rata-rata Progress per Universitas")
                fig_uni.update_traces(texttemplate="%{y:.0f}%")
                st.plotly_chart(fig_uni, use_container_width=True, config={"displayModeBar": False})

            selected_stage = safe_text(st.session_state.get("tracking_selected_stage"))
            if selected_stage:
                stage_rows = tracking[tracking["stage"].astype(str) == selected_stage].copy()
                render_tracking_stage_dialog(selected_stage, stage_rows, refs)

            st.markdown("### Priority Follow-up")
            follow_cols = ["nama_siswa", "program", "universitas", "stage", "progress_percent", "tanggal_update", "deadline", "pic", "prioritas", "next_action", "catatan"]
            follow_df = tracking[follow_cols].copy()
            follow_df = follow_df.sort_values(["prioritas", "deadline", "progress_percent"], ascending=[False, True, True]).head(20)
            follow_df["progress_percent"] = follow_df["progress_percent"].apply(lambda x: f"{x:.0f}%")
            st.dataframe(follow_df, use_container_width=True, hide_index=True)

    with tabs[1]:
        if tracking.empty:
            st.info("Belum ada data tracking.")
        else:
            f1, f2, f3, f4, f5 = st.columns([2, 1.2, 1.5, 1.2, 1.2])
            keyword = f1.text_input("Cari", placeholder="Nama, universitas, jurusan, username, PIC")
            program_options = sorted([x for x in tracking["program"].dropna().astype(str).unique().tolist() if x])
            uni_options = sorted([x for x in tracking["universitas"].dropna().astype(str).unique().tolist() if x])
            stage_options = sorted([x for x in tracking["stage"].dropna().astype(str).unique().tolist() if x])
            pic_options = sorted([x for x in tracking["pic"].dropna().astype(str).unique().tolist() if x])
            selected_program = f2.selectbox("Program", ["Semua"] + program_options)
            selected_uni = f3.selectbox("Universitas", ["Semua"] + uni_options)
            selected_stage = f4.selectbox("Status", ["Semua"] + stage_options)
            selected_pic = f5.selectbox("PIC", ["Semua"] + pic_options)

            filtered = filter_tracking_df(tracking, keyword, selected_program, selected_uni, selected_stage, selected_pic)
            show_password = st.toggle("Tampilkan password portal", value=False)
            show_cols = [
                "tracking_id", "student_id", "nama_siswa", "program", "universitas", "jurusan",
                "stage", "progress_percent", "sudah_submit", "interview", "loa", "scholarship", "visa",
                "web_pendaftaran", "portal_username", "portal_password", "deadline", "pic", "prioritas", "next_action", "tanggal_update", "updated_at",
            ]
            display = filtered[[c for c in show_cols if c in filtered.columns]].copy()
            if "progress_percent" in display.columns:
                display["progress_percent"] = display["progress_percent"].apply(lambda x: f"{x:.0f}%")
            if "portal_password" in display.columns and not show_password:
                display["portal_password"] = display["portal_password"].apply(lambda x: "••••••" if safe_text(x) else "")
            st.dataframe(display, use_container_width=True, hide_index=True)
            st.caption(f"Total data tampil: {len(filtered)}")

    with tabs[2]:
        mode = st.radio("Mode", ["Tambah tracking baru", "Edit tracking existing"], horizontal=True)
        current: Dict[str, Any] = {}
        if mode == "Edit tracking existing":
            options, mapping = build_tracking_options(tracking)
            if not options:
                st.info("Belum ada data untuk diedit.")
                st.stop()
            selected_label = st.selectbox(
                "Pilih data tracking",
                options,
                key="edit_tracking_select",
                help="Format pilihan: Tracking ID - Nama - Universitas - Program - Stage - Progress terbaru",
            )
            current = find_tracking(tracking, mapping[selected_label])

        student_labels, student_map = build_student_options(students_df)
        default_student_index = 0
        if current.get("student_id"):
            for i, label in enumerate([""] + student_labels):
                if label and student_map.get(label) == safe_text(current.get("student_id")):
                    default_student_index = i
                    break

        selected_student_label = st.selectbox(
            "Hubungkan ke data Calon Mahasiswa",
            [""] + student_labels,
            index=default_student_index,
            key="tracking_student_picker",
        )
        selected_student = find_student(students_df, student_map.get(selected_student_label, "")) if selected_student_label else {}

        with st.form("form_upsert_student_tracking"):
            st.markdown("### Identitas & Tujuan")
            col1, col2, col3 = st.columns(3)
            student_id = col1.text_input("Student ID", value=safe_text(selected_student.get("student_id") or current.get("student_id")))
            nama_siswa = col2.text_input("Nama Siswa", value=safe_text(selected_student.get("nama_lengkap") or current.get("nama_siswa")))
            program_options = tracking_ref_options(refs, "program_diminati", refs.get("program", []), current.get("program") or selected_student.get("program_diminati"))
            program = col3.selectbox("Program", program_options, index=option_index(program_options, current.get("program") or selected_student.get("program_diminati")))

            col4, col5, col6 = st.columns(3)
            universitas = col4.text_input("Universitas", value=safe_text(current.get("universitas") or selected_student.get("kampus_tujuan")))
            negara_kota = col5.text_input("Negara / Kota", value=safe_text(current.get("negara_kota") or selected_student.get("kota_tujuan")))
            jurusan = col6.text_input("Jurusan", value=safe_text(current.get("jurusan")))

            st.markdown("### Progress Pendaftaran")
            col7, col8, col9 = st.columns(3)
            status_options = tracking_ref_options(refs, "tracking_status", TRACKING_STATUS_DEFAULT, current.get("status_pendaftaran"))
            status_pendaftaran = col7.selectbox("Status Pendaftaran", status_options, index=option_index(status_options, current.get("status_pendaftaran")))
            submit_options = tracking_ref_options(refs, "submit_status", TRACKING_SUBMIT_DEFAULT, current.get("sudah_submit"))
            sudah_submit = col8.selectbox("Sudah Submit?", submit_options, index=option_index(submit_options, current.get("sudah_submit")))
            interview_options = tracking_ref_options(refs, "interview_status", TRACKING_INTERVIEW_DEFAULT, current.get("interview"))
            interview = col9.selectbox("Interview", interview_options, index=option_index(interview_options, current.get("interview")))

            col10, col11, col12 = st.columns(3)
            loa_options = tracking_ref_options(refs, "loa_status", TRACKING_LOA_DEFAULT, current.get("loa"))
            loa = col10.selectbox("LOA", loa_options, index=option_index(loa_options, current.get("loa")))
            scholarship_options = tracking_ref_options(refs, "scholarship_status", TRACKING_SCHOLARSHIP_DEFAULT, current.get("scholarship"))
            scholarship = col11.selectbox("Scholarship", scholarship_options, index=option_index(scholarship_options, current.get("scholarship")))
            visa_options = tracking_ref_options(refs, "visa_status", TRACKING_VISA_DEFAULT, current.get("visa"))
            visa = col12.selectbox("Visa", visa_options, index=option_index(visa_options, current.get("visa")))

            draft_for_score = {
                "progress_score": current.get("progress_score"),
                "status_pendaftaran": status_pendaftaran,
                "sudah_submit": sudah_submit,
                "interview": interview,
                "loa": loa,
                "scholarship": scholarship,
                "visa": visa,
            }
            default_progress = calculate_tracking_progress(draft_for_score)
            progress_score = st.slider("Progress Manual (%)", min_value=0, max_value=100, value=int(default_progress), step=5)
            render_progress_badge(progress_score, status_pendaftaran)

            st.markdown("### Portal Pendaftaran & Follow-up")
            col13, col14, col15 = st.columns(3)
            web_pendaftaran = col13.text_input("Link Website Pendaftaran", value=safe_text(current.get("web_pendaftaran")))
            portal_username = col14.text_input("Username Portal", value=safe_text(current.get("portal_username")))
            portal_password = col15.text_input("Password Portal", value=safe_text(current.get("portal_password")), type="password")

            col16, col17, col18 = st.columns(3)
            deadline = col16.text_input("Deadline", value=maybe_date(current.get("deadline")), placeholder="YYYY-MM-DD atau catatan deadline")
            pic_default_options = tracking_ref_options(refs, "pic_admin", [], current.get("pic") or selected_student.get("pic_admin"))
            pic = col17.selectbox("PIC", pic_default_options, index=option_index(pic_default_options, current.get("pic") or selected_student.get("pic_admin")))
            priority_options = tracking_ref_options(refs, "prioritas", TRACKING_PRIORITY_DEFAULT, current.get("prioritas"))
            prioritas = col18.selectbox("Prioritas", priority_options, index=option_index(priority_options, current.get("prioritas")))

            next_action = st.text_input("Next Action", value=safe_text(current.get("next_action")))
            catatan = st.text_area("Catatan", value=safe_text(current.get("catatan")))
            updated_by = st.text_input("Updated by", value=safe_text(current.get("updated_by") or pic or "Admin"))
            st.caption(f"Tanggal update otomatis: {format_date_id(current_update_date())}")

            submitted = st.form_submit_button("Simpan Tracking", type="primary")
            if submitted:
                if not nama_siswa.strip() and not student_id.strip():
                    st.error("Minimal isi Nama Siswa atau Student ID.")
                elif not universitas.strip():
                    st.error("Universitas wajib diisi agar tracking mudah difilter.")
                else:
                    payload = tracking_payload_from_form(
                        current,
                        selected_student,
                        {
                            "student_id": student_id,
                            "nama_siswa": nama_siswa,
                            "program": program,
                            "universitas": universitas,
                            "negara_kota": negara_kota,
                            "jurusan": jurusan,
                            "status_pendaftaran": status_pendaftaran,
                            "sudah_submit": sudah_submit,
                            "interview": interview,
                            "loa": loa,
                            "scholarship": scholarship,
                            "visa": visa,
                            "web_pendaftaran": web_pendaftaran,
                            "portal_username": portal_username,
                            "portal_password": portal_password,
                            "deadline": deadline,
                            "pic": pic,
                            "prioritas": prioritas,
                            "next_action": next_action,
                            "catatan": catatan,
                            "progress_score": progress_score,
                            "updated_by": updated_by,
                        },
                    )
                    result = api_post("upsert_student_tracking", {"payload": payload})
                    if result.get("ok"):
                        st.success(f"Student tracking berhasil disimpan. ID: {result.get('tracking_id', payload.get('tracking_id'))}")
                        clear_cache_and_rerun()
                    else:
                        st.error(result.get("error", "Gagal menyimpan student tracking"))

    with tabs[3]:
        if tracking.empty:
            st.info("Belum ada data tracking untuk diupdate.")
        else:
            stage_filter_options = ["Semua"] + sorted([
                x for x in tracking["stage"].dropna().astype(str).unique().tolist() if x
            ])

            selected_stage_filter = st.selectbox(
                "Filter stage / status",
                stage_filter_options,
                key="progress_stage_filter",
            )

            progress_pool = tracking.copy()

            if selected_stage_filter != "Semua":
                progress_pool = progress_pool[progress_pool["stage"].astype(str) == selected_stage_filter]

            options, mapping = build_tracking_options(progress_pool)

            if not options:
                st.info("Tidak ada data pada filter stage ini.")
            else:
                selected_label = st.selectbox(
                    "Pilih data tracking",
                    options,
                    key="progress_tracking_select",
                    help="Format pilihan: Tracking ID - Nama - Universitas - Program - Stage - Progress terbaru",
                )

                selected_tracking_id = mapping[selected_label]
                row = find_tracking(progress_pool, selected_tracking_id)

                st.markdown(
                    f"### {safe_text(row.get('nama_siswa'))} — "
                    f"{safe_text(row.get('universitas'))} — "
                    f"{safe_text(row.get('program'))} — "
                    f"{tracking_stage(row)}"
                )

                render_progress_badge(calculate_tracking_progress(row), tracking_stage(row))

                render_tracking_portal_account_panel(
                    row,
                    tracking_id=selected_tracking_id,
                    allow_update=True,
                    form_key_prefix="progress",
                )

                render_tracking_history_flow(row)

                with st.form("form_update_tracking_progress"):
                    col1, col2, col3 = st.columns(3)

                    status_options = tracking_ref_options(refs, "tracking_status", TRACKING_STATUS_DEFAULT, row.get("status_pendaftaran"))
                    status_pendaftaran = col1.selectbox("Status Pendaftaran", status_options, index=option_index(status_options, row.get("status_pendaftaran")))

                    submit_options = tracking_ref_options(refs, "submit_status", TRACKING_SUBMIT_DEFAULT, row.get("sudah_submit"))
                    sudah_submit = col2.selectbox("Sudah Submit?", submit_options, index=option_index(submit_options, row.get("sudah_submit")))

                    interview_options = tracking_ref_options(refs, "interview_status", TRACKING_INTERVIEW_DEFAULT, row.get("interview"))
                    interview = col3.selectbox("Interview", interview_options, index=option_index(interview_options, row.get("interview")))

                    col4, col5, col6 = st.columns(3)

                    loa_options = tracking_ref_options(refs, "loa_status", TRACKING_LOA_DEFAULT, row.get("loa"))
                    loa = col4.selectbox("LOA", loa_options, index=option_index(loa_options, row.get("loa")))

                    scholarship_options = tracking_ref_options(refs, "scholarship_status", TRACKING_SCHOLARSHIP_DEFAULT, row.get("scholarship"))
                    scholarship = col5.selectbox("Scholarship", scholarship_options, index=option_index(scholarship_options, row.get("scholarship")))

                    visa_options = tracking_ref_options(refs, "visa_status", TRACKING_VISA_DEFAULT, row.get("visa"))
                    visa = col6.selectbox("Visa", visa_options, index=option_index(visa_options, row.get("visa")))

                    progress_score = st.slider("Progress Manual (%)", 0, 100, int(calculate_tracking_progress(row)), 5)
                    next_action = st.text_input("Next Action", value=safe_text(row.get("next_action")))
                    catatan = st.text_area("Catatan update", value=safe_text(row.get("catatan")))
                    updated_by = st.text_input("Updated by", value=safe_text(row.get("updated_by") or row.get("pic") or "Admin"))
                    st.caption(f"Tanggal update otomatis: {format_date_id(current_update_date())}")

                    if st.form_submit_button("Update Progress", type="primary"):
                        update_payload = {
                            "status_pendaftaran": status_pendaftaran,
                            "sudah_submit": sudah_submit,
                            "interview": interview,
                            "loa": loa,
                            "scholarship": scholarship,
                            "visa": visa,
                            "progress_score": progress_score,
                            "next_action": next_action,
                            "catatan": catatan,
                            "updated_by": updated_by,
                        }
                        result = update_tracking_progress_with_history(selected_tracking_id, row, update_payload)

                        if result.get("ok"):
                            st.success("Progress tracking berhasil diperbarui.")
                            clear_cache_and_rerun()
                        else:
                            st.error(result.get("error", "Gagal update progress tracking"))


    with tabs[4]:
        render_ai_assistant_tab(tracking)


# ---------- Dashboard ----------
def render_dashboard(students_df: pd.DataFrame, invoices_df: pd.DataFrame, payments_df: pd.DataFrame) -> None:
    st.subheader("Dashboard")
    top_left, top_right = st.columns([1.6, 1])

    with top_left:
        st.markdown("""
        <div class="soft-card">
            <div class="section-title">Selamat datang di dashboard Nihaoma</div>
            <div style="color:#5b6f64; font-size:16px;">
                Pantau calon mahasiswa, dokumen, invoice, pembayaran, dan progress operasional
                dalam satu tampilan yang lebih rapi.
            </div>
        </div>
        """, unsafe_allow_html=True)

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

    pay = payments_df.copy()
    if not pay.empty:
        pay["jumlah_pembayaran"] = pay.get("jumlah_pembayaran", 0).apply(to_number)

    total_students = len(active_students)
    total_invoice = len(inv)
    total_nilai_invoice = inv["harga_program"].sum() if not inv.empty else 0
    total_dibayar = inv["sudah_dibayar"].sum() if not inv.empty else 0
    total_outstanding = inv["sisa_tagihan"].sum() if not inv.empty else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Calon Mahasiswa", total_students)
    c2.metric("Total Invoice", total_invoice)
    c3.metric("Nilai Invoice", format_currency(total_nilai_invoice))
    c4.metric("Sudah Dibayar", format_currency(total_dibayar))
    c5.metric("Outstanding", format_currency(total_outstanding))

    latest_new_students = st.session_state.get("latest_new_students", [])

    if latest_new_students:
        st.markdown("""
        <div class="soft-card">
            <div class="section-title">Notifikasi Mahasiswa Baru</div>
        </div>
        """, unsafe_allow_html=True)

        for item in latest_new_students[:5]:
            st.success(
                f"{item['nama_lengkap']} ({item['student_id']})"
                + (f" • {item['program_diminati']}" if item["program_diminati"] else "")
            )

    st.markdown("## At Glance")

    q1, q2, q3, q4 = st.columns(4)

    with q1:
        if st.button("🎓 Buka Calon Mahasiswa", key="quick_students", use_container_width=True):
            go_to_page("Calon Mahasiswa")

    with q2:
        if st.button(" 📄 Buka Dokumen", key="quick_documents", use_container_width=True):
            go_to_page("Dokumen")

    with q3:
        if st.button("🧾 Buka Invoice", key="quick_invoice", use_container_width=True):
            go_to_page("Invoice & Pembayaran")

    with q4:
        if st.button("💳 Buka Pembayaran", key="quick_payment", use_container_width=True):
            go_to_page("Invoice & Pembayaran")

    left, right = st.columns(2)

    with left:
        st.markdown("**Distribusi Status Proses**")
        if active_students.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            status_df = (
                active_students.assign(
                    status_proses=active_students.get("status_proses", "").replace("", "Belum Diisi")
                )
                .groupby("status_proses", dropna=False)
                .size()
                .reset_index(name="jumlah")
                .sort_values("jumlah", ascending=False)
            )
            fig_status = px.pie(
                status_df,
                names="status_proses",
                values="jumlah",
                color_discrete_sequence=ORANGE_COLORS,
            )
            fig_status = style_pie_chart(fig_status, "Distribusi Status Proses", hole=0.48)
            st.plotly_chart(fig_status, use_container_width=True, config={"displayModeBar": False})

    with right:
        st.markdown("**Distribusi PIC**")
        if active_students.empty:
            st.info("Belum ada PIC.")
        else:
            pic_df = (
                active_students.assign(pic_admin=active_students.get("pic_admin", "").replace("", "Belum Assign"))
                .groupby("pic_admin", dropna=False)
                .size()
                .reset_index(name="jumlah")
                .sort_values("jumlah", ascending=False)
            )
            fig_pic = px.pie(
                pic_df,
                names="pic_admin",
                values="jumlah",
                color_discrete_sequence=["#FDBA74", "#F97316", "#C2410C", "#FED7AA"],
            )
            fig_pic = style_pie_chart(fig_pic, "Distribusi PIC", hole=0.46)
            st.plotly_chart(fig_pic, use_container_width=True, config={"displayModeBar": False})

    lower_left, lower_right = st.columns(2)
    with lower_left:
        st.markdown("**Invoice berdasarkan Status Pelunasan**")
        if inv.empty:
            st.info("Belum ada invoice.")
        else:
            pelunasan = (
                inv.assign(status_pelunasan=inv.get("status_pelunasan", "").replace("", "Belum Diisi"))
                .groupby("status_pelunasan", dropna=False)
                .size()
                .reset_index(name="jumlah")
            )
            fig_pelunasan = px.pie(
                pelunasan,
                names="status_pelunasan",
                values="jumlah",
                color_discrete_sequence=["#C2410C", "#F97316", "#FDBA74", "#FED7AA"],
            )
            fig_pelunasan = style_pie_chart(fig_pelunasan, "Invoice berdasarkan Status Pelunasan", hole=0.42)
            st.plotly_chart(fig_pelunasan, use_container_width=True, config={"displayModeBar": False})

    with lower_right:
        st.markdown("**Outstanding per Program**")
        if inv.empty:
            st.info("Belum ada invoice.")
        else:
            outstanding_df = (
                inv.groupby("program", dropna=False)["sisa_tagihan"]
                .sum()
                .reset_index()
                .sort_values("sisa_tagihan", ascending=False)
            )
            fig_outstanding = px.bar(
                outstanding_df,
                x="program",
                y="sisa_tagihan",
                color="sisa_tagihan",
                color_continuous_scale=[
                    [0.00, "#FED7AA"],
                    [0.35, "#FDBA74"],
                    [0.70, "#F97316"],
                    [1.00, "#C2410C"],
                ],
            )
            fig_outstanding = style_bar_chart(fig_outstanding, "Outstanding per Program")
            st.plotly_chart(fig_outstanding, use_container_width=True, config={"displayModeBar": False})


# ---------- Students ----------
def render_student_list(students_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Modul Calon Mahasiswa")
    tabs = st.tabs(["Daftar Mahasiswa", "Tambah Data", "Detail & Progress"])

    with tabs[0]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            col_search, col_program, col_status, col_pic = st.columns([2, 1, 1, 1])
            keyword = col_search.text_input("Cari mahasiswa", placeholder="Nama, student_id, email, no WhatsApp")
            program_options = [x for x in refs.get("program_diminati", refs.get("program", [])) if x]
            status_options = [x for x in refs.get("status_proses", []) if x]
            pic_options = [x for x in refs.get("pic_admin", []) if x]

            selected_program = col_program.selectbox("Program", ["Semua"] + program_options)
            selected_status = col_status.selectbox("Status Proses", ["Semua"] + status_options)
            selected_pic = col_pic.selectbox("PIC", ["Semua"] + pic_options)

            filtered = students_df.copy()
            if keyword:
                kw = keyword.lower()
                mask = pd.Series(False, index=filtered.index)
                for col in [c for c in ["student_id", "nama_lengkap", "email", "no_whatsapp", "program_diminati"] if c in filtered.columns]:
                    mask = mask | filtered[col].astype(str).str.lower().str.contains(kw, na=False)
                filtered = filtered[mask]
            if selected_program != "Semua" and "program_diminati" in filtered.columns:
                filtered = filtered[filtered["program_diminati"] == selected_program]
            if selected_status != "Semua" and "status_proses" in filtered.columns:
                filtered = filtered[filtered["status_proses"] == selected_status]
            if selected_pic != "Semua" and "pic_admin" in filtered.columns:
                filtered = filtered[filtered["pic_admin"] == selected_pic]

            display_columns = [
                c for c in [
                    "student_id", "nama_lengkap", "program_diminati", "estimasi_biaya",
                    "intake", "pic_admin", "status_proses", "tanggal_input"
                ] if c in filtered.columns
            ]
            display_df = filtered[display_columns].copy() if display_columns else filtered.copy()
            if "estimasi_biaya" in display_df.columns:
                if "program_diminati" in display_df.columns:
                    display_df["estimasi_biaya"] = display_df.apply(
                        lambda row: format_currency(
                            get_program_total_fee(
                                row.get("program_diminati"),
                                row.get("estimasi_biaya")
                            )
                        ),
                        axis=1
                    )
            else:
                display_df["estimasi_biaya"] = display_df["estimasi_biaya"].apply(format_currency)
            st.dataframe(display_df, use_container_width=True, hide_index=True)
            st.caption(f"Total data tampil: {len(filtered)}")

            student_options, student_map = build_student_options(filtered)

            if student_options:
                selected_label = st.selectbox("Pilih mahasiswa untuk aksi", student_options, key="student_action_id")
                selected_id = student_map[selected_label]
                action_col1, action_col2, action_col3 = st.columns([1, 1, 3])
                if action_col1.button("Edit data", use_container_width=True):
                    st.session_state["edit_student_id"] = selected_id
                if action_col2.button("Hapus data", use_container_width=True):
                    st.session_state["delete_student_id"] = selected_id

                if st.session_state.get("edit_student_id"):
                    edit_id = st.session_state["edit_student_id"]
                    student = find_student(students_df, edit_id)
                    if student:
                        st.markdown("### Form Edit Mahasiswa")
                        render_edit_form(student, refs)

                if st.session_state.get("delete_student_id"):
                    delete_id = st.session_state["delete_student_id"]
                    st.markdown("### Konfirmasi Hapus Mahasiswa")
                    st.warning("Aksi ini akan menghapus students_master dan data terkait jika endpoint delete_student sudah dipasang di Apps Script.")
                    confirm_text = st.text_input(f"Ketik {delete_id} untuk konfirmasi hapus", key="confirm_delete_text")
                    del_col1, del_col2 = st.columns(2)
                    if del_col1.button("Ya, hapus sekarang", type="primary", use_container_width=True):
                        if confirm_text != delete_id:
                            st.error("Konfirmasi tidak cocok.")
                        else:
                            result = api_post("delete_student", {"student_id": delete_id})
                            if result.get("ok"):
                                st.success("Data mahasiswa berhasil dihapus.")
                                if result.get("kode_invoice"):
                                    st.write(f"**Kode Invoice:** {safe_text(result.get('kode_invoice'))}")
                                if result.get("file_name"):
                                    st.write(f"**File PDF:** {safe_text(result.get('file_name'))}")
                                if result.get("folder_name"):
                                    st.write(f"**Folder Drive:** {safe_text(result.get('folder_name'))}")
                                if result.get("file_url"):
                                    st.link_button("Buka PDF di Google Drive", result["file_url"], use_container_width=True)
                                if result.get("folder_url"):
                                    st.link_button("Buka Folder Invoices", result["folder_url"], use_container_width=True)
                                if result.get("preview_url"):
                                    st.link_button("Buka Preview Invoice", result["preview_url"], use_container_width=True)
                                clear_cache_and_rerun()
                            else:
                                st.error(result.get("error", "Gagal menghapus data"))
                    if del_col2.button("Batal", use_container_width=True):
                        st.session_state.pop("delete_student_id", None)
                        st.session_state.pop("confirm_delete_text", None)
                        st.rerun()

    with tabs[1]:
        render_add_form(refs)

    with tabs[2]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            detail_options, detail_map = build_student_options(students_df)
            selected_detail_label = st.selectbox("Pilih mahasiswa", detail_options, key="detail_student_id")
            selected_detail_id = detail_map[selected_detail_label]
            row_df = students_df[students_df["student_id"].astype(str) == str(selected_detail_id)]
            if row_df.empty:
                st.info("Data tidak ditemukan.")
            else:
                student = row_df.iloc[0].to_dict()
                left, right = st.columns(2)
                with left:
                    st.markdown("### Detail Mahasiswa")
                    for field in [
                        "student_id", "nama_lengkap", "email", "no_whatsapp", "program_diminati",
                        "intake", "kampus_tujuan", "kota_tujuan", "status_proses", "pic_admin"
                    ]:
                        st.write(f"**{field}**: {safe_text(student.get(field))}")
                with right:
                    st.markdown("### Update Progress")
                    with st.form("form_update_progress"):
                        status_options = refs.get("status_proses", [safe_text(student.get("status_proses"))]) or [safe_text(student.get("status_proses"))]
                        next_action_options = refs.get("next_action", []) or [safe_text(student.get("next_action"))]
                        status_baru = st.selectbox("Status Baru", status_options, index=option_index(status_options, student.get("status_proses")))
                        next_action = st.selectbox("Next Action", [""] + next_action_options, index=option_index([""] + next_action_options, student.get("next_action")))
                        tanggal_next_action = st.text_input("Tanggal Next Action", value=maybe_date(student.get("tanggal_next_action")))
                        catatan = st.text_area("Catatan Progress")
                        updated_by = st.text_input("Updated by", value=safe_text(student.get("pic_admin")) or "Admin")
                        if st.form_submit_button("Simpan Progress"):
                            result = api_post(
                                "update_progress",
                                {
                                    "student_id": selected_detail_id,
                                    "status_baru": status_baru,
                                    "next_action": next_action,
                                    "tanggal_next_action": tanggal_next_action,
                                    "catatan": catatan,
                                    "updated_by": updated_by,
                                },
                            )
                            if result.get("ok"):
                                st.success("Progress berhasil diperbarui.")
                                clear_cache_and_rerun()
                            else:
                                st.error(result.get("error", "Gagal update progress"))


# ---------- Forms ----------
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
        intake_options_fixed = ["", "Maret", "September"]
        intake = col9.selectbox(
            "Intake",
            intake_options_fixed,
            index=option_index(intake_options_fixed, student.get("intake")),
        )

        col10, col11, col12 = st.columns(3)
        program_diminati = col10.selectbox("Program", program_options, index=option_index(program_options, student.get("program_diminati")))
        kampus_tujuan = col11.text_input("Kampus Tujuan", value=safe_text(student.get("kampus_tujuan")))
        kota_tujuan = col12.text_input("Kota Tujuan", value=safe_text(student.get("kota_tujuan")))

        col13, col14, col15 = st.columns(3)
        negara_tujuan = col13.text_input("Negara Tujuan", value=safe_text(student.get("negara_tujuan")))
        pic_admin = col14.text_input("PIC", value=safe_text(student.get("pic_admin")))
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
                st.session_state.pop("edit_student_id", None)
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
        intake = col9.selectbox("Intake", ["", "Maret", "September"])

        col10, col11, col12 = st.columns(3)
        program_diminati = col10.selectbox("Program", [""] + program_options)
        kampus_tujuan = col11.text_input("Kampus Tujuan")
        kota_tujuan = col12.text_input("Kota Tujuan")

        col13, col14, col15 = st.columns(3)
        negara_tujuan = col13.text_input("Negara Tujuan", value="China")
        pic_admin = col14.text_input("PIC")
        status_proses = col15.selectbox("Status Proses", status_options, index=0 if status_options else None)

        col16, col17 = st.columns(2)
        sumber_leads = col16.selectbox("Sumber Leads", [""] + lead_options)
        prioritas = col17.selectbox("Prioritas", [""] + priority_options)

        alamat = st.text_area("Alamat")
        catatan_admin = st.text_area("Catatan Admin")

        if st.form_submit_button("Tambah Mahasiswa"):
            if not nama_lengkap.strip():
                st.error("Nama lengkap wajib diisi.")
            else:
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
                    if result.get("duplicate"):
                        st.warning(f"Data duplikat. student_id existing: {result.get('student_id')}")
                    else:
                        st.success(f"Mahasiswa berhasil ditambahkan. ID: {result.get('student_id')}")
                    clear_cache_and_rerun()
                else:
                    st.error(result.get("error", "Gagal menambah mahasiswa"))


# ---------- Documents ----------
def render_documents_module(students_df: pd.DataFrame, documents_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Dokumen")
    tabs = st.tabs(["Upload Dokumen", "Daftar Dokumen"])

    with tabs[0]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            student_options, student_map = build_student_options(students_df)
            selected_student_label = st.selectbox("Pilih mahasiswa", student_options, key="doc_student_id")
            selected_student_id = student_map[selected_student_label]
            student = find_student(students_df, selected_student_id)
            doc_types = refs.get("required_doc_types", []) or ["Passport", "Ijazah", "Transkrip", "Foto", "Bukti Pembayaran"]

            with st.form("form_upload_document"):
                c1, c2, c3 = st.columns(3)
                jenis_dokumen = c1.selectbox("Jenis Dokumen", [""] + doc_types)
                uploaded_by = c2.text_input("Uploaded by", value=safe_text(student.get("pic_admin")) or "Admin")
                versi_dokumen = c3.text_input("Versi Dokumen", value="v1")
                status_verifikasi = st.selectbox("Status Verifikasi", refs.get("status_verifikasi", ["Belum Dicek"]))
                catatan_verifikasi = st.text_area("Catatan", value="")
                file = st.file_uploader(
                    "Upload file",
                    type=["pdf", "jpg", "jpeg", "png", "doc", "docx", "zip"],
                    key="document_uploader",
                )
                if st.form_submit_button("Upload Dokumen"):
                    if not file:
                        st.error("Pilih file terlebih dahulu.")
                    elif not jenis_dokumen:
                        st.error("Jenis dokumen wajib dipilih.")
                    else:
                        b64 = base64.b64encode(file.read()).decode("utf-8")
                        result = api_post(
                            "upload_document",
                            {
                                "student_id": selected_student_id,
                                "nama_mahasiswa": safe_text(student.get("nama_lengkap")),
                                "jenis_dokumen": jenis_dokumen,
                                "nama_file": document_filename(
                                    selected_student_id,
                                    safe_text(student.get("nama_lengkap")),
                                    jenis_dokumen,
                                    file.name,
                                ),
                                "mime_type": file.type or "application/octet-stream",
                                "file_base64": b64,
                                "uploaded_by": uploaded_by,
                                "status_verifikasi": status_verifikasi,
                                "catatan_verifikasi": catatan_verifikasi,
                                "versi_dokumen": versi_dokumen,
                            },
                        )
                        if result.get("ok"):
                            st.success("Dokumen berhasil diupload. Folder mahasiswa akan dibuat otomatis di Google Drive.")
                            if result.get("link_file"):
                                st.link_button("Buka file di Google Drive", result["link_file"])
                            clear_cache_and_rerun()
                        else:
                            st.error(result.get("error", "Gagal upload dokumen"))

    with tabs[1]:
        if documents_df.empty:
            st.info("Belum ada dokumen.")
        else:
            docs = documents_df.copy()
            if "tanggal_upload" in docs.columns:
                docs["tanggal_upload"] = docs["tanggal_upload"].astype(str)
            filter_cols = st.columns(3)
            docs["student_display"] = docs.apply(
                lambda r: student_code_name(r.get("student_id"), r.get("nama_mahasiswa")),
                axis=1,
            )

            student_filter = filter_cols[0].selectbox(
                "Filter mahasiswa",
                ["Semua"] + sorted(docs["student_display"].astype(str).unique().tolist())
            )
            jenis_filter = filter_cols[1].selectbox("Filter jenis dokumen", ["Semua"] + sorted(docs["jenis_dokumen"].astype(str).unique().tolist()))
            verify_filter = filter_cols[2].selectbox("Filter status verifikasi", ["Semua"] + sorted(docs["status_verifikasi"].astype(str).unique().tolist()))
            if student_filter != "Semua":
                docs = docs[docs["student_display"].astype(str) == student_filter]
            if jenis_filter != "Semua":
                docs = docs[docs["jenis_dokumen"].astype(str) == jenis_filter]
            if verify_filter != "Semua":
                docs = docs[docs["status_verifikasi"].astype(str) == verify_filter]
            show_cols = [c for c in [
                "doc_id", "student_display", "student_id", "nama_mahasiswa", "jenis_dokumen", "nama_file",
                "tanggal_upload", "uploaded_by", "status_verifikasi", "link_file", "storage_path"
            ] if c in docs.columns]
            st.dataframe(docs[show_cols], use_container_width=True, hide_index=True)
            if "link_file" in docs.columns:
                for _, row in docs.head(10).iterrows():
                    if safe_text(row.get("link_file")):
                        st.markdown(f"- **{safe_text(row.get('nama_file'))}** — [Buka file]({safe_text(row.get('link_file'))})")


# ---------- Invoice & Payments ----------
def render_invoice_module(students_df: pd.DataFrame, invoices_df: pd.DataFrame, payments_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Invoice & Pembayaran")
    tabs = st.tabs([
        "Dashboard Invoice",
        "Buat Paket Invoice",
        "Buat Invoice Manual",
        "Record Pembayaran",
        "Invoice Styled",
    ])

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

            reg_paid = inv[(inv["invoice_type"] == "Pendaftaran") & (inv["status_pelunasan"] == "Lunas")]
            admin_outstanding = inv[(inv["invoice_type"] == "Admin") & (inv["sisa_tagihan"] > 0)]
            m1, m2 = st.columns(2)
            m1.metric("Invoice pendaftaran lunas", len(reg_paid))
            m2.metric("Outstanding admin", format_currency(admin_outstanding["sisa_tagihan"].sum()))

            ch1, ch2 = st.columns(2)
            with ch1:
                pel = inv.groupby("status_pelunasan", dropna=False).size().reset_index(name="jumlah")
                fig_invoice_status = px.pie(
                    pel,
                    names="status_pelunasan",
                    values="jumlah",
                    color_discrete_sequence=["#C2410C", "#F97316", "#FDBA74", "#FED7AA"],
                )
                fig_invoice_status = style_pie_chart(fig_invoice_status, "Status Pelunasan", hole=0.44)
                st.plotly_chart(fig_invoice_status, use_container_width=True, config={"displayModeBar": False})
            with ch2:
                typ = inv.groupby("invoice_type", dropna=False)["sisa_tagihan"].sum().reset_index()
                fig_invoice_type = px.bar(
                    typ,
                    x="invoice_type",
                    y="sisa_tagihan",
                    color="sisa_tagihan",
                    color_continuous_scale=[
                        [0.00, "#FED7AA"],
                        [0.35, "#FDBA74"],
                        [0.70, "#F97316"],
                        [1.00, "#C2410C"],
                    ],
                )
                fig_invoice_type = style_bar_chart(fig_invoice_type, "Outstanding per Jenis Invoice")
                st.plotly_chart(fig_invoice_type, use_container_width=True, config={"displayModeBar": False})
            st.markdown("### Ringkasan keuangan per mahasiswa")
            finance_df = group_student_finance(inv)
            show_finance = finance_df.copy()
            for col in ["total_tagihan", "total_dibayar", "total_outstanding"]:
                if col in show_finance.columns:
                    show_finance[col] = show_finance[col].apply(format_currency)
            st.dataframe(show_finance, use_container_width=True, hide_index=True)

            st.markdown("### Detail invoice")
            invoice_type_options = ["Semua"] + sorted([x for x in inv["invoice_type"].dropna().unique() if x])
            selected_type = st.selectbox("Filter jenis invoice", invoice_type_options, key="invoice_type_filter")
            filtered_inv = inv.copy()
            if selected_type != "Semua":
                filtered_inv = filtered_inv[filtered_inv["invoice_type"] == selected_type]

            show_cols = [c for c in [
                "kode_invoice", "invoice_type", "student_id", "nama_mahasiswa", "tanggal_invoice", "program",
                "harga_program", "sudah_dibayar", "sisa_tagihan", "status_pelunasan", "status_pengiriman"
            ] if c in filtered_inv.columns]
            show_df = filtered_inv[show_cols].copy()
            for money_col in ["harga_program", "sudah_dibayar", "sisa_tagihan"]:
                if money_col in show_df.columns:
                    show_df[money_col] = show_df[money_col].apply(format_currency)
            st.dataframe(show_df, use_container_width=True, hide_index=True)

            if not payments_df.empty:
                st.markdown("### Log pembayaran")
                pay = payments_df.copy()
                if "jumlah_pembayaran" in pay.columns:
                    pay["jumlah_pembayaran"] = pay["jumlah_pembayaran"].apply(format_currency)
                st.dataframe(pay, use_container_width=True, hide_index=True)

    with tabs[1]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            student_options, student_map = build_student_options(students_df)
            selected_student_label = st.selectbox(
                "Pilih mahasiswa untuk paket invoice",
                student_options,
                key="invoice_package_student_id"
            )
            selected_student_id = student_map[selected_student_label]
            student = find_student(students_df, selected_student_id)
            package = calculate_invoice_package(student)

            info1, info2, info3 = st.columns(3)
            info1.metric("Biaya pendaftaran", format_currency(package["registration_fee"]))
            info2.metric("Biaya admin", format_currency(package["admin_invoice_total"]))
            info3.metric("Total kewajiban", format_currency(package["grand_total"]))

            st.info(
                f"Invoice akan di-split menjadi: "
                f"Pendaftaran {format_currency(package['registration_fee'])} "
                f"+ Admin {format_currency(package['admin_invoice_total'])}. "
                f"Total keseluruhan kewajiban mahasiswa: {format_currency(package['grand_total'])}."
            )

            with st.form("form_create_invoice_package"):
                c1, c2 = st.columns(2)
                tanggal_invoice_val = c1.date_input("Tanggal Invoice", value=datetime.now().date(), key="package_tanggal_invoice")
                mata_uang = c2.selectbox("Mata Uang", ["IDR", "USD", "CNY"], key="package_currency")
                status_pengiriman = st.selectbox("Status Pengiriman", refs.get("status_pengiriman", ["Belum Dikirim"]), key="package_status_pengiriman")
                kirim_hari_ini = st.checkbox("Sudah dikirim hari ini", key="package_sent_today")
                tanggal_kirim_val = st.date_input("Tanggal Kirim", value=datetime.now().date(), disabled=not kirim_hari_ini, key="package_tanggal_kirim")
                catatan_invoice = st.text_area(
                    "Catatan Invoice Paket",
                        value="Invoice paket otomatis: Pendaftaran + Admin",
                        key="package_catatan_invoice",
                    )
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
                        created = result.get("created_invoices", [])
                        codes = ", ".join([safe_text(x.get("kode_invoice")) for x in created if safe_text(x.get("kode_invoice"))])
                        st.success(f"Paket invoice berhasil dibuat. {codes}")
                        clear_cache_and_rerun()
                    else:
                        st.error(result.get("error", "Gagal membuat paket invoice"))

    with tabs[2]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            student_options, student_map = build_student_options(students_df)
            selected_student_label = st.selectbox(
                "Pilih mahasiswa untuk invoice manual",
                student_options,
                key="manual_invoice_student_id"
            )
            selected_student_id = student_map[selected_student_label]
            student = find_student(students_df, selected_student_id)
            with st.form("form_create_invoice"):
                c1, c2, c3 = st.columns(3)
                tanggal_invoice_val = c1.date_input("Tanggal Invoice", value=datetime.now().date(), key="manual_tanggal_invoice")
                invoice_type = c2.selectbox("Jenis Invoice", ["Pendaftaran", "Admin", "Manual"], key="manual_invoice_type")
                mata_uang = c3.selectbox("Mata Uang", ["IDR", "USD", "CNY"], key="manual_currency")

                program_options = ensure_option_list(
                    refs.get("program_diminati", refs.get("program", [])),
                    student.get("program_diminati"),
                )
                program = st.selectbox(
                    "Program",
                    program_options,
                    index=option_index(program_options, student.get("program_diminati")),
                    key="manual_program",
                )

                default_nominal = 0.0
                if invoice_type == "Pendaftaran":
                    default_nominal = get_registration_fee(program)
                elif invoice_type == "Admin":
                    package_preview = calculate_invoice_package(
                        {
                            "program_diminati": program,
                            "estimasi_biaya": student.get("estimasi_biaya"),
                        }
                    )
                    default_nominal = package_preview["admin_invoice_total"]

                harga_program = st.number_input(
                    "Nominal Invoice",
                    min_value=0.0,
                    value=float(default_nominal),
                    step=100000.0,
                    key="manual_harga_program",
                )
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
                        st.success(f"Invoice berhasil dibuat: {result.get('kode_invoice')}")
                        clear_cache_and_rerun()
                    else:
                        st.error(result.get("error", "Gagal membuat invoice"))

    with tabs[3]:
        if inv.empty:
            st.info("Belum ada invoice.")
        else:
            invoice_options, invoice_map = build_invoice_options(inv)
            selected_label = st.selectbox("Pilih invoice", invoice_options, key="payment_invoice_label")
            selected_invoice_id = invoice_map[selected_label]
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
            invoice_options, invoice_map = build_invoice_options(inv)
            selected_label = st.selectbox("Pilih invoice styled", invoice_options, key="styled_invoice_label")
            selected_invoice_id = invoice_map[selected_label]
            invoice = inv[inv["invoice_id"].astype(str) == selected_invoice_id].iloc[0].to_dict()
            student = find_student(students_df, safe_text(invoice.get("student_id")))

            preview_url = build_preview_invoice_url(selected_invoice_id)
            pdf_record = invoice_row_for_pdf(invoice, student)
            pdf_bytes = generate_invoice_pdf(pdf_record, PROFILE_FIXED)
            expected_code = expected_invoice_code(invoice.get("tanggal_invoice"), invoice.get("student_id"))

            left, right = st.columns([1, 1])
            with left:
                st.markdown("### Preview")
                st.link_button("Buka Preview Invoice Styled", preview_url, use_container_width=True)

                st.download_button(
                    "Download PDF Invoice",
                    data=pdf_bytes,
                    file_name=invoice_pdf_filename(
                        safe_text(invoice.get("kode_invoice") or invoice.get("invoice_id")),
                        safe_text(invoice.get("nama_mahasiswa") or student.get("nama_lengkap")),
                        safe_text(invoice.get("invoice_type") or "Invoice"),
                    ),
                    mime="application/pdf",
                    use_container_width=True,
                )

                if st.button("Simpan PDF Bagus ke Google Drive", use_container_width=True):
                    try:
                        result = upload_invoice_pdf_to_drive(
                            invoice_id=safe_text(invoice.get("invoice_id")),
                            student_id=safe_text(invoice.get("student_id")),
                            nama_mahasiswa=safe_text(invoice.get("nama_mahasiswa") or student.get("nama_lengkap")),
                            kode_invoice=safe_text(invoice.get("kode_invoice") or invoice.get("invoice_id")),
                            invoice_type=safe_text(invoice.get("invoice_type") or "Invoice"),
                            pdf_bytes=pdf_bytes,
                        )
                    except Exception as exc:
                        st.error(f"Gagal upload PDF ke Drive: {exc}")
                    else:
                        if result.get("ok"):
                            st.success("PDF bagus berhasil disimpan ke Google Drive.")
                            if result.get("file_name"):
                                st.write(f"**File PDF:** {safe_text(result.get('file_name'))}")
                            if result.get("folder_name"):
                                st.write(f"**Folder Drive:** {safe_text(result.get('folder_name'))}")
                            if result.get("file_url"):
                                st.link_button("Buka PDF di Google Drive", result["file_url"], use_container_width=True)
                            if result.get("folder_url"):
                                st.link_button("Buka Folder Invoices", result["folder_url"], use_container_width=True)
                        else:
                            st.error(result.get("error", "Gagal menyimpan PDF ke Drive"))

                st.caption("Preview akan membuka template invoice kanan dari Apps Script.")
                st.info("Untuk hasil paling mirip preview, Anda juga bisa pakai Print > Save as PDF dari halaman preview.")

            with right:
                st.markdown("### Informasi Invoice")
                st.write(f"**Kode Invoice saat ini:** {safe_text(invoice.get('kode_invoice'))}")
                if expected_code:
                    st.write(f"**Format kode yang Anda mau:** {expected_code}")
                st.write(f"**Jenis Invoice:** {safe_text(invoice.get('invoice_type'))}")
                st.write(f"**Nama Mahasiswa:** {safe_text(invoice.get('nama_mahasiswa'))}")
                st.write(f"**Program:** {safe_text(invoice.get('program'))}")
                st.write(f"**Harga Invoice:** {format_currency(invoice.get('harga_program'))}")
                st.write(f"**Sudah Dibayar:** {format_currency(invoice.get('sudah_dibayar'))}")
                st.write(f"**Sisa Tagihan:** {format_currency(invoice.get('sisa_tagihan'))}")
                st.write(f"**Status Pelunasan:** {safe_text(invoice.get('status_pelunasan'))}")
                st.divider()
                st.markdown("### Hapus Invoice")
                st.warning(
                    "Aksi ini akan menghapus data invoice yang dipilih dari Google Sheet. "
                    "Log pembayaran yang terkait invoice ini juga akan dihapus."
                )

                confirm_delete_invoice = st.text_input(
                    f"Ketik invoice_id berikut untuk konfirmasi: {selected_invoice_id}",
                    key=f"confirm_delete_invoice_{selected_invoice_id}",
                )

                if st.button(
                    "Hapus invoice ini",
                    type="primary",
                    use_container_width=True,
                    key=f"delete_invoice_btn_{selected_invoice_id}",
                ):
                    if confirm_delete_invoice.strip() != selected_invoice_id:
                        st.error("Konfirmasi tidak cocok. Invoice belum dihapus.")
                    else:
                        try:
                            result = api_post(
                                "delete_invoice",
                                {
                                    "invoice_id": selected_invoice_id,
                                },
                            )
                        except Exception as exc:
                            st.error(f"Gagal menghubungi Apps Script: {exc}")
                        else:
                            if result.get("ok"):
                                st.success(
                                    f"Invoice berhasil dihapus: "
                                    f"{safe_text(result.get('kode_invoice'))} "
                                    f"({safe_text(result.get('invoice_type'))})"
                                )

                                deleted_payments = result.get("deleted_payments", 0)
                                if deleted_payments:
                                    st.info(f"Log pembayaran terkait yang ikut dihapus: {deleted_payments}")

                                clear_cache_and_rerun()
                            else:
                                st.error(result.get("error", "Gagal menghapus invoice"))

# ---------- SOP ----------
def render_help_module() -> None:
    st.subheader("Bantuan & SOP")
    tabs = st.tabs(["Cara Pakai", "Alur Operasional", "Checklist Harian"])

    with tabs[0]:
        st.markdown(
            """
            ### Cara pakai aplikasi
            1. **Calon Mahasiswa** untuk melihat, menambah, edit, dan update progress.
            2. **Dokumen** untuk upload dokumen ke Google Drive. Folder mahasiswa dibuat otomatis berdasarkan `student_id`.
            3. **Invoice & Pembayaran** untuk membuat invoice, mencatat pembayaran, dan invoice styled.
            4. **Dashboard** untuk memantau pipeline mahasiswa dan kondisi keuangan secara ringkas.
            """
        )
        st.info("Agar upload dokumen otomatis masuk ke folder yang rapi, isi `ROOT_FOLDER_ID` di Apps Script.")

    with tabs[1]:
        st.markdown(
            """
            ### SOP singkat operasional
            **Lead masuk dari GForm**
            - Data masuk ke `Form Responses 1`
            - Trigger Apps Script memindahkan data ke `students_master`
            - Tim assign PIC dan update status proses

            **Dokumen masuk**
            - Pilih mahasiswa di menu Dokumen
            - Upload file
            - Sistem membuat folder mahasiswa otomatis di Google Drive
            - Status verifikasi bisa diubah dari metadata dokumen

            **Invoice & pembayaran**
            - Buat paket invoice otomatis dari menu Invoice & Pembayaran
            - Sistem membuat 2 invoice: **Pendaftaran** dan **Admin**
            - Invoice admin otomatis memuat biaya admin + transport Rp 4.000.000
            - Saat pembayaran diterima, catat di menu Record Pembayaran
            - Status invoice dihitung per invoice, sedangkan dashboard keuangan tetap menampilkan outstanding total per mahasiswa
            - Preview dan PDF styled diambil dari Apps Script
            """
        )

    with tabs[2]:
        st.markdown(
            """
            ### Checklist admin harian
            - Cek lead baru dari GForm
            - Update PIC dan status proses
            - Follow up dokumen yang belum lengkap
            - Upload dokumen yang diterima ke folder student
            - Buat paket invoice untuk student yang siap pembayaran
            - Catat pembayaran yang masuk per invoice
            - Review dashboard outstanding admin dan ringkasan keuangan per mahasiswa
            """
        )


# ---------- Main ----------
def main() -> None:
    inject_ui_style()
    render_top_header()

    try:
        data = load_bootstrap()
    except Exception as exc:
        st.error(f"Gagal memuat data awal: {exc}")
        st.stop()

    students_df = normalize_df(as_df(data.get("students", [])))
    tracking_df = normalize_df(as_df(data.get("student_tracking", data.get("tracking", []))))
    documents_df = normalize_df(as_df(data.get("documents", [])))
    invoices_df = normalize_df(as_df(data.get("invoices", [])))
    payments_df = normalize_df(as_df(data.get("payments", [])))
    refs = data.get("references", {}) or {}

    new_students = detect_new_students(students_df)

    if new_students:
        if len(new_students) == 1:
            st.toast(f"Mahasiswa baru masuk: {new_students[0]['nama_lengkap']}")
        else:
            st.toast(f"Ada {len(new_students)} calon mahasiswa baru masuk")

    if "sidebar_page" not in st.session_state:
        st.session_state["sidebar_page"] = "Dashboard"

    PAGES = [
        "Dashboard",
        "Calon Mahasiswa",
        "Student Tracking",
        "Dokumen",
        "Invoice & Pembayaran",
        "Bantuan & SOP",
    ]

    if "sidebar_page" not in st.session_state:
        st.session_state["sidebar_page"] = "Dashboard"

    if "pending_page" in st.session_state:
        st.session_state["sidebar_page"] = st.session_state.pop("pending_page")

    with st.sidebar:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=120)

        st.markdown("## Nihaoma")
        st.caption("Education Center")

        page = st.radio(
            "Pilih Menu",
            PAGES,
            key="sidebar_page",
        )

        if HERO_STUDENT_PATH.exists():
            st.image(str(HERO_STUDENT_PATH), width=220)

        if st.button("Refresh data", use_container_width=True):
            clear_cache_and_rerun()

        latest_new_students = st.session_state.get("latest_new_students", [])

        if latest_new_students:
            st.markdown(f"### Notifikasi ({len(latest_new_students)})")
            for item in latest_new_students[:5]:
                st.caption(f"• {item['nama_lengkap']} ({item['student_id']})")

        st.caption(f"Data terakhir dimuat: {safe_text(data.get('meta', {}).get('generated_at'))}")


    if page == "Dashboard":
        render_dashboard(students_df, invoices_df, payments_df)
    elif page == "Calon Mahasiswa":
        render_student_list(students_df, refs)
    elif page == "Student Tracking":
        render_student_tracking_module(students_df, tracking_df, refs)
    elif page == "Dokumen":
        render_documents_module(students_df, documents_df, refs)
    elif page == "Invoice & Pembayaran":
        render_invoice_module(students_df, invoices_df, payments_df, refs)
    elif page == "Bantuan & SOP":
        render_help_module()


if __name__ == "__main__":
    main()
