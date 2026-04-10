import base64
import io
import os
from datetime import datetime
from typing import Any, Dict, List

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

st.set_page_config(page_title="Nihaoma Student Operations", layout="wide")

import streamlit as st
import hmac

import streamlit as st
import pandas as pd
import requests
import hmac

# function login di sini...

check_login()

st.title("Nihaoma Student Operations")

# isi app Anda mulai dari sini

with st.sidebar:
    st.success(f"Login sebagai: {st.session_state.get('auth_user', '-')}")
    st.button("Logout", on_click=logout, use_container_width=True)

    st.markdown("### Menu")
    page = st.radio("Pilih menu", ["Dashboard", "Calon Mahasiswa", "Dokumen", "Invoice & Pembayaran"])
def logout():
    st.session_state["auth_ok"] = False
    st.session_state["auth_user"] = None
    st.session_state["login_password"] = ""
    st.rerun()


def check_login():
    if "auth_ok" not in st.session_state:
        st.session_state["auth_ok"] = False
    if "auth_user" not in st.session_state:
        st.session_state["auth_user"] = None
    if "auth_error" not in st.session_state:
        st.session_state["auth_error"] = ""

    users = st.secrets.get("APP_USERS", {})

    def try_login():
        username = str(st.session_state.get("login_username", "")).strip()
        password = str(st.session_state.get("login_password", ""))

        saved_password = users.get(username)
        if saved_password and hmac.compare_digest(password, str(saved_password)):
            st.session_state["auth_ok"] = True
            st.session_state["auth_user"] = username
            st.session_state["auth_error"] = ""
            st.session_state["login_password"] = ""
        else:
            st.session_state["auth_ok"] = False
            st.session_state["auth_user"] = None
            st.session_state["auth_error"] = "Username atau password salah."
            st.session_state["login_password"] = ""

    if not st.session_state["auth_ok"]:
        st.markdown("## Login Nihaoma Student Operations")
        st.caption("Masuk dengan username dan password admin.")

        if st.session_state["auth_error"]:
            st.error(st.session_state["auth_error"])

        with st.form("login_form"):
            st.text_input("Username", key="login_username")
            st.text_input("Password", type="password", key="login_password")
            st.form_submit_button("Masuk", on_click=try_login)

        st.stop()

SCRIPT_URL = (
    st.secrets.get("SCRIPT_URL")
    or st.secrets.get("APPS_SCRIPT_URL")
    or os.getenv("SCRIPT_URL")
    or os.getenv("APPS_SCRIPT_URL")
    or ""
)
WRITE_TOKEN = st.secrets.get("WRITE_TOKEN", os.getenv("WRITE_TOKEN", ""))
TIMEOUT = 90


# ---------- Core helpers ----------
def ensure_config() -> None:
    if not SCRIPT_URL or not WRITE_TOKEN:
        st.error("SCRIPT_URL atau WRITE_TOKEN belum diisi di secrets / environment.")
        st.stop()



def api_get(action: str) -> Dict[str, Any]:
    ensure_config()
    resp = requests.get(
        SCRIPT_URL,
        params={"action": action, "token": WRITE_TOKEN},
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


@st.cache_data(ttl=60, show_spinner=False)
def load_bootstrap() -> Dict[str, Any]:
    result = api_get("bootstrap")
    if not result.get("ok"):
        raise RuntimeError(result.get("error", "Gagal memuat data awal"))
    return result



def clear_cache_and_rerun() -> None:
    st.cache_data.clear()
    st.rerun()


# ---------- Formatting helpers ----------
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


# ---------- PDF ----------
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

    story.append(Paragraph("<b>NIHAOMA STUDENT OPERATIONS</b>", styles["Title"]))
    story.append(Paragraph("Invoice Pembayaran", styles["Heading2"]))
    story.append(Spacer(1, 8))

    info_data = [
        ["Kode Invoice", safe_text(invoice.get("kode_invoice"))],
        ["Tanggal Invoice", maybe_date(invoice.get("tanggal_invoice"))],
        ["Student ID", safe_text(invoice.get("student_id"))],
        ["Nama Mahasiswa", safe_text(invoice.get("nama_mahasiswa") or student.get("nama_lengkap"))],
        ["Program", safe_text(invoice.get("program") or student.get("program_diminati"))],
        ["Intake", safe_text(student.get("intake"))],
        ["Kampus Tujuan", safe_text(student.get("kampus_tujuan"))],
    ]
    info_table = Table(info_data, colWidths=[45 * mm, 120 * mm])
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

    bill_data = [
        ["Deskripsi", "Nilai"],
        [safe_text(invoice.get("deskripsi_biaya") or "Biaya program"), format_currency(harga)],
        ["Sudah Dibayar", format_currency(dibayar)],
        ["Sisa Tagihan", format_currency(sisa)],
        ["Status Pelunasan", safe_text(invoice.get("status_pelunasan"))],
    ]
    bill_table = Table(bill_data, colWidths=[110 * mm, 55 * mm])
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


# ---------- Dashboard ----------
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
            fig_status = px.pie(status_df, names="status_proses", values="jumlah", hole=0.35)
            st.plotly_chart(fig_status, use_container_width=True)

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
            fig_pic = px.pie(pic_df, names="pic_admin", values="jumlah")
            st.plotly_chart(fig_pic, use_container_width=True)

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
            fig_pelunasan = px.pie(pelunasan, names="status_pelunasan", values="jumlah")
            st.plotly_chart(fig_pelunasan, use_container_width=True)

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
            fig_outstanding = px.bar(outstanding_df, x="program", y="sisa_tagihan")
            st.plotly_chart(fig_outstanding, use_container_width=True)


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
                display_df["estimasi_biaya"] = display_df["estimasi_biaya"].apply(format_currency)
            st.dataframe(display_df, use_container_width=True, hide_index=True)
            st.caption(f"Total data tampil: {len(filtered)}")

            student_options = filtered["student_id"].astype(str).tolist() if "student_id" in filtered.columns else []
            if student_options:
                selected_id = st.selectbox("Pilih student_id untuk aksi", student_options, key="student_action_id")
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
                                st.success(f"Data {delete_id} berhasil dihapus.")
                                st.session_state.pop("delete_student_id", None)
                                st.session_state.pop("confirm_delete_text", None)
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
            detail_options = students_df["student_id"].astype(str).tolist()
            selected_detail_id = st.selectbox("Pilih mahasiswa", detail_options, key="detail_student_id")
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
            student_ids = students_df["student_id"].astype(str).tolist()
            selected_student_id = st.selectbox("Pilih mahasiswa", student_ids, key="doc_student_id")
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
                                "nama_file": file.name,
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
            student_filter = filter_cols[0].selectbox("Filter student_id", ["Semua"] + sorted(docs["student_id"].astype(str).unique().tolist()))
            jenis_filter = filter_cols[1].selectbox("Filter jenis dokumen", ["Semua"] + sorted(docs["jenis_dokumen"].astype(str).unique().tolist()))
            verify_filter = filter_cols[2].selectbox("Filter status verifikasi", ["Semua"] + sorted(docs["status_verifikasi"].astype(str).unique().tolist()))
            if student_filter != "Semua":
                docs = docs[docs["student_id"].astype(str) == student_filter]
            if jenis_filter != "Semua":
                docs = docs[docs["jenis_dokumen"].astype(str) == jenis_filter]
            if verify_filter != "Semua":
                docs = docs[docs["status_verifikasi"].astype(str) == verify_filter]
            show_cols = [c for c in [
                "doc_id", "student_id", "nama_mahasiswa", "jenis_dokumen", "nama_file",
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
    tabs = st.tabs(["Dashboard Invoice", "Buat Invoice", "Record Pembayaran", "Download PDF Invoice"])

    with tabs[0]:
        if invoices_df.empty:
            st.info("Belum ada invoice.")
        else:
            inv = invoices_df.copy()
            inv["harga_program"] = inv.get("harga_program", 0).apply(to_number)
            inv["sudah_dibayar"] = inv.get("sudah_dibayar", 0).apply(to_number)
            inv["sisa_tagihan"] = inv.get("sisa_tagihan", 0).apply(to_number)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Invoice", len(inv))
            c2.metric("Total Nilai", format_currency(inv["harga_program"].sum()))
            c3.metric("Sudah Dibayar", format_currency(inv["sudah_dibayar"].sum()))
            c4.metric("Outstanding", format_currency(inv["sisa_tagihan"].sum()))

            ch1, ch2 = st.columns(2)
            with ch1:
                pel = inv.groupby("status_pelunasan", dropna=False).size().reset_index(name="jumlah")
                st.plotly_chart(px.pie(pel, names="status_pelunasan", values="jumlah"), use_container_width=True)
            with ch2:
                peng = inv.groupby("status_pengiriman", dropna=False).size().reset_index(name="jumlah")
                st.plotly_chart(px.bar(peng, x="status_pengiriman", y="jumlah"), use_container_width=True)

            show_cols = [c for c in [
                "kode_invoice", "student_id", "nama_mahasiswa", "tanggal_invoice", "program",
                "harga_program", "sudah_dibayar", "sisa_tagihan", "status_pelunasan", "status_pengiriman"
            ] if c in inv.columns]
            show_df = inv[show_cols].copy()
            for money_col in ["harga_program", "sudah_dibayar", "sisa_tagihan"]:
                if money_col in show_df.columns:
                    show_df[money_col] = show_df[money_col].apply(format_currency)
            st.dataframe(show_df, use_container_width=True, hide_index=True)

            if not payments_df.empty:
                st.markdown("### Log Pembayaran")
                pay = payments_df.copy()
                if "jumlah_pembayaran" in pay.columns:
                    pay["jumlah_pembayaran"] = pay["jumlah_pembayaran"].apply(format_currency)
                st.dataframe(pay, use_container_width=True, hide_index=True)

    with tabs[1]:
        if students_df.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            student_ids = students_df["student_id"].astype(str).tolist()
            selected_student_id = st.selectbox("Pilih mahasiswa untuk invoice", student_ids, key="invoice_student_id")
            student = find_student(students_df, selected_student_id)
            program_prices = refs.get("program_prices", [])
            suggested_price = 0
            for row in program_prices:
                if safe_text(row.get("program_diminati")) == safe_text(student.get("program_diminati")):
                    suggested_price = to_number(row.get("estimasi_biaya"))
                    break
            with st.form("form_create_invoice"):
                c1, c2, c3 = st.columns(3)
                tanggal_invoice = c1.text_input("Tanggal Invoice", value=str(datetime.now().date()))
                program = c2.text_input("Program", value=safe_text(student.get("program_diminati")))
                mata_uang = c3.selectbox("Mata Uang", ["IDR", "USD", "CNY"])
                harga_program = st.number_input("Harga Program", min_value=0.0, value=float(suggested_price), step=100000.0)
                deskripsi_biaya = st.text_area("Deskripsi Biaya", value=f"Biaya program {safe_text(student.get('program_diminati'))}")
                status_pengiriman = st.selectbox("Status Pengiriman", refs.get("status_pengiriman", ["Belum Dikirim"]))
                tanggal_kirim = st.text_input("Tanggal Kirim", value="")
                catatan_invoice = st.text_area("Catatan Invoice")
                if st.form_submit_button("Buat Invoice"):
                    result = api_post(
                        "create_invoice",
                        {
                            "student_id": selected_student_id,
                            "nama_mahasiswa": safe_text(student.get("nama_lengkap")),
                            "tanggal_invoice": tanggal_invoice,
                            "program": program,
                            "deskripsi_biaya": deskripsi_biaya,
                            "mata_uang": mata_uang,
                            "harga_program": harga_program,
                            "status_pengiriman": status_pengiriman,
                            "tanggal_kirim": tanggal_kirim,
                            "catatan_invoice": catatan_invoice,
                        },
                    )
                    if result.get("ok"):
                        st.success(f"Invoice berhasil dibuat: {result.get('kode_invoice')}")
                        clear_cache_and_rerun()
                    else:
                        st.error(result.get("error", "Gagal membuat invoice"))

    with tabs[2]:
        if invoices_df.empty:
            st.info("Belum ada invoice.")
        else:
            invoice_options = [
                f"{safe_text(row.get('invoice_id'))} | {safe_text(row.get('kode_invoice'))} | {safe_text(row.get('nama_mahasiswa'))}"
                for _, row in invoices_df.iterrows()
            ]
            selected_label = st.selectbox("Pilih invoice", invoice_options, key="payment_invoice_label")
            selected_invoice_id = selected_label.split("|")[0].strip()
            invoice_row = invoices_df[invoices_df["invoice_id"].astype(str) == selected_invoice_id].iloc[0].to_dict()
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

    with tabs[3]:
        if invoices_df.empty:
            st.info("Belum ada invoice.")
        else:
            invoice_options = [
                f"{safe_text(row.get('invoice_id'))} | {safe_text(row.get('kode_invoice'))} | {safe_text(row.get('nama_mahasiswa'))}"
                for _, row in invoices_df.iterrows()
            ]
            selected_label = st.selectbox("Pilih invoice untuk PDF", invoice_options, key="pdf_invoice_label")
            selected_invoice_id = selected_label.split("|")[0].strip()
            invoice = invoices_df[invoices_df["invoice_id"].astype(str) == selected_invoice_id].iloc[0].to_dict()
            student = find_student(students_df, safe_text(invoice.get("student_id")))

            preview_cols = st.columns(2)
            with preview_cols[0]:
                st.write(f"**Kode Invoice:** {safe_text(invoice.get('kode_invoice'))}")
                st.write(f"**Nama Mahasiswa:** {safe_text(invoice.get('nama_mahasiswa'))}")
                st.write(f"**Program:** {safe_text(invoice.get('program'))}")
            with preview_cols[1]:
                st.write(f"**Harga Program:** {format_currency(invoice.get('harga_program'))}")
                st.write(f"**Sudah Dibayar:** {format_currency(invoice.get('sudah_dibayar'))}")
                st.write(f"**Sisa Tagihan:** {format_currency(invoice.get('sisa_tagihan'))}")
            pdf_bytes = build_invoice_pdf(invoice, student)
            st.download_button(
                "Download PDF Invoice",
                data=pdf_bytes,
                file_name=f"invoice_{safe_text(invoice.get('kode_invoice') or invoice.get('invoice_id'))}.pdf",
                mime="application/pdf",
            )


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
            3. **Invoice & Pembayaran** untuk membuat invoice, mencatat pembayaran, dan download PDF invoice.
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
            - Buat invoice dari menu Invoice & Pembayaran
            - Download PDF untuk dikirim ke mahasiswa
            - Saat pembayaran diterima, catat di menu Record Pembayaran
            - Sistem akan mengupdate `sudah_dibayar`, `sisa_tagihan`, dan `status_pelunasan`
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
            - Buat invoice untuk student yang siap pembayaran
            - Catat pembayaran yang masuk
            - Review dashboard outstanding invoice
            """
        )


# ---------- Main ----------
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
        page = st.radio(
            "",
            ["Dashboard", "Calon Mahasiswa", "Dokumen", "Invoice & Pembayaran", "Bantuan & SOP"],
            label_visibility="collapsed",
        )
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
