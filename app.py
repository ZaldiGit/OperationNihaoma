from __future__ import annotations

import base64
import html
import io
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
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

st.set_page_config(page_title="Nihaoma Student Operations System", layout="wide")

APP_TITLE = "Nihaoma Student Operations System"
LOGO_PATH = Path(__file__).parent / "logo-nihaoma-rounded.png"
APPROVAL_PATH = Path(__file__).parent / "approval-composite.png"

BLUE = colors.HexColor("#1E88E5")
BLUE_DARK = colors.HexColor("#0F4C81")
ORANGE = colors.HexColor("#F59E0B")
ORANGE_DARK = colors.HexColor("#D97706")
TEXT_DARK = colors.HexColor("#1F2937")
MUTED = colors.HexColor("#6B7280")
LINE = colors.HexColor("#D6DCE5")
LIGHT_BLUE = colors.HexColor("#EEF6FF")
LIGHT_ORANGE = colors.HexColor("#FFF6E8")
WHITE = colors.white


def require_secrets() -> tuple[str, str]:
    try:
        endpoint = st.secrets["APPS_SCRIPT_URL"]
        token = st.secrets["WRITE_TOKEN"]
        return endpoint, token
    except Exception:
        st.error("Secrets APPS_SCRIPT_URL dan WRITE_TOKEN belum diisi di Streamlit.")
        st.stop()


def format_rupiah(value: float) -> str:
    return f"Rp {float(value or 0):,.0f}".replace(",", ".")


def format_date_id(value) -> str:
    if value is None or value == "" or pd.isna(value):
        return "-"
    ts = pd.to_datetime(value, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return "-"
    return ts.strftime("%d/%m/%Y")


def safe_paragraph_text(value) -> str:
    text = "" if value is None else str(value)
    return html.escape(text).replace("\n", "<br/>")


def normalize_frames(payload: dict[str, Any]) -> dict[str, Any]:
    students = pd.DataFrame(payload.get("students", []))
    progress = pd.DataFrame(payload.get("progress", []))
    documents = pd.DataFrame(payload.get("documents", []))
    invoices = pd.DataFrame(payload.get("invoices", []))
    payments = pd.DataFrame(payload.get("payments", []))

    for df, date_cols, num_cols in [
        (students, ["tanggal_input", "tanggal_lahir", "tanggal_follow_up_terakhir", "tanggal_next_action"], ["estimasi_biaya"]),
        (progress, ["tanggal_update", "tanggal_next_action"], []),
        (documents, ["tanggal_upload", "tanggal_verifikasi"], []),
        (invoices, ["tanggal_invoice", "tanggal_kirim"], ["harga_program", "sudah_dibayar", "sisa_tagihan"]),
        (payments, ["tanggal_pembayaran"], ["jumlah_pembayaran"]),
    ]:
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return {
        "students": students.fillna(""),
        "progress": progress.fillna(""),
        "documents": documents.fillna(""),
        "invoices": invoices.fillna(""),
        "payments": payments.fillna(""),
        "references": payload.get("references", {}),
        "meta": payload.get("meta", {}),
    }


@st.cache_data(ttl=60, show_spinner=False)
def fetch_bootstrap(endpoint: str, token: str) -> dict[str, Any]:
    res = requests.get(endpoint, params={"action": "bootstrap", "token": token}, timeout=60)
    res.raise_for_status()
    payload = res.json()
    if not payload.get("ok"):
        raise RuntimeError(payload.get("error", "Gagal membaca data dari Google Sheets"))
    return normalize_frames(payload)


def post_action(endpoint: str, token: str, action: str, data: dict[str, Any]) -> dict[str, Any]:
    payload = {"action": action, "token": token, **data}
    res = requests.post(endpoint, json=payload, timeout=120)
    res.raise_for_status()
    out = res.json()
    if not out.get("ok"):
        raise RuntimeError(out.get("error", f"Gagal action {action}"))
    fetch_bootstrap.clear()
    return out


def render_header(title: str, subtitle: str = "") -> None:
    c1, c2 = st.columns([1, 6])
    with c1:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=95)
    with c2:
        st.markdown(f"## {title}")
        if subtitle:
            st.write(subtitle)


def dashboard_page(data: dict[str, Any]) -> None:
    render_header(
        "Selamat Datang di Sistem Operasional Nihaoma",
        "Versi live: semua simpanan data langsung masuk ke Google Sheet.",
    )
    meta = data["meta"]
    st.caption(f"Spreadsheet: {meta.get('spreadsheet_name','-')} | Last sync: {meta.get('generated_at','-')}")

    students = data["students"]
    documents = data["documents"]
    invoices = data["invoices"]

    total_students = len(students)
    active_students = int((students["is_active"].astype(str).str.upper() == "TRUE").sum()) if not students.empty and "is_active" in students else 0
    total_invoices = len(invoices)
    outstanding = float(invoices["sisa_tagihan"].sum()) if not invoices.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Calon Mahasiswa", total_students)
    c2.metric("Mahasiswa Aktif", active_students)
    c3.metric("Total Invoice", total_invoices)
    c4.metric("Outstanding", format_rupiah(outstanding))

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### Status Proses Mahasiswa")
        if students.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            st.bar_chart(students.groupby("status_proses")["student_id"].count().sort_values(ascending=False))
    with col2:
        st.markdown("### Outstanding per Program")
        if invoices.empty:
            st.info("Belum ada data invoice.")
        else:
            st.bar_chart(invoices.groupby("program")["sisa_tagihan"].sum().sort_values(ascending=False))

    checklist_rows = []
    required_docs = data["references"].get("required_doc_types", [])
    if not students.empty:
        for _, stu in students.iterrows():
            sid = str(stu["student_id"])
            stu_docs = documents[documents["student_id"] == sid]["jenis_dokumen"].astype(str).tolist() if not documents.empty else []
            missing = len([x for x in required_docs if x not in stu_docs])
            checklist_rows.append({"student_id": sid, "nama_lengkap": stu.get("nama_lengkap", ""), "dokumen_belum_lengkap": missing})
    checklist_df = pd.DataFrame(checklist_rows)

    col3, col4 = st.columns(2)
    with col3:
        st.markdown("### Dokumen Belum Lengkap")
        if checklist_df.empty:
            st.info("Belum ada data.")
        else:
            st.dataframe(checklist_df.sort_values("dokumen_belum_lengkap", ascending=False).head(10), use_container_width=True, hide_index=True)
    with col4:
        st.markdown("### Invoice Belum Lunas")
        if invoices.empty:
            st.info("Belum ada data.")
        else:
            show = invoices[invoices["status_pelunasan"] != "Lunas"][["kode_invoice", "nama_mahasiswa", "sisa_tagihan", "status_pelunasan"]].copy()
            if show.empty:
                st.success("Semua invoice sudah lunas.")
            else:
                show["sisa_tagihan"] = show["sisa_tagihan"].map(format_rupiah)
                st.dataframe(show.head(10), use_container_width=True, hide_index=True)

    with st.expander("Validasi Data", expanded=False):
        dup = students[students["student_id"].duplicated(keep=False)] if not students.empty else pd.DataFrame()
        missing_contacts = students[(students["email"].astype(str) == "") | (students["no_whatsapp"].astype(str) == "")] if not students.empty else pd.DataFrame()
        overpaid = invoices[invoices["sudah_dibayar"] > invoices["harga_program"]] if not invoices.empty else pd.DataFrame()
        a, b, c = st.columns(3)
        a.metric("Student ID duplikat", len(dup))
        b.metric("Kontak belum lengkap", len(missing_contacts))
        c.metric("Invoice overpaid", len(overpaid))


def students_page(data: dict[str, Any], endpoint: str, token: str) -> None:
    render_header("Modul Calon Mahasiswa", "Data tersimpan langsung ke Google Sheet live.")
    tabs = st.tabs(["Daftar Mahasiswa", "Tambah Data", "Detail & Progress"])

    students = data["students"]
    progress = data["progress"]
    refs = data["references"]
    prices = refs.get("program_prices", [])
    price_map = {item["program_diminati"]: float(item["estimasi_biaya"] or 0) for item in prices}

    with tabs[0]:
        search = st.text_input("Cari mahasiswa")
        col1, col2, col3 = st.columns(3)
        selected_program = col1.multiselect("Program", sorted([p for p in students.get("program_diminati", pd.Series(dtype=str)).astype(str).unique().tolist() if p]), default=None)
        selected_status = col2.multiselect("Status Proses", refs.get("status_proses", []), default=None)
        selected_pic = col3.multiselect("PIC", sorted([p for p in students.get("pic_admin", pd.Series(dtype=str)).astype(str).unique().tolist() if p]), default=None)

        show = students.copy()
        if not show.empty:
            if search:
                kw = search.lower()
                show = show[
                    show["nama_lengkap"].astype(str).str.lower().str.contains(kw, na=False)
                    | show["student_id"].astype(str).str.lower().str.contains(kw, na=False)
                    | show["email"].astype(str).str.lower().str.contains(kw, na=False)
                ]
            if selected_program:
                show = show[show["program_diminati"].isin(selected_program)]
            if selected_status:
                show = show[show["status_proses"].isin(selected_status)]
            if selected_pic:
                show = show[show["pic_admin"].isin(selected_pic)]
            display = show[["student_id", "nama_lengkap", "program_diminati", "estimasi_biaya", "intake", "pic_admin", "status_proses", "tanggal_input"]].copy()
            display["estimasi_biaya"] = display["estimasi_biaya"].map(format_rupiah)
            display["tanggal_input"] = display["tanggal_input"].apply(format_date_id)
            st.dataframe(display, use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada data mahasiswa.")

    with tabs[1]:
        with st.form("student_live_form"):
            c1, c2, c3 = st.columns(3)
            nama_lengkap = c1.text_input("Nama Lengkap *")
            nama_panggilan = c2.text_input("Nama Panggilan")
            jenis_kelamin = c3.selectbox("Jenis Kelamin", ["", "Pria", "Wanita"])
            c4, c5, c6 = st.columns(3)
            tanggal_lahir = c4.date_input("Tanggal Lahir", value=None)
            kewarganegaraan = c5.text_input("Kewarganegaraan", value="Indonesia")
            no_whatsapp = c6.text_input("No. WhatsApp")
            c7, c8, c9 = st.columns(3)
            email = c7.text_input("Email")
            alamat = c8.text_input("Alamat")
            no_paspor_atau_nik = c9.text_input("No. Paspor / NIK")
            c10, c11, c12 = st.columns(3)
            program_options = [p["program_diminati"] for p in prices] if prices else [""]
            program_diminati = c10.selectbox("Program Diminati", program_options)
            kampus_tujuan = c11.text_input("Kampus Tujuan")
            kota_tujuan = c12.text_input("Kota Tujuan")
            c13, c14, c15 = st.columns(3)
            negara_tujuan = c13.text_input("Negara Tujuan", value="China")
            intake = c14.text_input("Intake")
            durasi_program = c15.text_input("Durasi Program")
            estimasi_biaya = price_map.get(program_diminati, 0.0)
            st.info(f"Estimasi biaya program: {format_rupiah(estimasi_biaya)}")
            c16, c17, c18 = st.columns(3)
            sumber_leads = c16.text_input("Sumber Leads")
            pic_admin = c17.text_input("PIC Admin")
            status_proses = c18.selectbox("Status Proses", refs.get("status_proses", []))
            c19, c20, c21 = st.columns(3)
            next_action = c19.selectbox("Next Action", refs.get("next_action", [""]))
            tanggal_next_action = c20.date_input("Tanggal Next Action", value=None)
            prioritas = c21.selectbox("Prioritas", refs.get("prioritas", []))
            catatan_admin = st.text_area("Catatan Admin")
            is_active = st.checkbox("Aktif", value=True)
            submit = st.form_submit_button("Simpan ke Google Sheet")
        if submit:
            if not nama_lengkap.strip():
                st.error("Nama Lengkap wajib diisi.")
            else:
                try:
                    post_action(endpoint, token, "add_student", {
                        "nama_lengkap": nama_lengkap,
                        "nama_panggilan": nama_panggilan,
                        "jenis_kelamin": jenis_kelamin,
                        "tanggal_lahir": tanggal_lahir.isoformat() if tanggal_lahir else "",
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
                        "durasi_program": durasi_program,
                        "sumber_leads": sumber_leads,
                        "pic_admin": pic_admin,
                        "status_proses": status_proses,
                        "next_action": next_action,
                        "tanggal_next_action": tanggal_next_action.isoformat() if tanggal_next_action else "",
                        "prioritas": prioritas,
                        "catatan_admin": catatan_admin,
                        "is_active": str(bool(is_active)).upper(),
                    })
                    st.success("Data mahasiswa berhasil masuk ke Google Sheet.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan: {e}")

    with tabs[2]:
        if students.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            options = students["student_id"].astype(str) + " - " + students["nama_lengkap"].astype(str)
            selected = st.selectbox("Pilih Mahasiswa", options)
            selected_id = selected.split(" - ", 1)[0]
            row = students[students["student_id"] == selected_id].iloc[0]
            st.markdown(f"### {row['nama_lengkap']}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Student ID", row["student_id"])
            c2.metric("Program", row["program_diminati"] or "-")
            c3.metric("Status", row["status_proses"] or "-")
            c4.metric("PIC", row["pic_admin"] or "-")

            t1, t2, t3 = st.tabs(["Profil", "Update Progress", "Riwayat"])
            with t1:
                show = pd.DataFrame({
                    "Field": ["Nama Lengkap", "WhatsApp", "Email", "Program", "Estimasi Biaya", "Next Action", "Tanggal Next Action", "Catatan"],
                    "Value": [
                        row["nama_lengkap"], row["no_whatsapp"], row["email"], row["program_diminati"],
                        format_rupiah(row["estimasi_biaya"]), row["next_action"], format_date_id(row["tanggal_next_action"]), row["catatan_admin"]
                    ],
                })
                st.dataframe(show, use_container_width=True, hide_index=True)

            with t2:
                with st.form("progress_update_form"):
                    status_options = refs.get("status_proses", [])
                    next_actions = refs.get("next_action", [""])
                    current_status = row["status_proses"] if row["status_proses"] in status_options else status_options[0]
                    current_next = row["next_action"] if row["next_action"] in next_actions else next_actions[0]
                    status_baru = st.selectbox("Status Baru", status_options, index=status_options.index(current_status))
                    next_action_baru = st.selectbox("Next Action", next_actions, index=next_actions.index(current_next))
                    tanggal_next = st.date_input("Tanggal Next Action", value=row["tanggal_next_action"].date() if pd.notna(row["tanggal_next_action"]) else None)
                    note = st.text_area("Catatan Update")
                    updated_by = st.text_input("Updated By", value=str(row["pic_admin"] or "Admin"))
                    save_update = st.form_submit_button("Simpan Update")
                if save_update:
                    try:
                        post_action(endpoint, token, "update_progress", {
                            "student_id": selected_id,
                            "status_baru": status_baru,
                            "updated_by": updated_by,
                            "catatan": note,
                            "next_action": next_action_baru,
                            "tanggal_next_action": tanggal_next.isoformat() if tanggal_next else "",
                        })
                        st.success("Progress berhasil diperbarui di Google Sheet.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal update progress: {e}")

            with t3:
                logs = progress[progress["student_id"] == selected_id].copy()
                if logs.empty:
                    st.info("Belum ada riwayat.")
                else:
                    logs["tanggal_update"] = logs["tanggal_update"].apply(format_date_id)
                    st.dataframe(logs[["tanggal_update", "updated_by", "status_lama", "status_baru", "catatan", "next_action"]], use_container_width=True, hide_index=True)


def documents_page(data: dict[str, Any], endpoint: str, token: str) -> None:
    render_header("Modul Dokumen", "Upload dokumen ke Drive dan catat metadata langsung di Google Sheet.")
    tabs = st.tabs(["Upload Dokumen", "Checklist", "Daftar & Download"])
    students = data["students"]
    documents = data["documents"]
    refs = data["references"]

    with tabs[0]:
        if students.empty:
            st.info("Tambahkan data mahasiswa terlebih dahulu.")
        else:
            student_options = students["student_id"].astype(str) + " - " + students["nama_lengkap"].astype(str)
            with st.form("doc_upload_live_form"):
                selected = st.selectbox("Pilih Mahasiswa", student_options)
                selected_id = selected.split(" - ", 1)[0]
                selected_name = selected.split(" - ", 1)[1]
                jenis_dokumen = st.selectbox("Jenis Dokumen", refs.get("required_doc_types", []) + ["Form Aplikasi", "Dokumen Visa", "Lainnya"])
                versi = st.text_input("Versi Dokumen", value="v1")
                uploaded_by = st.text_input("Uploaded By", value="Admin")
                status_verifikasi = st.selectbox("Status Verifikasi", refs.get("status_verifikasi", []))
                catatan = st.text_area("Catatan")
                file = st.file_uploader("Upload File")
                submit = st.form_submit_button("Upload ke Google Drive & Simpan Metadata")
            if submit:
                if file is None:
                    st.error("File wajib dipilih.")
                else:
                    try:
                        post_action(endpoint, token, "upload_document", {
                            "student_id": selected_id,
                            "nama_mahasiswa": selected_name,
                            "jenis_dokumen": jenis_dokumen,
                            "nama_file": file.name,
                            "mime_type": file.type or "application/octet-stream",
                            "file_base64": base64.b64encode(file.getvalue()).decode("utf-8"),
                            "uploaded_by": uploaded_by,
                            "status_verifikasi": status_verifikasi,
                            "catatan_verifikasi": catatan,
                            "versi_dokumen": versi,
                        })
                        st.success("Dokumen berhasil diupload ke Drive dan dicatat di Google Sheet.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal upload dokumen: {e}")

    with tabs[1]:
        if students.empty:
            st.info("Belum ada data mahasiswa.")
        else:
            rows = []
            req = refs.get("required_doc_types", [])
            for _, stu in students.iterrows():
                sid = str(stu["student_id"])
                stu_docs = documents[documents["student_id"] == sid]["jenis_dokumen"].astype(str).tolist() if not documents.empty else []
                row = {"student_id": sid, "nama_mahasiswa": stu["nama_lengkap"]}
                for dtype in req:
                    row[dtype] = "Ada" if dtype in stu_docs else "Belum"
                row["Total Missing"] = sum(1 for d in req if d not in stu_docs)
                rows.append(row)
            st.dataframe(pd.DataFrame(rows).sort_values("Total Missing", ascending=False), use_container_width=True, hide_index=True)

    with tabs[2]:
        if documents.empty:
            st.info("Belum ada dokumen.")
        else:
            show = documents.copy()
            show["tanggal_upload"] = show["tanggal_upload"].apply(format_date_id)
            st.dataframe(show[["doc_id", "student_id", "nama_mahasiswa", "jenis_dokumen", "nama_file", "tanggal_upload", "status_verifikasi", "link_file"]], use_container_width=True, hide_index=True)
            st.caption("Klik link_file di Google Sheet atau buka Drive untuk mengunduh file asli.")


def register_fonts():
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass


def generate_invoice_pdf(invoice_row: dict) -> bytes:
    register_fonts()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=14 * mm, leftMargin=14 * mm, topMargin=12 * mm, bottomMargin=12 * mm)

    styles = getSampleStyleSheet()
    brand_name = ParagraphStyle("BrandName", parent=styles["Title"], fontSize=20, leading=23, textColor=TEXT_DARK)
    invoice_title = ParagraphStyle("InvoiceTitle", parent=styles["Title"], fontSize=23, leading=26, textColor=ORANGE_DARK, alignment=1)
    body = ParagraphStyle("Body", parent=styles["BodyText"], fontSize=9.2, leading=12, textColor=TEXT_DARK)
    chinese = ParagraphStyle("Chinese", parent=body, fontName="STSong-Light")
    label = ParagraphStyle("Label", parent=body, fontSize=9, leading=12)
    value = ParagraphStyle("Value", parent=body, fontSize=9.4, leading=12)

    story = []
    accent = Table([["", "", ""]], colWidths=[112 * mm, 40 * mm, 28 * mm], rowHeights=[4 * mm])
    accent.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), BLUE_DARK),
        ("BACKGROUND", (1, 0), (1, 0), BLUE),
        ("BACKGROUND", (2, 0), (2, 0), ORANGE),
    ]))
    story.append(accent)
    story.append(Spacer(1, 6))

    logo_flowable = RLImage(str(LOGO_PATH), width=28 * mm, height=28 * mm) if LOGO_PATH.exists() else Paragraph("", body)
    identity_table = Table([
        [logo_flowable, Paragraph(f"<b>{safe_paragraph_text(PROFILE_FIXED['Nama Brand'])}</b>", brand_name)],
        ["", Paragraph(safe_paragraph_text(PROFILE_FIXED["Alamat ID"]), body)],
        ["", Paragraph(safe_paragraph_text(PROFILE_FIXED["Alamat CN"]), chinese)],
        ["", Paragraph(safe_paragraph_text(PROFILE_FIXED["Alamat EN"]), body)],
        ["", Paragraph(f"Email: {safe_paragraph_text(PROFILE_FIXED['Email'])}<br/>WhatsApp: {safe_paragraph_text(PROFILE_FIXED['Telepon / WA'])}", body)],
    ], colWidths=[32 * mm, 78 * mm])
    identity_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("VALIGN", (0, 0), (0, 0), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("SPAN", (0, 1), (0, 4)),
    ]))

    invoice_card = Table([
        [Paragraph("INVOICE", invoice_title)],
        [Paragraph(f"<b>Kode</b>: {safe_paragraph_text(invoice_row.get('kode_invoice', '-'))}", body)],
        [Paragraph(f"<b>Tanggal</b>: {format_date_id(invoice_row.get('tanggal_invoice'))}", body)],
        [Paragraph(f"<b>Status</b>: {safe_paragraph_text(invoice_row.get('status_pelunasan', '-'))}", body)],
    ], colWidths=[64 * mm])
    invoice_card.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT_ORANGE),
        ("BOX", (0, 0), (-1, -1), 0.9, ORANGE),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))

    top = Table([[identity_table, invoice_card]], colWidths=[112 * mm, 66 * mm])
    top.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(top)
    story.append(Spacer(1, 8))

    def section_title(text, color):
        return Table([[Paragraph(f"<font color='white'><b>{safe_paragraph_text(text)}</b></font>", styles["BodyText"])]], colWidths=[178 * mm], rowHeights=[8 * mm], style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), color), ("LEFTPADDING", (0, 0), (-1, -1), 8)]))

    story.append(section_title("Student & Invoice Details", BLUE))
    detail_rows = [
        [Paragraph("<b>Nama Student</b>", label), Paragraph(safe_paragraph_text(invoice_row.get("nama_mahasiswa", "-")), value), Paragraph("<b>Program</b>", label), Paragraph(safe_paragraph_text(invoice_row.get("program", "-")), value)],
        [Paragraph("<b>Kode Invoice</b>", label), Paragraph(safe_paragraph_text(invoice_row.get("kode_invoice", "-")), value), Paragraph("<b>Status Pelunasan</b>", label), Paragraph(safe_paragraph_text(invoice_row.get("status_pelunasan", "-")), value)],
        [Paragraph("<b>Tanggal Invoice</b>", label), Paragraph(format_date_id(invoice_row.get("tanggal_invoice")), value), Paragraph("<b>Status Pengiriman</b>", label), Paragraph(safe_paragraph_text(invoice_row.get("status_pengiriman", "-")), value)],
        [Paragraph("<b>Deskripsi</b>", label), Paragraph(safe_paragraph_text(invoice_row.get("deskripsi_biaya", "-")), value), Paragraph("<b>Mata Uang</b>", label), Paragraph(safe_paragraph_text(invoice_row.get("mata_uang", "IDR")), value)],
    ]
    detail_table = Table(detail_rows, colWidths=[30 * mm, 58 * mm, 33 * mm, 57 * mm])
    detail_table.setStyle(TableStyle([
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [WHITE, LIGHT_BLUE]),
        ("GRID", (0, 0), (-1, -1), 0.45, LINE),
        ("BOX", (0, 0), (-1, -1), 0.9, BLUE),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(detail_table)
    story.append(Spacer(1, 8))

    story.append(section_title("Program Charge", ORANGE))
    charge_rows = [
        [Paragraph("<font color='white'><b>No</b></font>", label), Paragraph("<font color='white'><b>Deskripsi</b></font>", label), Paragraph("<font color='white'><b>Mata Uang</b></font>", label), Paragraph("<font color='white'><b>Total</b></font>", label)],
        [Paragraph("1", value), Paragraph(safe_paragraph_text(invoice_row.get("deskripsi_biaya", "-")), value), Paragraph(safe_paragraph_text(invoice_row.get("mata_uang", "IDR")), value), Paragraph(format_rupiah(invoice_row.get("harga_program", 0)), value)],
    ]
    charge_table = Table(charge_rows, colWidths=[14 * mm, 110 * mm, 18 * mm, 36 * mm])
    charge_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), ORANGE),
        ("BACKGROUND", (0, 1), (-1, -1), LIGHT_ORANGE),
        ("GRID", (0, 0), (-1, -1), 0.45, LINE),
        ("BOX", (0, 0), (-1, -1), 0.9, ORANGE_DARK),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(charge_table)
    story.append(Spacer(1, 8))

    story.append(section_title("Pembayaran & Ringkasan", BLUE_DARK))
    summary_text = (
        f"<b>Total Program</b>: {format_rupiah(invoice_row.get('harga_program', 0))}<br/>"
        f"<b>Sudah Dibayar</b>: {format_rupiah(invoice_row.get('sudah_dibayar', 0))}<br/>"
        f"<b>Sisa Tagihan</b>: {format_rupiah(invoice_row.get('sisa_tagihan', 0))}"
    )
    payment_table = Table([[Paragraph(safe_paragraph_text(PROFILE_FIXED.get("Info Pembayaran", "-")), value), Paragraph(summary_text, value)]], colWidths=[112 * mm, 66 * mm])
    payment_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), LIGHT_BLUE),
        ("BACKGROUND", (1, 0), (1, 0), LIGHT_ORANGE),
        ("BOX", (0, 0), (-1, -1), 0.9, BLUE_DARK),
        ("INNERGRID", (0, 0), (-1, -1), 0.45, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(payment_table)
    story.append(Spacer(1, 8))

    footer_note = Table([[Paragraph(safe_paragraph_text(PROFILE_FIXED.get("Catatan Footer", "")), body)]], colWidths=[118 * mm])
    footer_note.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF7E6")),
        ("BOX", (0, 0), (-1, -1), 0.9, ORANGE),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    signature_title = ParagraphStyle("SignatureTitle", parent=body, fontSize=8.8, leading=11, textColor=MUTED)
    signature_name = ParagraphStyle("SignatureName", parent=body, fontSize=10, leading=12, textColor=TEXT_DARK, alignment=1)
    signature_role = ParagraphStyle("SignatureRole", parent=body, fontSize=9, leading=11, textColor=MUTED, alignment=1)

    if APPROVAL_PATH.exists():
        approval_img = RLImage(str(APPROVAL_PATH), width=56 * mm, height=20 * mm)
    else:
        approval_img = Paragraph("", body)

    approval_block = Table(
        [
            [Paragraph("<b>Authorized Signature</b>", signature_title)],
            [approval_img],
            [Paragraph("<b>Yenny Pricila</b>", signature_name)],
            [Paragraph("Management", signature_role)],
        ],
        colWidths=[60 * mm],
    )
    approval_block.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("LINEABOVE", (0, 0), (-1, 0), 0.4, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("ALIGN", (0, 1), (0, 1), "CENTER"),
    ]))

    footer_combo = Table([[footer_note, approval_block]], colWidths=[118 * mm, 60 * mm])
    footer_combo.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(footer_combo)

    doc.build(story)
    buffer.seek(0)
    return buffer.read()


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


def invoices_page(data: dict[str, Any], endpoint: str, token: str) -> None:
    render_header("Modul Invoice & Pembayaran", "Buat invoice dan catat pembayaran langsung ke Google Sheet.")
    tabs = st.tabs(["Dashboard Invoice", "Tambah Invoice", "Catat Pembayaran", "PDF Invoice"])

    students = data["students"]
    invoices = data["invoices"]
    payments = data["payments"]

    with tabs[0]:
        c1, c2, c3, c4 = st.columns(4)
        total_invoice = len(invoices)
        total_nilai = float(invoices["harga_program"].sum()) if not invoices.empty else 0
        total_paid = float(invoices["sudah_dibayar"].sum()) if not invoices.empty else 0
        total_outstanding = float(invoices["sisa_tagihan"].sum()) if not invoices.empty else 0
        c1.metric("Total Invoice", total_invoice)
        c2.metric("Total Nilai", format_rupiah(total_nilai))
        c3.metric("Sudah Dibayar", format_rupiah(total_paid))
        c4.metric("Outstanding", format_rupiah(total_outstanding))
        if not invoices.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("### Status Pelunasan")
                st.bar_chart(invoices.groupby("status_pelunasan")["invoice_id"].count().sort_values(ascending=False))
            with col2:
                st.markdown("### Outstanding per Program")
                st.bar_chart(invoices.groupby("program")["sisa_tagihan"].sum().sort_values(ascending=False))

    with tabs[1]:
        if students.empty:
            st.info("Tambahkan data mahasiswa terlebih dahulu.")
        else:
            options = students["student_id"].astype(str) + " - " + students["nama_lengkap"].astype(str)
            with st.form("create_invoice_live_form"):
                selected = st.selectbox("Pilih Mahasiswa", options)
                selected_id = selected.split(" - ", 1)[0]
                srow = students[students["student_id"] == selected_id].iloc[0]
                c1, c2 = st.columns(2)
                program = c1.text_input("Program", value=str(srow["program_diminati"]))
                deskripsi_biaya = c2.text_input("Deskripsi Biaya", value=f"Biaya program {srow['program_diminati']}")
                c3, c4, c5 = st.columns(3)
                mata_uang = c3.text_input("Mata Uang", value="IDR")
                harga_program = c4.number_input("Harga Program", min_value=0.0, value=float(srow["estimasi_biaya"] or 0), step=100000.0)
                status_pengiriman = c5.selectbox("Status Pengiriman", data["references"].get("status_pengiriman", []))
                c6, c7 = st.columns(2)
                tanggal_invoice = c6.date_input("Tanggal Invoice", value=datetime.now().date())
                tanggal_kirim = c7.date_input("Tanggal Kirim", value=None)
                catatan_invoice = st.text_area("Catatan Invoice")
                submit = st.form_submit_button("Buat Invoice Live")
            if submit:
                try:
                    post_action(endpoint, token, "create_invoice", {
                        "student_id": selected_id,
                        "nama_mahasiswa": srow["nama_lengkap"],
                        "program": program,
                        "deskripsi_biaya": deskripsi_biaya,
                        "mata_uang": mata_uang,
                        "harga_program": float(harga_program),
                        "tanggal_invoice": tanggal_invoice.isoformat(),
                        "status_pengiriman": status_pengiriman,
                        "tanggal_kirim": tanggal_kirim.isoformat() if tanggal_kirim else "",
                        "catatan_invoice": catatan_invoice,
                    })
                    st.success("Invoice berhasil dibuat di Google Sheet.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal membuat invoice: {e}")

    with tabs[2]:
        if invoices.empty:
            st.info("Belum ada invoice.")
        else:
            options = invoices["kode_invoice"].astype(str) + " - " + invoices["nama_mahasiswa"].astype(str)
            with st.form("record_payment_live_form"):
                selected = st.selectbox("Pilih Invoice", options)
                invoice_code = selected.split(" - ", 1)[0]
                row = invoices[invoices["kode_invoice"] == invoice_code].iloc[0]
                st.caption(f"Outstanding saat ini: {format_rupiah(row['sisa_tagihan'])}")
                c1, c2, c3 = st.columns(3)
                tanggal_pembayaran = c1.date_input("Tanggal Pembayaran", value=datetime.now().date())
                jumlah_pembayaran = c2.number_input("Jumlah Pembayaran", min_value=0.0, step=100000.0)
                metode_pembayaran = c3.selectbox("Metode Pembayaran", data["references"].get("metode_pembayaran", ["Transfer"]))
                c4, c5 = st.columns(2)
                bukti_pembayaran_link = c4.text_input("Link Bukti Pembayaran")
                dicatat_oleh = c5.text_input("Dicatat Oleh", value="Finance")
                catatan = st.text_area("Catatan Pembayaran")
                submit = st.form_submit_button("Simpan Pembayaran Live")
            if submit:
                try:
                    post_action(endpoint, token, "record_payment", {
                        "invoice_id": row["invoice_id"],
                        "kode_invoice": row["kode_invoice"],
                        "student_id": row["student_id"],
                        "tanggal_pembayaran": tanggal_pembayaran.isoformat(),
                        "jumlah_pembayaran": float(jumlah_pembayaran),
                        "metode_pembayaran": metode_pembayaran,
                        "bukti_pembayaran_link": bukti_pembayaran_link,
                        "dicatat_oleh": dicatat_oleh,
                        "catatan": catatan,
                    })
                    st.success("Pembayaran berhasil dicatat di Google Sheet.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan pembayaran: {e}")

            if not payments.empty:
                show = payments.copy()
                show["tanggal_pembayaran"] = show["tanggal_pembayaran"].apply(format_date_id)
                show["jumlah_pembayaran"] = show["jumlah_pembayaran"].map(format_rupiah)
                st.dataframe(show[["payment_id", "invoice_id", "student_id", "tanggal_pembayaran", "jumlah_pembayaran", "metode_pembayaran", "dicatat_oleh"]], use_container_width=True, hide_index=True)

    with tabs[3]:
        if invoices.empty:
            st.info("Belum ada invoice.")
        else:
            options = invoices["kode_invoice"].astype(str) + " - " + invoices["nama_mahasiswa"].astype(str)
            selected = st.selectbox("Pilih Invoice untuk PDF", options)
            invoice_code = selected.split(" - ", 1)[0]
            row = invoices[invoices["kode_invoice"] == invoice_code].iloc[0].to_dict()
            c1, c2, c3 = st.columns(3)
            c1.metric("Nama Mahasiswa", row.get("nama_mahasiswa", "-"))
            c2.metric("Program", row.get("program", "-"))
            c3.metric("Sisa Tagihan", format_rupiah(row.get("sisa_tagihan", 0)))
            pdf_bytes = generate_invoice_pdf(row)
            st.download_button("Download PDF Invoice", data=pdf_bytes, file_name=f"{invoice_code}.pdf", mime="application/pdf")


def help_page() -> None:
    render_header("Bantuan & SOP", "Panduan setup live Google Sheet.")
    st.markdown(
        "### Setup singkat\n"
        "1. Buat Google Sheet baru.\n"
        "2. Tempel file **Code.gs** ke Google Apps Script.\n"
        "3. Isi `SPREADSHEET_ID`, `ROOT_FOLDER_ID`, dan `WRITE_TOKEN` di script.\n"
        "4. Jalankan `setupNihaomaSheets()` sekali.\n"
        "5. Deploy sebagai **Web app** dan ambil URL-nya.\n"
        "6. Isi `APPS_SCRIPT_URL` dan `WRITE_TOKEN` di Streamlit Secrets.\n"
        "7. Deploy app Streamlit."
    )
    st.info("Versi ini tidak lagi memakai upload workbook Excel di app. Semua write masuk langsung ke Google Sheet live.")


def main() -> None:
    endpoint, token = require_secrets()
    try:
        data = fetch_bootstrap(endpoint, token)
    except Exception as e:
        st.error(f"Gagal menghubungkan app ke Google Sheets live: {e}")
        st.stop()

    menu = st.sidebar.radio("Menu", ["Dashboard", "Calon Mahasiswa", "Dokumen", "Invoice & Pembayaran", "Bantuan & SOP"])
    if menu == "Dashboard":
        dashboard_page(data)
    elif menu == "Calon Mahasiswa":
        students_page(data, endpoint, token)
    elif menu == "Dokumen":
        documents_page(data, endpoint, token)
    elif menu == "Invoice & Pembayaran":
        invoices_page(data, endpoint, token)
    else:
        help_page()


if __name__ == "__main__":
    main()


import os
from typing import Any, Dict, List

import pandas as pd
import plotly.express as px
import requests
import streamlit as st

st.set_page_config(page_title="Nihaoma Student Operations", layout="wide")

SCRIPT_URL = st.secrets.get("SCRIPT_URL", os.getenv("SCRIPT_URL", ""))
WRITE_TOKEN = st.secrets.get("WRITE_TOKEN", os.getenv("WRITE_TOKEN", ""))
TIMEOUT = 60


# ---------- API helpers ----------
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
    data = api_get("bootstrap")
    if not data.get("ok"):
        raise RuntimeError(data.get("error", "Gagal memuat bootstrap"))
    return data



def clear_cache_and_rerun() -> None:
    st.cache_data.clear()
    st.rerun()


# ---------- Formatting ----------
def as_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame(rows) if rows else pd.DataFrame()



def safe_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value)



def format_currency(value: Any) -> str:
    try:
        num = float(value or 0)
    except Exception:
        num = 0
    return f"Rp {num:,.0f}".replace(",", ".")



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


# ---------- Dashboard ----------
def render_dashboard(students_df: pd.DataFrame) -> None:
    st.subheader("Dashboard")

    if students_df.empty:
        st.info("Belum ada data mahasiswa.")
        return

    active_df = students_df.copy()
    if "is_active" in active_df.columns:
        active_df = active_df[
            active_df["is_active"].astype(str).str.upper().isin(["TRUE", "1", "YA", "YES", ""])
        ].copy()

    total_students = len(active_df)
    new_lead = int((active_df.get("status_proses", pd.Series(dtype=str)) == "New Lead").sum())
    active_pipeline = int(active_df.get("status_proses", pd.Series(dtype=str)).isin([
        "New Lead", "Follow Up", "Interested", "Dokumen Awal Masuk",
        "Siap Daftar", "Sudah Daftar", "Menunggu Pembayaran", "Proses Visa",
        "Siap Berangkat", "Aktif"
    ]).sum())
    total_estimated = pd.to_numeric(active_df.get("estimasi_biaya", 0), errors="coerce").fillna(0).sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Mahasiswa", total_students)
    c2.metric("New Lead", new_lead)
    c3.metric("Pipeline Aktif", active_pipeline)
    c4.metric("Estimasi Nilai", format_currency(total_estimated))

    chart_left, chart_right = st.columns(2)

    with chart_left:
        st.markdown("**Distribusi Status Proses**")
        status_df = (
            active_df.assign(status_proses=active_df.get("status_proses", "").replace("", "Belum Diisi"))
            .groupby("status_proses", dropna=False)
            .size()
            .reset_index(name="jumlah")
            .sort_values("jumlah", ascending=False)
        )
        fig_status = px.pie(status_df, names="status_proses", values="jumlah", hole=0.35)
        st.plotly_chart(fig_status, use_container_width=True)

    with chart_right:
        st.markdown("**Distribusi Program Diminati**")
        prog_df = (
            active_df.assign(program_diminati=active_df.get("program_diminati", "").replace("", "Belum Diisi"))
            .groupby("program_diminati", dropna=False)
            .size()
            .reset_index(name="jumlah")
            .sort_values("jumlah", ascending=False)
        )
        fig_program = px.bar(prog_df, x="program_diminati", y="jumlah")
        st.plotly_chart(fig_program, use_container_width=True)

    bottom_left, bottom_right = st.columns(2)
    with bottom_left:
        st.markdown("**Distribusi PIC**")
        pic_df = (
            active_df.assign(pic_admin=active_df.get("pic_admin", "").replace("", "Belum Assign"))
            .groupby("pic_admin", dropna=False)
            .size()
            .reset_index(name="jumlah")
            .sort_values("jumlah", ascending=False)
        )
        fig_pic = px.pie(pic_df, names="pic_admin", values="jumlah")
        st.plotly_chart(fig_pic, use_container_width=True)

    with bottom_right:
        st.markdown("**Mahasiswa per Intake**")
        intake_df = (
            active_df.assign(intake=active_df.get("intake", "").replace("", "Belum Diisi"))
            .groupby("intake", dropna=False)
            .size()
            .reset_index(name="jumlah")
            .sort_values("jumlah", ascending=False)
        )
        fig_intake = px.bar(intake_df, x="intake", y="jumlah")
        st.plotly_chart(fig_intake, use_container_width=True)


# ---------- Student list ----------
def render_student_list(students_df: pd.DataFrame, refs: Dict[str, Any]) -> None:
    st.subheader("Modul Calon Mahasiswa")

    if students_df.empty:
        st.info("Belum ada data mahasiswa.")
        return

    tabs = st.tabs(["Daftar Mahasiswa", "Tambah Data", "Detail & Progress"])

    with tabs[0]:
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
            searchable_cols = [c for c in ["student_id", "nama_lengkap", "email", "no_whatsapp", "program_diminati"] if c in filtered.columns]
            mask = pd.Series(False, index=filtered.index)
            for col in searchable_cols:
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
        if not student_options:
            st.warning("Tidak ada data yang cocok dengan filter.")
            return

        selected_id = st.selectbox("Pilih student_id untuk aksi", student_options, key="selected_student_action")
        action_col1, action_col2, action_col3 = st.columns([1, 1, 3])
        if action_col1.button("Edit data", use_container_width=True):
            st.session_state["edit_student_id"] = selected_id
        if action_col2.button("Hapus data", use_container_width=True):
            st.session_state["delete_student_id"] = selected_id

        if st.session_state.get("edit_student_id"):
            edit_id = st.session_state["edit_student_id"]
            row_df = students_df[students_df["student_id"].astype(str) == str(edit_id)]
            if not row_df.empty:
                student = row_df.iloc[0].to_dict()
                st.markdown("### Form Edit Mahasiswa")
                render_edit_form(student, refs)

        if st.session_state.get("delete_student_id"):
            delete_id = st.session_state["delete_student_id"]
            st.markdown("### Konfirmasi Hapus Mahasiswa")
            st.warning(
                "Aksi ini akan menghapus data mahasiswa dari students_master. "
                "Kalau backend delete lengkap dipasang, log progress dan data terkait juga bisa ikut terhapus."
            )
            confirm_text = st.text_input(
                f"Ketik {delete_id} untuk konfirmasi hapus",
                key="confirm_delete_text"
            )
            del_col1, del_col2 = st.columns(2)
            if del_col1.button("Ya, hapus sekarang", type="primary", use_container_width=True):
                if confirm_text != delete_id:
                    st.error("Konfirmasi tidak cocok.")
                else:
                    try:
                        result = api_post("delete_student", {"student_id": delete_id})
                        if result.get("ok"):
                            st.success(f"Data {delete_id} berhasil dihapus.")
                            st.session_state.pop("delete_student_id", None)
                            st.session_state.pop("confirm_delete_text", None)
                            clear_cache_and_rerun()
                        else:
                            st.error(result.get("error", "Gagal menghapus data"))
                    except Exception as exc:
                        st.error(f"Gagal menghapus data: {exc}")
            if del_col2.button("Batal", use_container_width=True):
                st.session_state.pop("delete_student_id", None)
                st.session_state.pop("confirm_delete_text", None)
                st.rerun()

    with tabs[1]:
        render_add_form(refs)

    with tabs[2]:
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
                    if field in student:
                        st.write(f"**{field}**: {safe_text(student.get(field))}")
            with right:
                st.markdown("### Update Progress")
                with st.form("form_update_progress"):
                    status_options = refs.get("status_proses", [safe_text(student.get("status_proses"))]) or [safe_text(student.get("status_proses"))]
                    next_action_options = refs.get("next_action", []) or [safe_text(student.get("next_action"))]
                    status_baru = st.selectbox(
                        "Status Baru",
                        status_options,
                        index=option_index(status_options, student.get("status_proses"))
                    )
                    next_action = st.selectbox(
                        "Next Action",
                        [""] + next_action_options,
                        index=option_index([""] + next_action_options, student.get("next_action"))
                    )
                    tanggal_next_action = st.date_input("Tanggal Next Action", value=None, format="YYYY-MM-DD")
                    catatan = st.text_area("Catatan Progress")
                    updated_by = st.text_input("Updated by", value=safe_text(student.get("pic_admin")) or "Admin")
                    submit_progress = st.form_submit_button("Simpan Progress")
                    if submit_progress:
                        try:
                            result = api_post(
                                "update_progress",
                                {
                                    "student_id": selected_detail_id,
                                    "status_baru": status_baru,
                                    "next_action": next_action,
                                    "tanggal_next_action": str(tanggal_next_action) if tanggal_next_action else "",
                                    "catatan": catatan,
                                    "updated_by": updated_by,
                                },
                            )
                            if result.get("ok"):
                                st.success("Progress berhasil diperbarui.")
                                clear_cache_and_rerun()
                            else:
                                st.error(result.get("error", "Gagal update progress"))
                        except Exception as exc:
                            st.error(f"Gagal update progress: {exc}")


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

        submit = st.form_submit_button("Simpan Perubahan")
        if submit:
            try:
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
            except Exception as exc:
                st.error(f"Gagal update mahasiswa: {exc}")



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

        submit = st.form_submit_button("Tambah Mahasiswa")
        if submit:
            if not nama_lengkap.strip():
                st.error("Nama lengkap wajib diisi.")
                return
            try:
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
            except Exception as exc:
                st.error(f"Gagal menambah mahasiswa: {exc}")


# ---------- Main ----------
def main() -> None:
    st.title("Nihaoma Student Operations")
    st.caption("Dashboard operasional calon mahasiswa yang terhubung ke Google Sheet live")

    try:
        bootstrap = load_bootstrap()
    except Exception as exc:
        st.error(f"Gagal memuat data dari Apps Script: {exc}")
        st.stop()

    students_df = normalize_df(as_df(bootstrap.get("students", [])))
    refs = bootstrap.get("references", {})

    menu = st.sidebar.radio(
        "Menu",
        ["Dashboard", "Calon Mahasiswa"],
        index=0,
    )

    if st.sidebar.button("Refresh data", use_container_width=True):
        clear_cache_and_rerun()

    if menu == "Dashboard":
        render_dashboard(students_df)
    elif menu == "Calon Mahasiswa":
        render_student_list(students_df, refs)


if __name__ == "__main__":
    main()
