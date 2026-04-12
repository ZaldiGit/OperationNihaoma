import streamlit as st
import pandas as pd
import requests
import plotly.express as px
from datetime import date
from typing import Any, Dict, List

st.set_page_config(page_title="Nihaoma Student Operations", layout="wide")

TIMEOUT = 60
SCRIPT_URL = st.secrets.get("SCRIPT_URL", "")
WRITE_TOKEN = st.secrets.get("WRITE_TOKEN", "")


def ensure_config():
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
    body = {"action": action, "token": WRITE_TOKEN, **payload}
    resp = requests.post(SCRIPT_URL, json=body, timeout=TIMEOUT)
    resp.raise_for_status()
    return resp.json()


@st.cache_data(ttl=60)
def load_bootstrap() -> Dict[str, Any]:
    return api_get("bootstrap")


def rerun_and_clear_cache():
    load_bootstrap.clear()
    st.rerun()


def as_df(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    return pd.DataFrame(rows if rows else [])


def format_rp(value: Any) -> str:
    try:
        n = float(value or 0)
    except Exception:
        n = 0.0
    return f"Rp {n:,.0f}".replace(",", ".")


def get_program_options(refs: Dict[str, Any], students_df: pd.DataFrame) -> List[str]:
    opts = refs.get("program", []) or refs.get("program_diminati", []) or []
    opts = [str(x).strip() for x in opts if str(x).strip()]
    if students_df is not None and not students_df.empty and "program_diminati" in students_df.columns:
        for p in students_df["program_diminati"].dropna().astype(str).tolist():
            p = p.strip()
            if p and p not in opts:
                opts.append(p)
    return opts


def default_program_for_student(student_row: Dict[str, Any], program_options: List[str]) -> tuple[list[str], int]:
    default_program = str(student_row.get("program_diminati", "") or "").strip()
    options = list(program_options)
    if default_program and default_program not in options:
        options = [default_program] + options
    if not options:
        options = [default_program or "-"]
    idx = options.index(default_program) if default_program in options else 0
    return options, idx


def render_dashboard(bootstrap: Dict[str, Any]):
    students = as_df(bootstrap.get("students", []))
    invoices = as_df(bootstrap.get("invoices", []))
    payments = as_df(bootstrap.get("payments", []))

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Mahasiswa", len(students))
    c2.metric("Total Invoice", len(invoices))
    c3.metric("Total Pembayaran", len(payments))
    outstanding = 0
    if not invoices.empty and "sisa_tagihan" in invoices.columns:
        outstanding = pd.to_numeric(invoices["sisa_tagihan"], errors="coerce").fillna(0).sum()
    c4.metric("Outstanding", format_rp(outstanding))

    if not invoices.empty and "status_pelunasan" in invoices.columns:
        status_counts = invoices["status_pelunasan"].fillna("Belum Lunas").value_counts().reset_index()
        status_counts.columns = ["status_pelunasan", "jumlah"]
        st.plotly_chart(
            px.pie(status_counts, names="status_pelunasan", values="jumlah", title="Distribusi Status Pelunasan"),
            use_container_width=True,
        )


def render_invoice_manual(student_row: Dict[str, Any], refs: Dict[str, Any]):
    st.subheader("Buat Invoice Manual")

    program_options = get_program_options(refs, as_df([]))
    program_options, default_idx = default_program_for_student(student_row, program_options)

    invoice_type_options = ["Pendaftaran", "Admin", "Transport", "Lainnya"]
    currency_options = ["IDR", "USD", "CNY"]
    shipping_options = refs.get("status_pengiriman", []) or ["Belum Dikirim", "Sudah Dikirim"]

    with st.form("manual_invoice_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            tanggal_invoice = st.date_input("Tanggal Invoice", value=date.today())
        with col2:
            invoice_type = st.selectbox("Jenis Invoice", invoice_type_options)
        with col3:
            mata_uang = st.selectbox("Mata Uang", currency_options, index=0)

        program = st.selectbox("Program", options=program_options, index=default_idx)
        nominal = st.number_input("Nominal Invoice", min_value=0.0, value=0.0, step=100000.0)
        deskripsi = st.text_area("Deskripsi Biaya", value="")
        status_pengiriman = st.selectbox("Status Pengiriman", shipping_options, index=0)
        kirim_hari_ini = st.checkbox("Sudah dikirim hari ini")
        tanggal_kirim = st.date_input("Tanggal Kirim", value=date.today(), disabled=not kirim_hari_ini)
        catatan_invoice = st.text_area("Catatan Invoice", value="")

        submitted = st.form_submit_button("Buat Invoice Manual")

    if submitted:
        payload = {
            "student_id": student_row.get("student_id", ""),
            "nama_mahasiswa": student_row.get("nama_lengkap", ""),
            "tanggal_invoice": str(tanggal_invoice),
            "program": program,
            "deskripsi_biaya": deskripsi or f"Invoice {invoice_type}",
            "mata_uang": mata_uang,
            "harga_program": nominal,
            "status_pengiriman": status_pengiriman,
            "tanggal_kirim": str(tanggal_kirim) if kirim_hari_ini else "",
            "catatan_invoice": catatan_invoice,
        }
        res = api_post("create_invoice", payload)
        if res.get("ok"):
            st.success(f"Invoice berhasil dibuat: {res.get('kode_invoice', '-')}")
            rerun_and_clear_cache()
        else:
            st.error(res.get("error", "Gagal membuat invoice"))


def render_invoice_package(student_row: Dict[str, Any], refs: Dict[str, Any]):
    st.subheader("Buat Paket Invoice")

    estimasi_biaya = float(student_row.get("estimasi_biaya") or 0)
    program_name = str(student_row.get("program_diminati") or "")
    is_bahasa = "bahasa" in program_name.lower()
    biaya_pendaftaran = 2_000_000 if is_bahasa else 3_000_000
    biaya_transport = 4_000_000
    biaya_admin = max(estimasi_biaya - biaya_pendaftaran, 0) + biaya_transport
    total = biaya_pendaftaran + biaya_admin

    k1, k2, k3 = st.columns(3)
    k1.metric("Biaya Pendaftaran", format_rp(biaya_pendaftaran))
    k2.metric("Biaya Admin", format_rp(biaya_admin))
    k3.metric("Biaya Transport", format_rp(biaya_transport))
    st.info(f"Total invoice admin yang akan dibuat: {format_rp(biaya_admin)}. Total keseluruhan kewajiban mahasiswa: {format_rp(total)}.")

    shipping_options = refs.get("status_pengiriman", []) or ["Belum Dikirim", "Sudah Dikirim"]

    with st.form("package_invoice_form"):
        col1, col2 = st.columns(2)
        with col1:
            tanggal_invoice = st.date_input("Tanggal Invoice", value=date.today(), key="pkg_tgl")
        with col2:
            mata_uang = st.selectbox("Mata Uang", ["IDR", "USD", "CNY"], index=0, key="pkg_curr")

        status_pengiriman = st.selectbox("Status Pengiriman", shipping_options, index=0, key="pkg_ship")
        kirim_hari_ini = st.checkbox("Sudah dikirim hari ini", key="pkg_sent_today")
        tanggal_kirim = st.date_input("Tanggal Kirim", value=date.today(), disabled=not kirim_hari_ini, key="pkg_sent_date")
        catatan = st.text_area("Catatan Invoice Paket", value="Invoice paket otomatis: Pendaftaran + Admin/Transport")
        submitted = st.form_submit_button("Buat 2 invoice otomatis")

    if submitted:
        payload = {
            "student_id": student_row.get("student_id", ""),
            "nama_mahasiswa": student_row.get("nama_lengkap", ""),
            "program": student_row.get("program_diminati", ""),
            "estimasi_biaya": estimasi_biaya,
            "mata_uang": mata_uang,
            "tanggal_invoice": str(tanggal_invoice),
            "status_pengiriman": status_pengiriman,
            "tanggal_kirim": str(tanggal_kirim) if kirim_hari_ini else "",
            "catatan_invoice": catatan,
        }
        res = api_post("create_invoice_package", payload)
        if res.get("ok"):
            st.success("Paket invoice berhasil dibuat.")
            rerun_and_clear_cache()
        else:
            st.error(res.get("error", "Unknown action"))


def render_record_payment(bootstrap: Dict[str, Any]):
    st.subheader("Record Pembayaran")
    invoices = as_df(bootstrap.get("invoices", []))
    if invoices.empty:
        st.info("Belum ada invoice.")
        return

    invoices = invoices.copy()
    if "kode_invoice" not in invoices.columns:
        invoices["kode_invoice"] = invoices["invoice_id"]

    label_map = {}
    labels = []
    for _, row in invoices.iterrows():
        label = f"{row.get('kode_invoice', row.get('invoice_id', '-'))} | {row.get('nama_mahasiswa', '-')}"
        labels.append(label)
        label_map[label] = row.to_dict()

    selected_label = st.selectbox("Pilih Invoice", labels)
    row = label_map[selected_label]

    with st.form("payment_form"):
        col1, col2 = st.columns(2)
        with col1:
            tanggal_bayar = st.date_input("Tanggal Pembayaran", value=date.today())
        with col2:
            metode = st.selectbox("Metode Pembayaran", ["Transfer", "Cash", "QRIS", "EDC", "Lainnya"])
        jumlah = st.number_input("Jumlah Pembayaran", min_value=0.0, value=0.0, step=100000.0)
        bukti = st.text_input("Link Bukti Pembayaran", value="")
        catatan = st.text_area("Catatan", value="")
        submitted = st.form_submit_button("Simpan Pembayaran")

    if submitted:
        payload = {
            "invoice_id": row.get("invoice_id", ""),
            "student_id": row.get("student_id", ""),
            "tanggal_pembayaran": str(tanggal_bayar),
            "jumlah_pembayaran": jumlah,
            "metode_pembayaran": metode,
            "bukti_pembayaran_link": bukti,
            "catatan": catatan,
        }
        res = api_post("record_payment", payload)
        if res.get("ok"):
            st.success("Pembayaran berhasil dicatat.")
            rerun_and_clear_cache()
        else:
            st.error(res.get("error", "Gagal mencatat pembayaran"))


def main():
    try:
        bootstrap = load_bootstrap()
    except Exception as e:
        st.title("Nihaoma Student Operations")
        st.caption("Dashboard operasional calon mahasiswa yang terhubung ke Google Sheet live")
        st.error(f"Gagal memuat data awal: {e}")
        st.stop()

    refs = bootstrap.get("references", {})
    students = as_df(bootstrap.get("students", []))
    invoices = as_df(bootstrap.get("invoices", []))

    with st.sidebar:
        st.markdown("### Menu")
        page = st.radio("Menu", ["Dashboard", "Calon Mahasiswa", "Dokumen", "Invoice & Pembayaran", "Bantuan & SOP"], label_visibility="collapsed")
        if st.button("Refresh data", use_container_width=True):
            rerun_and_clear_cache()
        meta = bootstrap.get("meta", {})
        if meta.get("generated_at"):
            st.caption(f"Data terakhir dimuat: {meta['generated_at']}")

    st.title("Nihaoma Student Operations")
    st.caption("Dashboard operasional calon mahasiswa yang terhubung ke Google Sheet live")

    if page == "Dashboard":
        render_dashboard(bootstrap)
        return

    if page != "Invoice & Pembayaran":
        st.info("Halaman ini dipertahankan seperti versi Anda sebelumnya. File ini fokus memperbaiki modul invoice.")
        return

    st.subheader("Invoice & Pembayaran")

    if students.empty:
        st.warning("Belum ada data mahasiswa.")
        return

    selected_sid = st.selectbox(
        "Pilih mahasiswa",
        options=students["student_id"].astype(str).tolist(),
        index=0,
    )
    student_row = students[students["student_id"].astype(str) == str(selected_sid)].iloc[0].to_dict()

    tabs = st.tabs([
        "Dashboard Invoice",
        "Buat Paket Invoice",
        "Buat Invoice Manual",
        "Record Pembayaran",
        "Download PDF Invoice",
    ])

    with tabs[0]:
        if invoices.empty:
            st.info("Belum ada invoice.")
        else:
            df = invoices.copy()
            if "sisa_tagihan" in df.columns:
                df["sisa_tagihan_num"] = pd.to_numeric(df["sisa_tagihan"], errors="coerce").fillna(0)
            else:
                df["sisa_tagihan_num"] = 0
            st.dataframe(df, use_container_width=True)

    with tabs[1]:
        render_invoice_package(student_row, refs)

    with tabs[2]:
        render_invoice_manual(student_row, refs)

    with tabs[3]:
        render_record_payment(bootstrap)

    with tabs[4]:
        st.info("Tab download PDF bisa tetap memakai implementasi Anda yang lama. Fokus file ini adalah perbaikan form invoice.")


if __name__ == "__main__":
    main()
