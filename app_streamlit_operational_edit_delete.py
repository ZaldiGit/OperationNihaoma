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
