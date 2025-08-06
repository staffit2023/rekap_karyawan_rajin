import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📋 Rekapan Karyawan Rajin")

uploaded_file = st.file_uploader("Upload file absensi Excel (.xlsx/.xls)", type=["xlsx", "xls"])

if uploaded_file:
    # Baca Excel, header di baris kedua (index 1)
    try:
        df = pd.read_excel(uploaded_file, header=1)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        st.stop()

    # Hapus baris kosong
    df.dropna(how='all', inplace=True)

    # Bersihkan nama kolom: spasi jadi underscore dan huruf kecil
    df.columns = df.columns.str.strip().str.replace(r'\s+', '_', regex=True).str.lower()

    # Tampilkan data awal
    st.subheader("📄 Data Absensi (Preview)")
    st.dataframe(df)

    # Daftar kata-kata izin
    kata_izin = ['izin', 'sakit', 'cuti', 'dispensasi', 'alpha']

    # Hitung total hari kerja per karyawan
    hari_kerja = df.groupby('nama').size().reset_index(name='jumlah_hari_kerja')

    # Fungsi cek kehadiran bersih
    def bersih(row):
        no_telat = pd.isna(row['terlambat']) or row['terlambat'] == 0

        # Cek kolom 'izin' harus kosong atau 0
        no_izin_kolom = pd.isna(row.get('izin')) or row['izin'] == 0

        # Cek tidak ada kata izin di keterangan
        no_izin_keterangan = True
        if pd.notna(row.get('keterangan')):
            ket = str(row['keterangan']).lower()
            no_izin_keterangan = not any(k in ket for k in kata_izin)

        # Cek jam_kerja tidak "Tidak hadir"
        jam_kerja_oke = True
        if pd.notna(row.get('jam_kerja')):
            jam_kerja_oke = str(row['jam_kerja']).strip().lower() != 'tidak hadir'

        # Cek scan_pulang harus ada dan tidak lebih awal dari jam_pulang
        pulang_oke = True
        if pd.notna(row.get('jam_pulang')) and pd.notna(row.get('scan_pulang')):
            try:
                jam_pulang = pd.to_datetime(str(row['jam_pulang'])).time()
                scan_pulang = pd.to_datetime(str(row['scan_pulang'])).time()
                pulang_oke = scan_pulang >= jam_pulang
            except:
                pulang_oke = False
        else:
            pulang_oke = False

        return no_telat and no_izin_kolom and no_izin_keterangan and jam_kerja_oke and pulang_oke

    df['bersih'] = df.apply(bersih, axis=1)

    # Hitung hadir bersih
    hadir_bersih = df[df['bersih']].groupby('nama').size().reset_index(name='hadir_tanpa_telat_izin')

    # Gabungkan semua
    rekap = pd.merge(hari_kerja, hadir_bersih, on='nama', how='left')
    rekap['hadir_tanpa_telat_izin'] = rekap['hadir_tanpa_telat_izin'].fillna(0).astype(int)

    # Tandai status rajin
    rekap['status'] = rekap.apply(
        lambda row: 'Rajin' if row['jumlah_hari_kerja'] == row['hadir_tanpa_telat_izin'] else 'Tidak', axis=1
    )

    # Pilihan untuk tampilkan semua atau hanya yang rajin
    filter_rajin = st.checkbox("Tampilkan hanya karyawan rajin", value=True)

    if filter_rajin:
        rekap_tampil = rekap[rekap['status'] == 'Rajin']
    else:
        rekap_tampil = rekap

    st.subheader("📊 Rekapan Karyawan Rajin")
    st.dataframe(rekap_tampil)

    # Filter data lengkap untuk karyawan yang rajin
    nama_rajin = rekap[rekap['status'] == 'Rajin']['nama'].tolist()
    df_rajin_detail = df[df['nama'].isin(nama_rajin)]

    output_rekap = BytesIO()
    with pd.ExcelWriter(output_rekap, engine='xlsxwriter') as writer:
        rekap_tampil.to_excel(writer, index=False, sheet_name='Rekap_Rajin')

    # Tombol download Excel untuk Rekap
    st.download_button(
        label="⬇️ Download Rekap (Excel)",
        data=output_rekap.getvalue(),
        file_name="rekap_karyawan_rajin.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Tampilkan data lengkap
    st.subheader("📑 Detail Absensi Karyawan Rajin")
    st.dataframe(df_rajin_detail)

    # Kolom yang ingin ditampilkan
    kolom_dipilih = ['tanggal', 'nama', 'jam_masuk', 'scan_masuk', 'jam_pulang', 'scan_pulang', 'departemen']
    kolom_tersedia = [kol for kol in kolom_dipilih if kol in df_rajin_detail.columns]

    # Filter hanya kolom yang tersedia
    df_excel = df_rajin_detail[kolom_tersedia]

    # Simpan ke Excel menggunakan BytesIO
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        df_excel.to_excel(writer, index=False, sheet_name='Detail_Rajin')

    # Tombol download Excel
    st.download_button(
        label="⬇️ Download Detail (Excel)",
        data=output_excel.getvalue(),
        file_name="detail_karyawan_rajin.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success("✅ Proses selesai! Silakan cek tabel di atas.")
