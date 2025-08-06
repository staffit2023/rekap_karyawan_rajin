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

    # Bersihkan nama kolom: spasi jadi underscore
    df.columns = df.columns.str.strip().str.replace(r'\s+', '_', regex=True)

    # Tampilkan data awal
    st.subheader("📄 Data Absensi (Preview)")
    st.dataframe(df)

    # Daftar kata-kata izin
    kata_izin = ['izin', 'sakit', 'cuti', 'dispensasi', 'alpha']
    pengecualian_scan_pulang = ['Libur Rutin', 'tugas kantor','Izin dinas (Izin keperluan kantor)']

    # Hitung total hari kerja per karyawan
    hari_kerja = df.groupby('Nama').size().reset_index(name='Jumlah_Hari_Kerja')

    # Fungsi cek kehadiran bersih
    def bersih(row):
        no_telat = pd.isna(row['Terlambat']) or row['Terlambat'] == 0

        # Cek kolom 'Izin' harus kosong atau 0
        no_izin_kolom = pd.isna(row.get('Izin')) or row['Izin'] == 0

        # Cek tidak ada kata izin di keterangan
        no_izin_keterangan = True
        if pd.notna(row.get('Keterangan')):
            ket = str(row['Keterangan']).lower()
            no_izin_keterangan = not any(k in ket for k in kata_izin)

        # Cek jam_kerja tidak "Tidak hadir"
        jam_kerja_oke = True
        if pd.notna(row.get('Jam_kerja')):
            jam_kerja_oke = str(row['Jam_kerja']).strip().lower() != 'tidak hadir'

        # Cek scan_pulang harus ada dan tidak lebih awal dari jam_pulang,
        # kecuali jika jam_kerja mengandung 'Libur Rutin' atau 'Tugas Kantor'
        pulang_oke = True
        jam_kerja_str = str(row.get('Jam_kerja', '')).lower()
        if pd.isna(row.get('Scan_pulang')):
            if any(exc in jam_kerja_str for exc in pengecualian_scan_pulang):
                pulang_oke = True
            else:
                pulang_oke = False
        elif pd.notna(row.get('Jam_Pulang')):
            try:
                jam_pulang = pd.to_datetime(str(row['Jam_Pulang'])).time()
                scan_pulang = pd.to_datetime(str(row['Scan_pulang'])).time()
                pulang_oke = scan_pulang >= jam_pulang
            except:
                pulang_oke = False
        else:
            pulang_oke = False

        return no_telat and no_izin_kolom and no_izin_keterangan and jam_kerja_oke and pulang_oke

    df['Bersih'] = df.apply(bersih, axis=1)

    # Hitung hadir bersih
    hadir_bersih = df[df['Bersih']].groupby('Nama').size().reset_index(name='Hadir_Tanpa_Telat_Izin')

    # Gabungkan semua
    rekap = pd.merge(hari_kerja, hadir_bersih, on='Nama', how='left')
    rekap['Hadir_Tanpa_Telat_Izin'] = rekap['Hadir_Tanpa_Telat_Izin'].fillna(0).astype(int)

    # Tandai status rajin
    rekap['Status'] = rekap.apply(
        lambda row: 'Rajin' if row['Jumlah_Hari_Kerja'] == row['Hadir_Tanpa_Telat_Izin'] else 'Tidak', axis=1
    )

    # Pilihan untuk tampilkan semua atau hanya yang rajin
    filter_rajin = st.checkbox("Tampilkan hanya karyawan rajin", value=True)

    if filter_rajin:
        rekap_tampil = rekap[rekap['Status'] == 'Rajin']
    else:
        rekap_tampil = rekap

    st.subheader("📊 Rekapan Karyawan Rajin")
    st.dataframe(rekap_tampil)

    # Filter data lengkap untuk karyawan yang rajin
    nama_rajin = rekap[rekap['Status'] == 'Rajin']['Nama'].tolist()
    df_rajin_detail = df[df['Nama'].isin(nama_rajin)]

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
    kolom_dipilih = ['Tanggal', 'Nama', 'Jam_Masuk', 'Scan_masuk', 'Jam_Pulang', 'Scan_pulang', 'Departemen']
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
