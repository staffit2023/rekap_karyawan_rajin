import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

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

    # Validasi kolom wajib
    required_columns = ['Nama', 'Tanggal']
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        st.error(f"Kolom wajib tidak ditemukan: {missing_cols}")
        st.stop()

    # Cek apakah ada data setelah cleaning
    if df.empty:
        st.warning("Tidak ada data yang valid dalam file.")
        st.stop()

    # Tampilkan data awal
    st.subheader("📄 Data Absensi (Preview)")
    st.dataframe(df)

    # Daftar kata-kata izin
    kata_izin = ['izin', 'sakit', 'cuti', 'dispensasi', 'alpha']

    # Hitung total hari kerja per karyawan
    hari_kerja = df.groupby('Nama').size().reset_index(name='Jumlah_Hari_Kerja')

    # Fungsi untuk mengkonversi waktu ke menit
    def waktu_ke_menit(waktu_str):
        """Konversi string waktu (HH:MM:SS atau HH:MM) ke menit"""
        if pd.isna(waktu_str) or waktu_str == '':
            return None
        try:
            waktu_str = str(waktu_str).strip()
            if ':' in waktu_str:
                parts = waktu_str.split(':')
                if len(parts) >= 2:
                    jam = int(parts[0])
                    menit = int(parts[1])
                    # Detik diabaikan untuk perhitungan (parts[2] jika ada)
                    return jam * 60 + menit
                else:
                    return None
            else:
                # Jika hanya jam tanpa menit
                return int(waktu_str) * 60
        except (ValueError, AttributeError):
            return None

    # Fungsi cek kehadiran bersih
    def bersih(row):
        # Cek tidak terlambat
        no_telat = pd.isna(row.get('Terlambat')) or row.get('Terlambat', 0) == 0

        # Cek kolom 'Izin' harus kosong atau 0
        no_izin_kolom = pd.isna(row.get('Izin')) or row.get('Izin', 0) == 0

        # Cek tidak ada kata izin di keterangan
        no_izin_keterangan = True
        if pd.notna(row.get('Keterangan')):
            ket = str(row['Keterangan']).lower()
            no_izin_keterangan = not any(k in ket for k in kata_izin)

        # Cek jam_kerja tidak "Tidak hadir"
        jam_kerja_oke = True
        if pd.notna(row.get('Jam_kerja')):
            jam_kerja_oke = str(row['Jam_kerja']).strip().lower() != 'tidak hadir'

        # PENGECEKAN BARU: Scan Pulang
        scan_pulang_oke = True
        
        # Cek pengecualian untuk Izin dinas dan libur rutin
        jadwal_izin_dinas = False
        jam_kerja_libur = False
        
        if pd.notna(row.get('Jadwal')):
            jadwal_str = str(row['Jadwal']).strip()
            jadwal_izin_dinas = jadwal_str == "Izin dinas (Izin keperluan kantor)"
        
        if pd.notna(row.get('Jam_kerja')):
            jam_kerja_str = str(row['Jam_kerja']).strip().lower()
            jam_kerja_libur = jam_kerja_str == "libur rutin"
        
        # Jika ada pengecualian, skip pengecekan scan pulang
        if jadwal_izin_dinas or jam_kerja_libur:
            scan_pulang_oke = True
        else:
            # Cek scan pulang harus ada dan >= jam pulang
            jam_pulang = row.get('Jam_Pulang')
            scan_pulang = row.get('Scan_pulang')
            
            # Jika ada kolom jam pulang dan scan pulang
            if pd.notna(jam_pulang) and pd.notna(scan_pulang):
                jam_pulang_menit = waktu_ke_menit(jam_pulang)
                scan_pulang_menit = waktu_ke_menit(scan_pulang)
                
                if jam_pulang_menit is not None and scan_pulang_menit is not None:
                    # Scan pulang harus >= jam pulang (tidak pulang lebih awal)
                    scan_pulang_oke = scan_pulang_menit >= jam_pulang_menit
                else:
                    scan_pulang_oke = False
            elif pd.notna(jam_pulang):
                # Jika ada jam pulang tapi tidak ada scan pulang
                scan_pulang_oke = False
            # Jika tidak ada jam pulang, tidak perlu cek scan pulang
            else:
                scan_pulang_oke = True

        return no_telat and no_izin_kolom and no_izin_keterangan and jam_kerja_oke and scan_pulang_oke

    # Tambahkan kolom untuk debugging
    df['Bersih'] = df.apply(bersih, axis=1)

    # Fungsi untuk menampilkan alasan tidak rajin (untuk debugging)
    def alasan_tidak_rajin(row):
        alasan = []
        
        # Cek terlambat
        if not (pd.isna(row.get('Terlambat')) or row.get('Terlambat', 0) == 0):
            alasan.append("Terlambat")
        
        # Cek izin kolom
        if not (pd.isna(row.get('Izin')) or row.get('Izin', 0) == 0):
            alasan.append("Ada izin di kolom")
        
        # Cek keterangan izin
        if pd.notna(row.get('Keterangan')):
            ket = str(row['Keterangan']).lower()
            if any(k in ket for k in kata_izin):
                alasan.append("Ada kata izin di keterangan")
        
        # Cek jam kerja
        if pd.notna(row.get('Jam_kerja')):
            if str(row['Jam_kerja']).strip().lower() == 'tidak hadir':
                alasan.append("Tidak hadir")
        
        # Cek scan pulang
        jadwal_izin_dinas = False
        jam_kerja_libur = False
        
        if pd.notna(row.get('Jadwal')):
            jadwal_str = str(row['Jadwal']).strip()
            jadwal_izin_dinas = jadwal_str == "Izin dinas (Izin keperluan kantor)"
        
        if pd.notna(row.get('Jam_kerja')):
            jam_kerja_str = str(row['Jam_kerja']).strip().lower()
            jam_kerja_libur = jam_kerja_str == "libur rutin"
        
        if not (jadwal_izin_dinas or jam_kerja_libur):
            jam_pulang = row.get('Jam_Pulang')
            scan_pulang = row.get('Scan_pulang')
            
            if pd.notna(jam_pulang):
                if pd.isna(scan_pulang):
                    alasan.append("Tidak ada scan pulang")
                else:
                    jam_pulang_menit = waktu_ke_menit(jam_pulang)
                    scan_pulang_menit = waktu_ke_menit(scan_pulang)
                    
                    if jam_pulang_menit is not None and scan_pulang_menit is not None:
                        if scan_pulang_menit < jam_pulang_menit:
                            alasan.append("Scan pulang lebih awal dari jam pulang")
                    else:
                        alasan.append("Format waktu tidak valid")
        
        return "; ".join(alasan) if alasan else "Rajin"

    df['Alasan'] = df.apply(alasan_tidak_rajin, axis=1)

    # Hitung hadir bersih
    hadir_bersih = df[df['Bersih']].groupby('Nama').size().reset_index(name='Hadir_Tanpa_Telat_Izin')

    # Gabungkan semua
    rekap = pd.merge(hari_kerja, hadir_bersih, on='Nama', how='left')
    rekap['Hadir_Tanpa_Telat_Izin'] = rekap['Hadir_Tanpa_Telat_Izin'].fillna(0).astype(int)

    # Tandai status rajin
    rekap['Status'] = rekap.apply(
        lambda row: 'Rajin' if row['Jumlah_Hari_Kerja'] == row['Hadir_Tanpa_Telat_Izin'] else 'Tidak Rajin', axis=1
    )

    # Hitung persentase rajin
    rekap['Persentase_Rajin'] = ((rekap['Hadir_Tanpa_Telat_Izin'] / rekap['Jumlah_Hari_Kerja']) * 100).round(2)

    # FITUR FILTER LENGKAP
    st.subheader("🔍 Filter Data")
    
    # Buat kolom untuk filter
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Filter berdasarkan status
        status_options = ['Semua', 'Rajin', 'Tidak Rajin']
        selected_status = st.selectbox("Filter Status:", status_options, index=1)
    
    with col2:
        # Filter berdasarkan departemen (jika ada kolom departemen)
        if 'Departemen' in rekap.columns or 'Departemen' in df.columns:
            # Ambil departemen dari df karena rekap tidak memiliki kolom departemen
            if 'Departemen' in df.columns:
                dept_from_df = df.groupby('Nama')['Departemen'].first().reset_index()
                rekap = pd.merge(rekap, dept_from_df, on='Nama', how='left')
            
            departemen_list = ['Semua'] + sorted([d for d in rekap['Departemen'].unique() if pd.notna(d)])
            selected_dept = st.selectbox("Filter Departemen:", departemen_list)
        else:
            selected_dept = 'Semua'
    
    with col3:
        # Filter berdasarkan jabatan (jika ada kolom jabatan)
        if 'Jabatan' in df.columns:
            # Ambil jabatan dari df
            jabatan_from_df = df.groupby('Nama')['Jabatan'].first().reset_index()
            rekap = pd.merge(rekap, jabatan_from_df, on='Nama', how='left')
            
            jabatan_list = ['Semua'] + sorted([j for j in rekap['Jabatan'].unique() if pd.notna(j)])
            selected_jabatan = st.selectbox("Filter Jabatan:", jabatan_list)
        else:
            selected_jabatan = 'Semua'
    
    # Filter berdasarkan rentang tanggal
    if 'Tanggal' in df.columns:
        st.subheader("📅 Filter Tanggal")
        
        # Konversi kolom tanggal dan format hanya tanggal (hilangkan jam)
        try:
            df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date
            min_date = df['Tanggal'].min()
            max_date = df['Tanggal'].max()
            
            col_date1, col_date2 = st.columns(2)
            with col_date1:
                start_date = st.date_input("Tanggal Mulai:", min_date, min_value=min_date, max_value=max_date)
            with col_date2:
                end_date = st.date_input("Tanggal Selesai:", max_date, min_value=min_date, max_value=max_date)
            
            # Filter data berdasarkan tanggal
            if start_date <= end_date:
                mask_tanggal = (df['Tanggal'] >= start_date) & (df['Tanggal'] <= end_date)
                df_filtered = df[mask_tanggal]
                
                # Hitung ulang statistik berdasarkan filter tanggal
                hari_kerja_filtered = df_filtered.groupby('Nama').size().reset_index(name='Jumlah_Hari_Kerja')
                
                # Terapkan fungsi bersih pada data yang difilter
                df_filtered['Bersih'] = df_filtered.apply(bersih, axis=1)
                hadir_bersih_filtered = df_filtered[df_filtered['Bersih']].groupby('Nama').size().reset_index(name='Hadir_Tanpa_Telat_Izin')
                
                # Gabungkan statistik yang difilter
                rekap = pd.merge(hari_kerja_filtered, hadir_bersih_filtered, on='Nama', how='left')
                rekap['Hadir_Tanpa_Telat_Izin'] = rekap['Hadir_Tanpa_Telat_Izin'].fillna(0).astype(int)
                rekap['Status'] = rekap.apply(
                    lambda row: 'Rajin' if row['Jumlah_Hari_Kerja'] == row['Hadir_Tanpa_Telat_Izin'] else 'Tidak Rajin', axis=1
                )
                rekap['Persentase_Rajin'] = ((rekap['Hadir_Tanpa_Telat_Izin'] / rekap['Jumlah_Hari_Kerja']) * 100).round(2)
                
                # Tambahkan kembali departemen jika ada
                if 'Departemen' in df_filtered.columns:
                    dept_from_df_filtered = df_filtered.groupby('Nama')['Departemen'].first().reset_index()
                    rekap = pd.merge(rekap, dept_from_df_filtered, on='Nama', how='left')
                
                # Tambahkan kembali jabatan jika ada
                if 'Jabatan' in df_filtered.columns:
                    jabatan_from_df_filtered = df_filtered.groupby('Nama')['Jabatan'].first().reset_index()
                    rekap = pd.merge(rekap, jabatan_from_df_filtered, on='Nama', how='left')
                
            else:
                st.error("Tanggal mulai harus lebih kecil atau sama dengan tanggal selesai!")
        except:
            st.warning("Format tanggal tidak valid. Filter tanggal dinonaktifkan.")

    # Filter berdasarkan nama karyawan
    st.subheader("👤 Filter Nama Karyawan")
    nama_list = ['Semua'] + sorted(rekap['Nama'].unique().tolist())
    selected_names = st.multiselect("Pilih Karyawan:", nama_list, default=['Semua'])

    # Terapkan semua filter
    rekap_filtered = rekap.copy()

    # Filter status
    if selected_status != 'Semua':
        rekap_filtered = rekap_filtered[rekap_filtered['Status'] == selected_status]

    # Filter departemen
    if selected_dept != 'Semua' and 'Departemen' in rekap_filtered.columns:
        rekap_filtered = rekap_filtered[rekap_filtered['Departemen'] == selected_dept]

    # Filter jabatan
    if selected_jabatan != 'Semua' and 'Jabatan' in rekap_filtered.columns:
        rekap_filtered = rekap_filtered[rekap_filtered['Jabatan'] == selected_jabatan]

    # Filter nama karyawan
    if 'Semua' not in selected_names and selected_names:
        rekap_filtered = rekap_filtered[rekap_filtered['Nama'].isin(selected_names)]

    st.subheader("📊 Rekapan Karyawan Rajin (Hasil Filter)")
    st.dataframe(rekap_filtered)

    # Update rekap_tampil untuk konsistensi dengan kode selanjutnya
    rekap_tampil = rekap_filtered

    # Statistik ringkas berdasarkan filter
    total_karyawan_filtered = len(rekap_filtered)
    karyawan_rajin_filtered = len(rekap_filtered[rekap_filtered['Status'] == 'Rajin'])
    persentase_rajin_filtered = (karyawan_rajin_filtered / total_karyawan_filtered * 100) if total_karyawan_filtered > 0 else 0

    st.subheader("📈 Statistik Hasil Filter")
    col_stat1, col_stat2, col_stat3 = st.columns(3)
    with col_stat1:
        st.metric("Total Karyawan (Filter)", total_karyawan_filtered)
    with col_stat2:
        st.metric("Karyawan Rajin (Filter)", karyawan_rajin_filtered)
    with col_stat3:
        st.metric("Persentase Rajin (Filter)", f"{persentase_rajin_filtered:.1f}%")

    # Statistik berdasarkan departemen dan jabatan (jika ada)
    if 'Departemen' in rekap_filtered.columns and len(rekap_filtered) > 0:
        st.subheader("📊 Statistik per Departemen")
        dept_stats = rekap_filtered.groupby('Departemen').agg({
            'Nama': 'count',
            'Status': lambda x: (x == 'Rajin').sum(),
            'Persentase_Rajin': 'mean'
        }).round(2)
        dept_stats.columns = ['Total_Karyawan', 'Karyawan_Rajin', 'Rata2_Persentase_Rajin']
        dept_stats['Persentase_Dept_Rajin'] = ((dept_stats['Karyawan_Rajin'] / dept_stats['Total_Karyawan']) * 100).round(2)
        st.dataframe(dept_stats)

    if 'Jabatan' in rekap_filtered.columns and len(rekap_filtered) > 0:
        st.subheader("📊 Statistik per Jabatan")
        jabatan_stats = rekap_filtered.groupby('Jabatan').agg({
            'Nama': 'count',
            'Status': lambda x: (x == 'Rajin').sum(),
            'Persentase_Rajin': 'mean'
        }).round(2)
        jabatan_stats.columns = ['Total_Karyawan', 'Karyawan_Rajin', 'Rata2_Persentase_Rajin']
        jabatan_stats['Persentase_Jabatan_Rajin'] = ((jabatan_stats['Karyawan_Rajin'] / jabatan_stats['Total_Karyawan']) * 100).round(2)
        st.dataframe(jabatan_stats)

    # Filter data lengkap berdasarkan filter yang diterapkan
    if 'Tanggal' in df.columns and 'df_filtered' in locals():
        # Jika ada filter tanggal, gunakan df_filtered
        nama_filtered = rekap_filtered['Nama'].tolist()
        df_detail_filtered = df_filtered[df_filtered['Nama'].isin(nama_filtered)]
    else:
        # Jika tidak ada filter tanggal, gunakan df original
        nama_filtered = rekap_filtered['Nama'].tolist()
        df_detail_filtered = df[df['Nama'].isin(nama_filtered)]

    # Download rekap hasil filter
    output_rekap = BytesIO()
    with pd.ExcelWriter(output_rekap, engine='xlsxwriter') as writer:
        rekap_filtered.to_excel(writer, index=False, sheet_name='Rekap_Filter')

    st.download_button(
        label="⬇️ Download Rekap Filter (Excel)",
        data=output_rekap.getvalue(),
        file_name="rekap_karyawan_filter.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Tampilkan data lengkap berdasarkan filter
    st.subheader("📑 Detail Absensi Hasil Filter")
    
    # Opsi untuk menampilkan detail
    show_detail_options = st.radio(
        "Pilihan tampilan detail:",
        ["Hasil Filter", "Semua Data Original"],
        index=0
    )
    
    if show_detail_options == "Hasil Filter":
        if len(df_detail_filtered) > 0:
            st.dataframe(df_detail_filtered)
            df_for_export = df_detail_filtered
            filename = "detail_hasil_filter.xlsx"
            sheet_name = "Detail_Filter"
        else:
            st.info("Tidak ada data yang sesuai dengan filter yang dipilih.")
            df_for_export = pd.DataFrame()
    else:
        st.dataframe(df)
        df_for_export = df
        filename = "detail_semua_data.xlsx"
        sheet_name = "Detail_Semua"

    # Export detail berdasarkan pilihan
    if not df_for_export.empty:
        # Kolom yang ingin ditampilkan untuk export
        kolom_dipilih = ['Tanggal', 'Nama', 'Jam_Masuk', 'Scan_masuk', 'Jam_Pulang', 'Scan_pulang', 'Departemen', 'Jabatan', 'Jadwal', 'Jam_kerja', 'Alasan']
        kolom_tersedia = [kol for kol in kolom_dipilih if kol in df_for_export.columns]
        df_excel = df_for_export[kolom_tersedia]

        # Simpan ke Excel
        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df_excel.to_excel(writer, index=False, sheet_name=sheet_name)

        # Tombol download Excel
        st.download_button(
            label="⬇️ Download Detail (Excel)",
            data=output_excel.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Debug section (opsional) - berdasarkan filter
    with st.expander("🔍 Debug: Lihat alasan karyawan tidak rajin (berdasarkan filter)"):
        if 'df_detail_filtered' in locals() and len(df_detail_filtered) > 0:
            df_tidak_rajin_filtered = df_detail_filtered[~df_detail_filtered['Bersih']]
            if not df_tidak_rajin_filtered.empty:
                st.dataframe(df_tidak_rajin_filtered[['Nama', 'Tanggal', 'Alasan']])
            else:
                st.success("Semua record dalam filter menunjukkan kehadiran rajin!")
        else:
            st.info("Tidak ada data untuk ditampilkan berdasarkan filter yang dipilih.")
