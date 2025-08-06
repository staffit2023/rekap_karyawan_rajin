# Cek scan_pulang harus ada dan tidak lebih awal dari jam_pulang
pulang_oke = True
if pd.notna(row.get('Jam_Pulang')) and pd.notna(row.get('Scan_pulang')):
    try:
        jam_pulang = pd.to_datetime(str(row['Jam_Pulang'])).time()
        scan_pulang = pd.to_datetime(str(row['Scan_pulang'])).time()
        pulang_oke = scan_pulang >= jam_pulang
    except:
        pulang_oke = False
else:
    pulang_oke = False
