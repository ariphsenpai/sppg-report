# SPPG Report — Satuan Pelayanan Pemenuhan Gizi

Format laporan Excel untuk program SPPG (Satuan Pelayanan Pemenuhan Gizi).

## Struktur File Excel

| Sheet | Isi |
|---|---|
| **Ringkasan** | Dashboard utama — total penerima, total hari, total biaya, ringkasan bulanan |
| **Data Penerima** | Master data penerima manfaat (nama, usia, alamat, status) |
| **Distribusi Harian** | Log distribusi per hari per penerima |
| **Menu Makanan** | Database menu & biaya satuan |
| **Rekap Bulanan** | Rekap per penerima per bulan |
| **Rekap Anggaran** | Total pengeluaran vs anggaran |

## Cara Pakai

```bash
python3 generate.py
```

Output: `SPPG_Report_<bulan>_<tahun>.xlsx`
