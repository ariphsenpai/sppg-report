#!/usr/bin/env python3
"""
SPPG Tools Runner — main entry point
Mode: CLI menu untuk akses semua tools
"""

import sys
import os
from pathlib import Path

# Add tools dir
sys.path.insert(0, str(Path(__file__).parent / "tools"))

WELCOME = """
╔══════════════════════════════════════════╗
║     🍽️  SPPG REPORT TOOLS v1.0         ║
║     Satuan Pelayanan Pemenuhan Gizi      ║
╚══════════════════════════════════════════╝
"""

MENU = """
Menu Tools:

📋 [1] Laporan Operasional Harian
     → Input data produksi, distribusi, waste, masalah

📊 [2] Laporan Mingguan
     → Summary dari data harian + analisa + rekomendasi

📋 [3] Report Penerima Manfaat
     → Dari Google Sheets / CSV / manual

📈 [4] Auto-Summary Mingguan
     → Generate mingguan otomatis dari data harian yang ada

📚 [5] JUKNIS Quick Reference
     → Akses ringkasan juknis SPPG

❌ [0] Keluar
"""


def main():
    print(WELCOME)

    while True:
        print(MENU)
        choice = input("Pilih [0-5]: ").strip()

        if choice == "1":
            from laporan_harian import main as harian
            harian()
        elif choice == "2":
            from laporan_mingguan import main as mingguan
            mingguan()
        elif choice == "3":
            from report_penerima import main as penerima
            penerima()
        elif choice == "4":
            from auto_summary import main as auto
            auto()
        elif choice == "5":
            show_juknis_ref()
        elif choice == "0":
            print("\n👋 Sampai jumpa! Gas operasional!")
            break
        else:
            print("❌ Pilihan tidak valid")

        input("\nTekan Enter untuk kembali ke menu...")


def show_juknis_ref():
    """Quick reference to JUKNIS"""
    print("""
╔══════════════════════════════════════════╗
║   📚 JUKNIS QUICK REFERENCE             ║
╚══════════════════════════════════════════╝

🏢 STRUKTUR SPPG:
  Ka.SPPG, Pengawas Gizi, Pengawas Keuangan
  Pengawas Sanitasi, Juru Masak, Asisten Lapangan
  + 40-45 relawan (persiapan, masak, porsi, packing, dll)

💰 BIAYA ACUAN (AT COST):
  • Balita/TK/SD kls 1-3 / ATS <9: Rp 8.000/porsi
  • SD 4-6/SMP/SMA/3B/Tendik: Rp 10.000/porsi
  • Operasional: Rp 3.000/porsi
  • Insentif Fasilitas SPPG: Rp 6.000.000/hari (lumpsum)

📅 HARI OPERASIONAL: 313 hari/tahun (Senin-Sabtu)
   Max 24 hari/bulan × bukan Minggu

📊 LAPORAN WAJIB:
  • Harian: Rekap porsi (30a), Penggunaan dana (30b)
  • 2 Mingguan: Laporan penggunaan dana (30c)
  • Bulanan: Summary + anggaran + PM (30d)
  • Buku Besar, Petty Cash, Bahan Pangan, Operasional (30e-30h)

👥 PENERIMA MANFAAT:
  • Peserta Didik: PAUD/TK s/d SMA/SMK + SLB + Pesantren
  • Tendik: Pendidik & Tenaga Kependidikan
  • 3B: Ibu Hamil, Ibu Menyusui, Balita 6-59bln

🔬 UJI ORGANOLEPTIK:
  • 2x: saat tiba di lokasi & sebelum dikonsumsi
  • Parameter: Rasa, Warna, Aroma, Tekstur (skor 1-5)
""")
    input("Tekan Enter untuk kembali...")


if __name__ == "__main__":
    main()
