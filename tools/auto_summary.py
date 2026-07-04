#!/usr/bin/env python3
"""
AUTO-SUMMARY MINGGUAN — Generate weekly report from daily JSON data
No manual input needed for production/distribution/waste sections
"""

import sys
from pathlib import Path
from datetime import date, datetime, timedelta
import json

sys.path.insert(0, str(Path(__file__).parent))
from laporan_mingguan import render_excel, load_week_data
from utils import (
    MINGGUAN_DIR, HARIAN_DIR, tgl_indo, tgl_str,
    save_json, load_json, init_dirs, range_minggu
)


def auto_analisa(days):
    """Auto-generate analysis from daily data"""
    analisa = {
        "temuan": "",
        "root_cause": "",
        "tindakan_korektif": ""
    }

    # Find near-miss / issues
    masalah_hari = [d for d in days if d.get("masalah") and str(d["masalah"]).strip()]
    if masalah_hari:
        worst = masalah_hari[0]
        analisa["temuan"] = f"{tgl_indo(worst['tanggal'])}: {worst['masalah']}"
        analisa["root_cause"] = worst.get("tindakan", "Perlu investigasi lebih lanjut")
        analisa["tindakan_korektif"] = worst.get("tindakan", "Evaluasi SOP dan perketat QC")

    # Waste analysis
    waste_nasi = sum(d.get("waste_nasi", 0) for d in days)
    waste_sayur = sum(d.get("waste_sayur", 0) for d in days)
    waste_lauk = sum(d.get("waste_lauk", 0) for d in days)
    waste_buah = sum(d.get("waste_buah", 0) for d in days)

    waste_data = [
        ("Sayur", waste_sayur),
        ("Karbohidrat", waste_nasi),
        ("Protein", waste_lauk),
        ("Buah", waste_buah),
    ]
    waste_data.sort(key=lambda x: x[1], reverse=True)

    if waste_data and waste_data[0][1] > 0:
        top_waste = waste_data[0]
        analisa["temuan_waste"] = f"{top_waste[0]} menjadi penyumbang waste terbesar ({top_waste[1]} kg)"

    return analisa


def auto_evaluasi_divisi(days):
    """Auto-generate division evaluations from incidents"""
    evaluasi = []

    # Check for specific issues
    ada_ulat = any("ulat" in str(d.get("masalah", "")).lower() for d in days)
    ada_komplain_rasa = any("rasa" in str(d.get("feedback_pm", "")).lower() for d in days)
    ada_komplain_kualitas = any("asam" in str(d.get("feedback_pm", "")).lower() for d in days)
    ada_komplain_karbo = any("double karbo" in str(d.get("feedback_pm", "")).lower() for d in days)

    evaluasi.append({
        "divisi": "Admin Gudang",
        "status": "Sangat Baik",
        "temuan": "-" if not ada_komplain_kualitas else "Perlu monitoring kualitas sayur yang masuk",
        "rekomendasi": "Pertahankan monitoring stok dan spesifikasi barang"
    })

    if ada_ulat:
        evaluasi.append({
            "divisi": "Persiapan",
            "status": "Perlu Evaluasi",
            "temuan": "Sayur tidak lolos QC hingga tahap pemorsian",
            "rekomendasi": "Perketat sortir sayur + tambah QC setelah pencucian"
        })
    else:
        evaluasi.append({
            "divisi": "Persiapan",
            "status": "Baik",
            "temuan": "-",
            "rekomendasi": "Pertahankan kualitas persiapan bahan"
        })

    evaluasi.append({
        "divisi": "Cooking",
        "status": "Sangat Baik",
        "temuan": "-",
        "rekomendasi": "Pertahankan kualitas masakan dan cita rasa"
    })

    evaluasi.append({
        "divisi": "Pemorsian",
        "status": "Sangat Baik" if not ada_ulat else "Baik",
        "temuan": "Berfungsi sebagai QC lapis akhir" if not ada_ulat else "Berhasil menemukan kontaminasi sebelum distribusi — fungsi kontrol berjalan",
        "rekomendasi": "Pertahankan kewaspadaan dalam pemorsian"
    })

    evaluasi.append({
        "divisi": "Distribusi",
        "status": "Sangat Baik",
        "temuan": "-",
        "rekomendasi": "Pertahankan ketepatan waktu distribusi"
    })

    evaluasi.append({
        "divisi": "Cuci Ompreng",
        "status": "Sangat Baik",
        "temuan": "-",
        "rekomendasi": "Pertahankan standar kebersihan"
    })

    return evaluasi


def auto_rekomendasi(days, analisa):
    """Auto-generate recommendations"""
    rekom = []

    # Waste-based
    waste_sayur = sum(d.get("waste_sayur", 0) for d in days)
    waste_nasi = sum(d.get("waste_nasi", 0) for d in days)

    if waste_sayur > 10:
        rekom.append("Evaluasi jenis sayur yang digunakan — kombinasikan dengan sayur yang lebih diterima")
    if waste_nasi > 10:
        rekom.append("Pertimbangkan pengurangan porsi nasi atau variasi karbohidrat non-nasi")

    # Feedback-based
    feedbacks = [d.get("feedback_pm", "") for d in days if d.get("feedback_pm")]
    for fb in feedbacks:
        if "asam" in fb.lower():
            rekom.append("Evaluasi kualitas sayur saat distribusi agar tidak terjadi perubahan rasa")
        if "double" in fb.lower() or "karbo" in fb.lower():
            rekom.append("Evaluasi keseimbangan komposisi menu — hindari persepsi double karbo")
        if "buncis" in fb.lower():
            rekom.append("Kurangi frekuensi buncis atau kombinasikan dengan sayuran lain")

    rekom.append("Pertahankan variasi menu non-nasi (kentang, jagung) sebagai alternatif karbohidrat")
    rekom.append("Pertahankan sistem kontrol pemorsian sebagai lapisan QC terakhir")

    return rekom


def auto_summary(days, analisa):
    """Auto-generate executive summary"""
    total_prod = sum(d["produksi"] for d in days)
    total_target = sum(d["target"] for d in days)
    hari_efektif = sum(1 for d in days if d["produksi"] > 0)
    tepat = sum(1 for d in days if d.get("tepat_waktu", False))

    masalah = [d for d in days if d.get("masalah") and str(d["masalah"]).strip()]
    feedbacks = [d for d in days if d.get("feedback_pm") and str(d["feedback_pm"]).strip()]

    summary = []
    summary.append(f"Periode ini merupakan minggu operasional SPPG.")
    summary.append("")
    summary.append("✅ Pencapaian:")
    summary.append(f"✅ Produksi {total_prod:,} porsi ({total_target:,} target)")
    summary.append(f"✅ Tepat waktu {(tepat/hari_efektif*100) if hari_efektif else 0:.0f}%")
    summary.append(f"✅ Tidak ada masalah distribusi besar")

    if masalah:
        summary.append("")
        summary.append("⚠️ Temuan:")
        for m in masalah[:3]:
            summary.append(f"⚠️ {m['masalah']}")

    if feedbacks:
        summary.append("")
        summary.append("💬 Feedback PM:")
        for f in feedbacks[:3]:
            summary.append(f"• {f['feedback_pm']}")

    return "\n".join(summary)


def main():
    init_dirs()

    print("\n" + "="*50)
    print("📈 AUTO-SUMMARY MINGGUAN SPPG")
    print("="*50)
    nama_sppg = input("\n🏢 Nama SPPG [SPPG Wonodri 3]: ").strip() or "SPPG Wonodri 3"
    penyusun = input("👤 Penyusun [Auto-Generated]: ").strip() or "Auto-Generated"

    # Date range
    print("\n📅 Periode (biarkan kosong untuk minggu ini):")
    tgl_start = input("  Tanggal mulai (YYYY-MM-DD): ").strip()
    tgl_end = input("  Tanggal selesai (YYYY-MM-DD): ").strip()

    days, start_d, end_d = load_week_data(
        tgl_start if tgl_start else None,
        tgl_end if tgl_end else None
    )

    if not tgl_start:
        print(f"\n  → Auto: {tgl_indo(str(start_d))} – {tgl_indo(str(end_d))}")

    # Filter out zero-production days (weekends, holidays)
    aktif_days = [d for d in days if d["produksi"] > 0]

    if not aktif_days:
        print("❌ Tidak ada data harian di periode ini. Isi laporan harian dulu ya.")
        return

    print(f"\n📂 Data harian ditemukan: {len(aktif_days)} hari aktif")

    # Auto-analisa
    analisa = auto_analisa(aktif_days)
    evaluasi = auto_evaluasi_divisi(aktif_days)
    rekom = auto_rekomendasi(aktif_days, analisa)
    summary = auto_summary(aktif_days, analisa)

    # Build report
    report = {
        "nama_sppg": nama_sppg,
        "penyusun": penyusun,
        "tgl_start": str(start_d),
        "tgl_end": str(end_d),
        "tgl_laporan": str(date.today()),
        "days": days,
        "analisa": analisa,
        "evaluasi_divisi": evaluasi,
        "rekomendasi": rekom,
        "executive_summary": summary,
        "insight": "Auto-generated from daily report data. Review manually before submission.",
        "status_performa": "Baik" if len([m for m in aktif_days if m.get("masalah")]) == 0 else "Cukup"
    }

    # Preview
    print("\n" + "="*50)
    print("📋 PREVIEW ANALISA OTOMATIS:")
    print(f"  Temuan: {analisa.get('temuan', '-')}")
    print(f"  Root Cause: {analisa.get('root_cause', '-')}")
    print(f"  Rekomendasi: {len(rekom)} item")
    print(f"  Status: {report['status_performa']}")
    print("="*50)

    print("\n🔄 Hasil auto-summary bisa di-edit manual nanti.")
    confirm = input("\nSimpan laporan? (y/n) [y]: ").strip().lower() or "y"

    if confirm != "y":
        print("❌ Dibatalkan")
        return

    # Save JSON
    tgl_key = f"{report['tgl_start']}_to_{report['tgl_end']}"
    json_path = MINGGUAN_DIR / f"auto_mingguan_{tgl_key}.json"
    save_json(json_path, report)

    # Save Excel
    xlsx_path = MINGGUAN_DIR / f"auto_laporan_mingguan_{tgl_key}.xlsx"
    render_excel(report, xlsx_path)

    print("\n" + "="*50)
    print("✅ AUTO-SUMMARY MINGGUAN TERSIMPAN")
    print(f"📄 Excel: {xlsx_path}")
    print(f"📋 JSON:  {json_path}")
    total_prod = sum(d["produksi"] for d in aktif_days)
    print(f"📊 Total Produksi: {total_prod:,} porsi")
    print("="*50)


if __name__ == "__main__":
    main()
