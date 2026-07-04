#!/usr/bin/env python3
"""
SPPG Utils — shared functions for all report tools
"""

import json
from datetime import datetime, date, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
HARIAN_DIR = DATA_DIR / "harian"
MINGGUAN_DIR = DATA_DIR / "mingguan"
PENERIMA_DIR = DATA_DIR / "penerima"
TEMPLATE_DIR = BASE_DIR / "template"

HARI_NAMA = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

# Standard categories from Juknis
KELOMPOK_BESAR = ["SD/MI Kls 4-6", "SMP/MTs", "SMA/MA/SMK", "ATS 9-18th"]
KELOMPOK_KECIL = ["PAUD/TK/RA", "SD/MI Kls 1-3", "Balita", "ATS <9th"]
KELOMPOK_TENDIK = ["Pendidik", "Tenaga Kependidikan"]
KELOMPOK_3B = ["Ibu Hamil", "Ibu Menyusui", "Balita (6-59 bln)"]


def init_dirs():
    for d in [HARIAN_DIR, MINGGUAN_DIR, PENERIMA_DIR, TEMPLATE_DIR]:
        d.mkdir(parents=True, exist_ok=True)


def tgl_str(tgl=None):
    if tgl is None:
        tgl = date.today()
    elif isinstance(tgl, str):
        return tgl
    return tgl.strftime("%Y-%m-%d")


def tgl_indo(tgl):
    """Format tanggal Indonesia: 19 Juni 2026"""
    bulan = [
        "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    if isinstance(tgl, str):
        tgl = datetime.strptime(tgl, "%Y-%m-%d").date()
    return f"{tgl.day} {bulan[tgl.month]} {tgl.year}"


def nama_hari(tgl):
    if isinstance(tgl, str):
        tgl = datetime.strptime(tgl, "%Y-%m-%d").date()
    return HARI_NAMA[tgl.weekday()]


def minggu_ke(tgl):
    """Minggu ke- dalam bulan"""
    if isinstance(tgl, str):
        tgl = datetime.strptime(tgl, "%Y-%m-%d").date()
    return (tgl.day - 1) // 7 + 1


def range_minggu(tgl):
    """Cari Senin-Jumat/Minggu dari tanggal tertentu"""
    if isinstance(tgl, str):
        tgl = datetime.strptime(tgl, "%Y-%m-%d").date()
    senin = tgl - timedelta(days=tgl.weekday())
    jumat = senin + timedelta(days=4)
    return senin, jumat


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_json(path, default=None):
    if not path.exists():
        return default if default is not None else {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def rupiah(angka):
    if angka is None or angka == "":
        return "Rp 0"
    return f"Rp {int(angka):,}".replace(",", ".")
