#!/usr/bin/env python3
"""
Aplikasi Rapor SD Kelas 1-6 - Kurikulum Merdeka
Streamlit Web App untuk Input Nilai Mudah & Generate PDF Rapor Otomatis
Sesuai Panduan Pembelajaran dan Asesmen (Format Laporan Hasil Belajar SD)
"""

import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from io import BytesIO
from datetime import datetime
from PIL import Image
import os
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==================== KONFIGURASI ====================
st.set_page_config(
    page_title="Aplikasi Rapor SD | Kelas 1-6",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': "# 📘 Aplikasi Rapor SD\nVersi PWA - Bisa di-install sebagai aplikasi di Android & Desktop"
    }
)

# Custom CSS for better look
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a5276;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #2c3e50;
        text-align: center;
        margin-bottom: 1.5rem;
    }
    .section-header {
        background-color: #eaf2f8;
        padding: 8px 12px;
        border-radius: 6px;
        font-weight: 600;
        color: #1a5276;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .stNumberInput, .stTextArea, .stTextInput, .stSelectbox {
        margin-bottom: 0.3rem;
    }
    .info-box {
        background-color: #fef9e7;
        padding: 10px;
        border-left: 4px solid #f39c12;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ==================== FUNGSI BANTU ====================
def get_fase(kelas):
    if kelas in [1, 2]:
        return "A"
    elif kelas in [3, 4]:
        return "B"
    else:
        return "C"

def get_subjects(kelas, agama_label):
    """Daftar mata pelajaran sesuai kelas (Kurikulum Merdeka)"""
    subjects = [
        f"Pendidikan Agama {agama_label} dan Budi Pekerti",
        "Pendidikan Pancasila",
        "Bahasa Indonesia",
        "Matematika",
    ]
    if kelas >= 3:
        subjects.append("Ilmu Pengetahuan Alam dan Sosial (IPAS)")
    subjects.extend([
        "Pendidikan Jasmani, Olahraga dan Kesehatan (PJOK)",
        "Seni Budaya",
    ])
    if kelas >= 5:
        subjects.append("Bahasa Inggris")
    subjects.append("Muatan Lokal")
    return subjects

def generate_deskripsi_otomatis(nama_lengkap, mapel, nilai):
    """Generate deskripsi capaian kompetensi yang positif, memotivasi, dan sesuai Kurikulum Merdeka"""
    if not nama_lengkap:
        nama_depan = "Ananda"
    else:
        nama_depan = nama_lengkap.split()[0].capitalize()
    
    # Level berdasarkan nilai
    if nilai >= 95:
        level = "sangat baik sekali"
        capaian = "telah mencapai dan melampaui semua kompetensi yang diharapkan dengan sangat baik"
        saran = "Terus kembangkan potensi luar biasa ini dengan konsisten dan mandiri."
    elif nilai >= 85:
        level = "sangat baik"
        capaian = "telah mencapai kompetensi yang diharapkan dengan sangat baik"
        saran = "Pertahankan prestasi ini dan tingkatkan lagi kemampuan analisis serta kreativitas."
    elif nilai >= 75:
        level = "baik"
        capaian = "telah mencapai sebagian besar kompetensi yang diharapkan dengan baik"
        saran = "Terus berlatih dan aktif bertanya untuk memperdalam pemahaman."
    elif nilai >= 65:
        level = "cukup"
        capaian = "telah mencapai kompetensi dasar yang diharapkan"
        saran = "Perlu meningkatkan latihan mandiri dan memanfaatkan bimbingan guru/orang tua."
    else:
        level = "perlu bimbingan lebih lanjut"
        capaian = "sedang dalam proses mencapai kompetensi yang diharapkan"
        saran = "Jangan menyerah! Dengan latihan rutin dan bimbingan, Ananda pasti bisa lebih baik."
    
    # Spesifik per mata pelajaran (ringkas & relevan)
    mapel_lower = mapel.lower()
    if "agama" in mapel_lower:
        specific = "memahami dan mengamalkan ajaran agama serta nilai-nilai budi pekerti dalam kehidupan sehari-hari"
        contoh = "Ananda sudah mampu melaksanakan ibadah dengan khusyuk dan menunjukkan sikap toleransi."
    elif "pancasila" in mapel_lower:
        specific = "memahami dan mengamalkan nilai-nilai Pancasila dalam kehidupan bermasyarakat"
        contoh = "Ananda mampu menyebutkan sila-sila Pancasila dan memberikan contoh penerapannya di sekolah dan rumah."
    elif "bahasa indonesia" in mapel_lower:
        specific = "keterampilan berbahasa Indonesia (membaca, menulis, menyimak, dan berbicara)"
        contoh = "Ananda sudah lancar membaca teks dan menulis paragraf sederhana dengan ejaan yang tepat."
    elif "matematika" in mapel_lower:
        specific = "pemahaman konsep bilangan, operasi hitung, geometri, dan pemecahan masalah"
        contoh = "Ananda mampu menyelesaikan soal cerita dan menggambar bangun datar dengan akurat."
    elif "ipas" in mapel_lower or "pengetahuan alam" in mapel_lower:
        specific = "memahami konsep IPA dan IPS serta keterampilan berpikir ilmiah dan sosial"
        contoh = "Ananda mampu mengamati fenomena alam sekitar dan menjelaskan hubungan sebab-akibat sederhana."
    elif "jasmani" in mapel_lower or "pjok" in mapel_lower:
        specific = "keterampilan gerak dasar, kebugaran jasmani, dan sportivitas"
        contoh = "Ananda aktif berpartisipasi dalam permainan dan menunjukkan sikap kerja sama yang baik."
    elif "seni" in mapel_lower:
        specific = "ekspresi kreatif melalui seni rupa, musik, tari, atau teater"
        contoh = "Ananda mampu mengekspresikan ide melalui karya seni dan menikmati berbagai bentuk seni budaya."
    elif "inggris" in mapel_lower:
        specific = "keterampilan dasar berbahasa Inggris (vocabulary, simple sentences, listening & speaking)"
        contoh = "Ananda sudah mampu memperkenalkan diri dan menyebutkan benda-benda di sekitarnya dalam bahasa Inggris sederhana."
    else:  # Muatan Lokal
        specific = "pemahaman dan pelestarian budaya serta kearifan lokal daerah"
        contoh = "Ananda menunjukkan antusiasme dalam mempelajari bahasa dan tradisi daerah."
    
    deskripsi = (
        f"{nama_depan} menunjukkan perkembangan {level} dalam {specific}. "
        f"Ananda {capaian}. {contoh} {saran} "
        f"Semangat terus belajar, Ananda memiliki potensi besar untuk berkembang lebih optimal!"
    )
    return deskripsi


# ==================== KETERANGAN MATA PELAJARAN (untuk bantuan input nilai) ====================
SUBJECT_KETERANGAN = {
    "Pendidikan Agama": "Membentuk karakter religius, pemahaman ajaran agama, ibadah, dan akhlak mulia sesuai keyakinan siswa.",
    "Pendidikan Pancasila": "Mengembangkan pemahaman & pengamalan nilai-nilai Pancasila, karakter bangsa, dan kehidupan bermasyarakat yang demokratis.",
    "Bahasa Indonesia": "Melatih 4 keterampilan berbahasa (menyimak, berbicara, membaca, menulis) serta menghargai sastra dan budaya Indonesia.",
    "Matematika": "Mengembangkan kemampuan berpikir logis, pemecahan masalah, konsep bilangan, geometri, pengukuran, dan statistika dasar.",
    "Ilmu Pengetahuan Alam dan Sosial (IPAS)": "Mengenal konsep IPA & IPS secara terpadu, mengembangkan keterampilan berpikir ilmiah, observasi, dan kesadaran sosial-lingkungan.",
    "Pendidikan Jasmani, Olahraga dan Kesehatan (PJOK)": "Mengembangkan kebugaran jasmani, keterampilan gerak dasar, sportivitas, kerja sama tim, dan gaya hidup sehat.",
    "Seni Budaya": "Mengekspresikan kreativitas melalui seni rupa, musik, tari, dan teater serta menghargai keberagaman seni budaya Nusantara.",
    "Bahasa Inggris": "Mengenal kosakata dasar, kalimat sederhana, listening & speaking untuk berkomunikasi dalam konteks kehidupan sehari-hari.",
    "Muatan Lokal": "Mempelajari dan melestarikan bahasa daerah, budaya, tradisi, serta kearifan lokal sesuai potensi daerah masing-masing.",
}


def create_rapor_pdf(data):
    """Generate PDF Rapor menggunakan reportlab canvas (layout profesional & rapi)"""
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # Margin
    left_margin = 1.5 * cm
    right_margin = 1.5 * cm
    top_margin = 1.2 * cm
    usable_width = width - left_margin - right_margin
    
    y = height - top_margin
    
    # ========== HEADER ==========
    # Kotak header
    c.setFillColor(colors.HexColor("#1a5276"))
    c.rect(0, height - 2.8*cm, width, 2.8*cm, fill=1, stroke=0)
    
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 11)
    c.drawCentredString(width/2, height - 0.9*cm, "KEMENTERIAN PENDIDIKAN, KEBUDAYAAN, RISET, DAN TEKNOLOGI")
    c.setFont("Helvetica", 9)
    c.drawCentredString(width/2, height - 1.35*cm, "REPUBLIK INDONESIA")
    
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width/2, height - 2.0*cm, "RAPOR PESERTA DIDIK")
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width/2, height - 2.45*cm, "SEKOLAH DASAR (SD)")
    
    y = height - 3.3*cm
    
    # Nama Sekolah
    c.setFillColor(colors.HexColor("#1a5276"))
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width/2, y, data.get('nama_sekolah', 'SEKOLAH DASAR'))
    y -= 0.45*cm
    c.setFont("Helvetica", 9)
    c.setFillColor(colors.black)
    c.drawCentredString(width/2, y, f"NPSN: {data.get('npsn', '-')} | {data.get('alamat_sekolah', '')}")
    y -= 0.7*cm
    
    # ========== LOGO SEKOLAH (opsional) - digambar setelah header agar tidak tertutup ==========
    logo_bytes = data.get('logo_bytes')
    if logo_bytes:
        try:
            logo_reader = ImageReader(BytesIO(logo_bytes))
            logo_w = 2.0 * cm
            logo_h = 2.0 * cm
            # Posisi pojok kiri atas, overlay di area header biru sebelah kiri
            c.drawImage(logo_reader, left_margin + 0.3*cm, height - 2.55*cm,
                        width=logo_w, height=logo_h,
                        preserveAspectRatio=True, mask='auto')
        except Exception:
            pass  # abaikan jika logo bermasalah
    
    # Garis pemisah
    c.setStrokeColor(colors.HexColor("#1a5276"))
    c.setLineWidth(1.5)
    c.line(left_margin, y, width - right_margin, y)
    y -= 0.6*cm
    
    # ========== IDENTITAS SISWA ==========
    c.setFillColor(colors.HexColor("#1a5276"))
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin, y, "IDENTITAS PESERTA DIDIK")
    y -= 0.45*cm
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 9)
    
    # Dua kolom identitas - RAPi & PROFESIONAL: semua titik dua (colon) disejajarkan secara vertikal
    # Ini membuat tampilan lebih rapi, seperti formulir resmi (label kiri, colon sejajar, nilai setelahnya)
    col1_x = left_margin
    col2_x = left_margin + usable_width * 0.52
    
    # Data identitas (label, nilai) - pisahkan agar mudah diatur posisi colon
    identitas_kiri = [
        ("Nama Peserta Didik", data.get('nama_siswa', '-')),
        ("NISN", data.get('nisn', '-')),
        ("Kelas / Fase", f"{data.get('kelas', '-')} / Fase {data.get('fase', '-')}"),
    ]
    identitas_kanan = [
        ("Semester", data.get('semester', '-')),
        ("Tahun Ajaran", data.get('tahun_ajaran', '-')),
        ("Agama", data.get('agama', '-')),
    ]
    
    # Posisi horizontal titik dua yang KONSISTEN (supaya sejajar vertikal di setiap kolom)
    # Nilai ini sudah disesuaikan dengan lebar font Helvetica 9pt + padding rapi
    colon_x_kiri = col1_x + 3.35 * cm      # colon untuk kolom kiri (cukup untuk "Nama Peserta Didik")
    value_x_kiri = colon_x_kiri + 0.28 * cm
    
    colon_x_kanan = col2_x + 2.35 * cm     # colon untuk kolom kanan (cukup untuk "Tahun Ajaran")
    value_x_kanan = colon_x_kanan + 0.28 * cm
    
    line_height = 0.45 * cm   # sedikit lebih lega untuk keterbacaan
    
    for i, (label, val) in enumerate(identitas_kiri):
        y_pos = y - i * line_height
        c.drawString(col1_x, y_pos, label)
        c.drawString(colon_x_kiri, y_pos, ":")
        c.drawString(value_x_kiri, y_pos, str(val))
    
    for i, (label, val) in enumerate(identitas_kanan):
        y_pos = y - i * line_height
        c.drawString(col2_x, y_pos, label)
        c.drawString(colon_x_kanan, y_pos, ":")
        c.drawString(value_x_kanan, y_pos, str(val))
    
    y -= 3 * line_height + 0.15 * cm   # spacing setelah section identitas
    
    # Garis tipis
    c.setStrokeColor(colors.grey)
    c.setLineWidth(0.5)
    c.line(left_margin, y, width - right_margin, y)
    y -= 0.5*cm
    
    # ========== A. NILAI DAN CAPAIAN KOMPETENSI ==========
    c.setFillColor(colors.HexColor("#1a5276"))
    c.setFont("Helvetica-Bold", 10)
    c.drawString(left_margin, y, "A. NILAI DAN CAPAIAN KOMPETENSI")
    y -= 0.5*cm
    
    # Header tabel
    c.setFillColor(colors.HexColor("#d5dbdb"))
    c.rect(left_margin, y - 0.35*cm, usable_width, 0.55*cm, fill=1, stroke=0)
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 8)
    
    # Kolom lebar - proporsional & rapi (lebih banyak ruang untuk deskripsi capaian)
    col_no = left_margin + 0.15*cm
    col_mapel = left_margin + 0.55*cm
    mapel_width = usable_width * 0.30
    col_nilai = left_margin + usable_width * 0.37
    col_desk = left_margin + usable_width * 0.46
    desk_width = usable_width * 0.54
    
    c.drawString(col_no, y - 0.2*cm, "No")
    c.drawString(col_mapel, y - 0.2*cm, "Mata Pelajaran")
    c.drawCentredString(col_nilai + 0.6*cm, y - 0.2*cm, "Nilai Akhir")
    c.drawString(col_desk, y - 0.2*cm, "Capaian Kompetensi")
    
    y -= 0.55*cm
    
    # Isi tabel nilai
    subjects = data.get('subjects', [])
    nilai_list = data.get('nilai_list', [])
    deskripsi_list = data.get('deskripsi_list', [])
    
    c.setFont("Helvetica", 7.5)
    row_height = 1.50*cm  # lebih lega untuk keterbacaan & deskripsi multi-baris
    
    for idx, (mapel, nilai, desk) in enumerate(zip(subjects, nilai_list, deskripsi_list)):
        # Background selang-seling
        if idx % 2 == 0:
            c.setFillColor(colors.HexColor("#f8f9f9"))
            c.rect(left_margin, y - row_height + 0.1*cm, usable_width, row_height, fill=1, stroke=0)
        
        c.setFillColor(colors.black)
        
        # No
        c.drawCentredString(col_no + 0.25*cm, y - 0.3*cm, str(idx + 1))
        
        # Mata Pelajaran (wrap jika panjang) - gunakan mapel_width dinamis
        c.setFont("Helvetica", 7.5)
        mapel_lines = []
        words = mapel.split()
        current_line = ""
        for word in words:
            test_line = current_line + " " + word if current_line else word
            if c.stringWidth(test_line, "Helvetica", 7.5) < mapel_width:
                current_line = test_line
            else:
                if current_line:
                    mapel_lines.append(current_line)
                current_line = word
        if current_line:
            mapel_lines.append(current_line)
        
        mapel_y = y - 0.22*cm
        for line in mapel_lines[:3]:  # max 3 baris
            c.drawString(col_mapel, mapel_y, line)
            mapel_y -= 0.26*cm
        
        # Nilai
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(col_nilai + 0.6*cm, y - 0.45*cm, str(nilai))
        
        # Deskripsi - wrap DINAMIS berdasarkan lebar kolom (bukan hardcoded char!)
        # Ini memperbaiki teks yang sebelumnya overflow atau terpotong
        c.setFont("Helvetica", 6.8)
        desk_lines = []
        words = desk.split()
        current_line = ""
        for word in words:
            test_line = current_line + " " + word if current_line else word
            if c.stringWidth(test_line, "Helvetica", 6.8) < desk_width - 0.15*cm:
                current_line = test_line
            else:
                if current_line:
                    desk_lines.append(current_line)
                current_line = word
        if current_line:
            desk_lines.append(current_line)
        
        desk_y = y - 0.18*cm
        for line in desk_lines[:5]:  # max 5 baris (lebih fleksibel)
            c.drawString(col_desk, desk_y, line.strip())
            desk_y -= 0.23*cm
        
        y -= row_height
    
    # Border tabel + garis vertikal pemisah kolom (tampilan lebih profesional & rapi)
    c.setStrokeColor(colors.HexColor("#1a5276"))
    c.setLineWidth(0.8)
    table_height = (len(subjects) * row_height) + 0.55*cm
    c.rect(left_margin, y, usable_width, table_height, fill=0, stroke=1)
    
    # Garis vertikal pemisah antar kolom tabel
    c.setLineWidth(0.4)
    c.line(col_mapel - 0.05*cm, y, col_mapel - 0.05*cm, y + table_height)
    c.line(col_nilai - 0.15*cm, y, col_nilai - 0.15*cm, y + table_height)
    c.line(col_desk - 0.1*cm, y, col_desk - 0.1*cm, y + table_height)
    
    y -= 0.55*cm
    
    # ========== B/C. CAPAIAN PEMBELAJARAN (CP) & TUJUAN PEMBELAJARAN (TP) (opsional) ==========
    cp_tp = data.get('cp_tp_ringkasan', '').strip()
    cp_shown = False
    if cp_tp and cp_tp != "":
        c.setFillColor(colors.HexColor("#1a5276"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(left_margin, y, "B. CAPAIAN PEMBELAJARAN (CP) & TUJUAN PEMBELAJARAN (TP)")
        y -= 0.4*cm
        
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 8)
        
        # Wrap teks CP/TP (lebar dinamis, sudah rapi)
        cp_lines = []
        words = cp_tp.split()
        current = ""
        for w in words:
            if c.stringWidth(current + " " + w, "Helvetica", 8) < usable_width - 0.3*cm:
                current = current + " " + w if current else w
            else:
                cp_lines.append(current)
                current = w
        if current:
            cp_lines.append(current)
        
        for line in cp_lines[:6]:  # max 6 baris
            c.drawString(left_margin, y, line)
            y -= 0.30*cm
        
        y -= 0.25*cm
        cp_shown = True
    
    # ========== C/D. KETIDAKHADIRAN ==========
    c.setFillColor(colors.HexColor("#1a5276"))
    c.setFont("Helvetica-Bold", 9)
    ket_label = "C. KETIDAKHADIRAN" if cp_shown else "B. KETIDAKHADIRAN"
    c.drawString(left_margin, y, ket_label)
    y -= 0.38*cm
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8)
    
    sakit = data.get('sakit', 0)
    izin = data.get('izin', 0)
    tanpa = data.get('tanpa_keterangan', 0)
    
    c.drawString(left_margin, y, f"Sakit              : {sakit} hari")
    c.drawString(left_margin + 5*cm, y, f"Izin               : {izin} hari")
    c.drawString(left_margin + 10*cm, y, f"Tanpa Keterangan : {tanpa} hari")
    y -= 0.55*cm
    
    # ========== D/E. CATATAN WALI KELAS ==========
    c.setFillColor(colors.HexColor("#1a5276"))
    c.setFont("Helvetica-Bold", 9)
    cat_label = "D. CATATAN WALI KELAS" if cp_shown else "C. CATATAN WALI KELAS"
    c.drawString(left_margin, y, cat_label)
    y -= 0.32*cm
    
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8)
    catatan = data.get('catatan_wali', 'Ananda menunjukkan perkembangan yang baik selama satu semester ini. Pertahankan semangat belajar dan tingkatkan kedisiplinan.')
    
    # Wrap catatan (lebar dinamis)
    catatan_lines = []
    words = catatan.split()
    current = ""
    for w in words:
        if c.stringWidth(current + " " + w, "Helvetica", 8) < usable_width - 0.5*cm:
            current = current + " " + w if current else w
        else:
            catatan_lines.append(current)
            current = w
    if current:
        catatan_lines.append(current)
    
    for line in catatan_lines[:5]:
        c.drawString(left_margin, y, line)
        y -= 0.30*cm
    
    y -= 0.25*cm
    
    # ========== TANDA TANGAN ==========
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.5)
    
    # Posisi tanda tangan
    sig_width = usable_width / 3
    sig_y = y - 0.3*cm
    
    # Kolom 1: Wali Kelas
    c.setFont("Helvetica", 8)
    c.drawCentredString(left_margin + sig_width/2, sig_y, "Mengetahui,")
    c.drawCentredString(left_margin + sig_width/2, sig_y - 0.35*cm, "Wali Kelas")
    c.line(left_margin + 0.8*cm, sig_y - 2.2*cm, left_margin + sig_width - 0.8*cm, sig_y - 2.2*cm)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(left_margin + sig_width/2, sig_y - 2.5*cm, data.get('nama_wali_kelas', '........................'))
    c.setFont("Helvetica", 7)
    c.drawCentredString(left_margin + sig_width/2, sig_y - 2.75*cm, f"NIP. {data.get('nip_wali', '-')}")
    
    # Kolom 2: Kepala Sekolah
    c.setFont("Helvetica", 8)
    c.drawCentredString(left_margin + sig_width + sig_width/2, sig_y, "Kepala Sekolah")
    c.line(left_margin + sig_width + 0.8*cm, sig_y - 2.2*cm, left_margin + 2*sig_width - 0.8*cm, sig_y - 2.2*cm)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(left_margin + sig_width + sig_width/2, sig_y - 2.5*cm, data.get('nama_kepala', '........................'))
    c.setFont("Helvetica", 7)
    c.drawCentredString(left_margin + sig_width + sig_width/2, sig_y - 2.75*cm, f"NIP. {data.get('nip_kepala', '-')}")
    
    # Kolom 3: Orang Tua
    c.setFont("Helvetica", 8)
    c.drawCentredString(left_margin + 2*sig_width + sig_width/2, sig_y, "Orang Tua / Wali Murid")
    c.line(left_margin + 2*sig_width + 0.8*cm, sig_y - 2.2*cm, left_margin + 3*sig_width - 0.8*cm, sig_y - 2.2*cm)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(left_margin + 2*sig_width + sig_width/2, sig_y - 2.5*cm, "........................")
    
    # Tanggal di bawah
    y_sig = sig_y - 3.3*cm
    c.setFont("Helvetica", 8)
    tempat_tanggal = data.get('tempat_tanggal', f"{data.get('kota', '....................')}, {datetime.now().strftime('%d %B %Y')}")
    c.drawRightString(width - right_margin, y_sig, tempat_tanggal)
    
    # Footer kecil
    c.setFont("Helvetica-Oblique", 6)
    c.setFillColor(colors.grey)
    c.drawCentredString(width/2, 0.8*cm, "Dokumen ini dihasilkan secara otomatis oleh Aplikasi Rapor SD | Format sesuai Panduan Kurikulum Merdeka")
    
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# ==================== FUNGSI BATCH EXCEL (BARU) ====================
def create_excel_template():
    """Membuat template Excel siap pakai untuk input data banyak siswa"""
    wb = Workbook()
    
    # Sheet 1: Instruksi
    ws_inst = wb.active
    ws_inst.title = "INSTRUKSI"
    
    ws_inst['A1'] = "PETUNJUK PENGGUNAAN TEMPLATE RAPOR BATCH"
    ws_inst['A1'].font = Font(bold=True, size=14, color="1a5276")
    ws_inst.merge_cells('A1:G1')
    
    instructions = [
        "",
        "1. Isi data pada sheet 'Data_Siswa' (satu baris = satu siswa)",
        "2. Kolom wajib: Nama_Lengkap, NISN, Kelas (1-6), Semester, Tahun_Ajaran, Agama",
        "3. Isi nilai hanya pada kolom yang relevan dengan kelas siswa tersebut",
        "4. Kolom Nilai yang kosong akan diabaikan",
        "5. Setelah selesai, simpan file Excel ini",
        "6. Kembali ke aplikasi → Tab 'Batch Excel' → Upload file ini",
        "7. Klik 'Generate Semua PDF' → Anda akan mendapat file ZIP berisi semua rapor PDF",
        "",
        "Catatan Penting:",
        "- Agama harus ditulis persis: Islam, Kristen Protestan, Katolik, Hindu, Buddha, atau Konghucu",
        "- Semester: tulis '1 (Ganjil)' atau '2 (Genap)'",
        "- Tahun_Ajaran contoh: 2025/2026",
        "- Untuk deskripsi: sistem akan generate otomatis berdasarkan nilai (sangat disarankan)",
        "- Jika ingin deskripsi custom, gunakan mode Single Student atau edit PDF setelah dihasilkan",
        "",
        "Kolom Nilai yang tersedia:",
        "- Nilai_Agama, Nilai_Pancasila, Nilai_Bahasa_Indonesia, Nilai_Matematika",
        "- Nilai_IPAS (untuk kelas 3+), Nilai_PJOK, Nilai_Seni_Budaya",
        "- Nilai_Bahasa_Inggris (untuk kelas 5+), Nilai_Muatan_Lokal",
    ]
    
    for i, line in enumerate(instructions, start=3):
        ws_inst[f'A{i}'] = line
        if line.startswith("Catatan") or line.startswith("Kolom Nilai"):
            ws_inst[f'A{i}'].font = Font(bold=True)
    
    ws_inst.column_dimensions['A'].width = 100
    
    # Sheet 2: Data_Siswa (Template utama)
    ws = wb.create_sheet("Data_Siswa")
    
    # Header styling
    header_fill = PatternFill(start_color="1a5276", end_color="1a5276", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Header columns
    headers = [
        "No", "Nama_Lengkap", "NISN", "Kelas", "Semester", "Tahun_Ajaran", "Agama",
        "Nilai_Agama", "Nilai_Pancasila", "Nilai_Bahasa_Indonesia", "Nilai_Matematika",
        "Nilai_IPAS", "Nilai_PJOK", "Nilai_Seni_Budaya", "Nilai_Bahasa_Inggris", "Nilai_Muatan_Lokal",
        "Sakit", "Izin", "Tanpa_Keterangan",
        "Catatan_Wali_Kelas",
        "Nama_Wali_Kelas", "NIP_Wali_Kelas",
        "Nama_Kepala_Sekolah", "NIP_Kepala_Sekolah",
        "Tempat_Tanggal",
        "Ringkasan_CP_TP"
    ]
    
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    
    # Sample data row (contoh)
    sample_data = [
        1, "Ahmad Fauzan", "1234567890", 3, "1 (Ganjil)", "2025/2026", "Islam",
        88, 85, 90, 82, 87, 91, 84, "", 80,   # nilai (IPAS & Inggris kosong karena contoh)
        2, 0, 0,
        "Ananda menunjukkan perkembangan yang sangat baik. Pertahankan!",
        "Siti Aminah, S.Pd.", "198505152010012003",
        "Drs. Budi Santoso, M.Pd.", "196812101990011001",
        "Jakarta, 18 Juni 2026",
        "Pada semester ini peserta didik telah mencapai Capaian Pembelajaran Fase B dengan fokus pada penguatan literasi dan numerasi."
    ]
    
    for col, value in enumerate(sample_data, start=1):
        cell = ws.cell(row=2, column=col, value=value)
        cell.border = thin_border
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    
    # Set column widths
    col_widths = {
        'A': 5, 'B': 22, 'C': 14, 'D': 8, 'E': 14, 'F': 14, 'G': 18,
        'H': 12, 'I': 14, 'J': 18, 'K': 14, 'L': 12, 'M': 12, 'N': 14, 'O': 16, 'P': 14,
        'Q': 8, 'R': 8, 'S': 16,
        'T': 45,
        'U': 20, 'V': 20,
        'W': 22, 'X': 20,
        'Y': 22
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width
    
    # Row height for header
    ws.row_dimensions[1].height = 35
    
    # Add more empty rows for user to fill (20 rows ready)
    for row in range(3, 23):
        ws.cell(row=row, column=1, value=row-1)
        for col in range(1, len(headers)+1):
            ws.cell(row=row, column=col).border = thin_border
    
    # Freeze header
    ws.freeze_panes = 'A2'
    
    # Sheet 3: Contoh Hasil (optional info)
    ws3 = wb.create_sheet("Contoh_Hasil")
    ws3['A1'] = "Contoh: Setelah generate, Anda akan mendapat ZIP berisi file PDF terpisah untuk setiap siswa"
    ws3['A1'].font = Font(bold=True)
    ws3.column_dimensions['A'].width = 80
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def process_batch_excel(df, nama_sekolah="SD Negeri Contoh", npsn="00000000", alamat_sekolah="", kota="Kota Contoh", logo_bytes=None):
    """
    Memproses DataFrame dari Excel dan menghasilkan ZIP berisi semua PDF rapor.
    Menggunakan auto-generate deskripsi berdasarkan nilai.
    School identity diambil dari parameter UI batch (sama untuk semua siswa).
    logo_bytes: bytes dari file logo yang sama untuk semua rapor (opsional).
    """
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
        success_count = 0
        error_list = []
        
        for idx, row in df.iterrows():
            try:
                # Ambil data dasar
                nama = str(row.get('Nama_Lengkap', '')).strip()
                if not nama or nama == 'nan':
                    continue
                
                nisn = str(row.get('NISN', '')).strip()
                kelas = int(row.get('Kelas', 1))
                semester = str(row.get('Semester', '1 (Ganjil)')).strip()
                tahun_ajaran = str(row.get('Tahun_Ajaran', '2025/2026')).strip()
                agama = str(row.get('Agama', 'Islam')).strip()
                
                # Siapkan data dict seperti single mode - gunakan identitas sekolah dari parameter batch
                data = {
                    'nama_sekolah': nama_sekolah,
                    'npsn': npsn,
                    'alamat_sekolah': alamat_sekolah,
                    'kota': kota,
                    'nama_siswa': nama,
                    'nisn': nisn,
                    'kelas': kelas,
                    'fase': get_fase(kelas),
                    'semester': semester,
                    'tahun_ajaran': tahun_ajaran,
                    'agama': agama,
                    'logo_bytes': logo_bytes,
                }
                
                # Bangun daftar mapel & nilai berdasarkan kelas
                subjects = get_subjects(kelas, agama)
                nilai_list = []
                deskripsi_list = []
                
                # Mapping nama kolom Excel ke mapel
                mapel_to_col = {
                    f"Pendidikan Agama {agama} dan Budi Pekerti": "Nilai_Agama",
                    "Pendidikan Pancasila": "Nilai_Pancasila",
                    "Bahasa Indonesia": "Nilai_Bahasa_Indonesia",
                    "Matematika": "Nilai_Matematika",
                    "Ilmu Pengetahuan Alam dan Sosial (IPAS)": "Nilai_IPAS",
                    "Pendidikan Jasmani, Olahraga dan Kesehatan (PJOK)": "Nilai_PJOK",
                    "Seni Budaya": "Nilai_Seni_Budaya",
                    "Bahasa Inggris": "Nilai_Bahasa_Inggris",
                    "Muatan Lokal": "Nilai_Muatan_Lokal",
                }
                
                for mapel in subjects:
                    col_name = mapel_to_col.get(mapel, "")
                    nilai = row.get(col_name, 0)
                    if pd.isna(nilai) or nilai == "" or nilai == 0:
                        nilai = 75  # default jika kosong
                    
                    nilai = int(float(nilai))
                    nilai_list.append(nilai)
                    
                    # Generate deskripsi otomatis
                    desk = generate_deskripsi_otomatis(nama, mapel, nilai)
                    deskripsi_list.append(desk)
                
                data['subjects'] = subjects
                data['nilai_list'] = nilai_list
                data['deskripsi_list'] = deskripsi_list
                
                # Kehadiran
                data['sakit'] = int(row.get('Sakit', 0) or 0)
                data['izin'] = int(row.get('Izin', 0) or 0)
                data['tanpa_keterangan'] = int(row.get('Tanpa_Keterangan', 0) or 0)
                
                # Catatan & Tanda tangan
                catatan = str(row.get('Catatan_Wali_Kelas', '')).strip()
                if not catatan or catatan == 'nan':
                    catatan = "Ananda menunjukkan perkembangan yang baik selama semester ini. Pertahankan semangat belajar dan tingkatkan kedisiplinan."
                data['catatan_wali'] = catatan
                
                data['nama_wali_kelas'] = str(row.get('Nama_Wali_Kelas', '........................')).strip()
                data['nip_wali'] = str(row.get('NIP_Wali_Kelas', '-')).strip()
                data['nama_kepala'] = str(row.get('Nama_Kepala_Sekolah', '........................')).strip()
                data['nip_kepala'] = str(row.get('NIP_Kepala_Sekolah', '-')).strip()
                
                tempat_tanggal = str(row.get('Tempat_Tanggal', '')).strip()
                if not tempat_tanggal or tempat_tanggal == 'nan':
                    tempat_tanggal = f"{data['kota']}, {datetime.now().strftime('%d %B %Y')}"
                data['tempat_tanggal'] = tempat_tanggal
                
                # CP & TP ringkasan (opsional dari Excel)
                cp_tp = str(row.get('Ringkasan_CP_TP', '')).strip()
                if not cp_tp or cp_tp == 'nan':
                    cp_tp = f"Pada semester ini peserta didik telah mencapai Capaian Pembelajaran Fase {get_fase(kelas)} pada mata pelajaran yang diajarkan."
                data['cp_tp_ringkasan'] = cp_tp
                
                # Generate PDF
                pdf_bytes = create_rapor_pdf(data)
                
                # Nama file PDF
                safe_nama = nama.replace(" ", "_").replace("/", "-")
                pdf_filename = f"Rapor_{safe_nama}_Kelas{kelas}_{semester.replace(' ', '')}_{tahun_ajaran.replace('/', '-')}.pdf"
                
                zipf.writestr(pdf_filename, pdf_bytes)
                success_count += 1
                
            except Exception as e:
                error_list.append(f"Baris {idx+2}: {nama} - Error: {str(e)}")
        
        # Tambahkan file ringkasan
        summary = f"""RINGKASAN GENERATE RAPOR BATCH
================================
Tanggal Generate : {datetime.now().strftime('%d %B %Y %H:%M')}
Total Siswa Berhasil : {success_count}
Total Error : {len(error_list)}

"""
        if error_list:
            summary += "DAFTAR ERROR:\n" + "\n".join(error_list)
        
        zipf.writestr("RINGKASAN_GENERATE.txt", summary.encode('utf-8'))
    
    zip_buffer.seek(0)
    return zip_buffer.getvalue(), success_count, error_list


# ==================== UI UTAMA ====================
st.markdown('<h1 class="main-header">📘 APLIKASI RAPOR SD</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Kelas 1–6 | Kurikulum Merdeka | Input Mudah • Deskripsi Otomatis • PDF Siap Unduh</p>', unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.header("⚙️ Pengaturan")
    st.info("Aplikasi ini mendukung semua mata pelajaran SD sesuai Kurikulum Merdeka (termasuk IPAS mulai kelas 3 & Bahasa Inggris kelas 5-6).")
    
    st.markdown("---")
    st.subheader("Petunjuk Singkat")
    st.markdown("""
    1. Isi **Identitas Sekolah** & **Siswa**
    2. Pilih **Kelas** & **Agama** (untuk mata pelajaran Agama)
    3. Input **Nilai Akhir** (0-100)
    4. Edit atau gunakan **Deskripsi Otomatis** yang sudah sesuai kaidah Kurikulum Merdeka
    5. Lengkapi kehadiran & catatan
    6. Klik **Generate PDF** → Langsung unduh!
    """)
    
    st.markdown("---")
    st.caption("Versi 1.0 | Format mengikuti Panduan Pembelajaran dan Asesmen 2025")

    # === PWA INSTALLATION GUIDE ===
    with st.expander("📱 Install sebagai Aplikasi di Android (PWA)", expanded=False):
        st.markdown("""
        **Cara Install di HP Android:**
        
        1. Deploy aplikasi ini ke internet (bisa pakai Streamlit Cloud gratis)
        2. Buka link aplikasi menggunakan **Chrome** di HP Android
        3. Ketuk ikon **⋮** (tiga titik) di pojok kanan atas
        4. Pilih **"Add to Home screen"** atau **"Install app"**
        5. Konfirmasi → Ikon aplikasi akan muncul di layar utama HP Anda
        
        Setelah di-install, aplikasi akan terbuka dalam mode fullscreen seperti aplikasi native.
        """)

# Tab utama
tab1, tab2, tab3 = st.tabs(["📝 Input Data Rapor (Single)", "📦 Batch Banyak Siswa (Excel)", "📖 Panduan & Contoh"])

with tab1:
    # === IDENTITAS SEKOLAH ===
    st.markdown('<div class="section-header">🏫 IDENTITAS SEKOLAH</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        nama_sekolah = st.text_input("Nama Sekolah Dasar", value="SD Negeri 01 Contoh", key="nama_sekolah")
        npsn = st.text_input("NPSN", value="12345678", key="npsn")
    with col2:
        alamat_sekolah = st.text_input("Alamat Sekolah", value="Jl. Pendidikan No. 1, Kota Contoh", key="alamat")
        kota = st.text_input("Kota/Kabupaten", value="Kota Contoh", key="kota")
    
    # Logo upload (opsional) - akan muncul di pojok kiri atas PDF
    logo_file = st.file_uploader(
        "🖼️ Upload Logo Sekolah (PNG / JPG, opsional)", 
        type=["png", "jpg", "jpeg"],
        help="Logo akan ditampilkan di pojok kiri atas halaman rapor. Gunakan logo resmi sekolah dengan background transparan atau putih untuk hasil terbaik. Ukuran ideal: persegi."
    )
    
    # === IDENTITAS SISWA ===
    st.markdown('<div class="section-header">👤 IDENTITAS PESERTA DIDIK</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        nama_siswa = st.text_input("Nama Lengkap Siswa", value="Ahmad Fauzan", key="nama_siswa")
    with col2:
        nisn = st.text_input("NISN", value="1234567890", key="nisn")
    with col3:
        kelas = st.selectbox("Kelas", options=list(range(1, 7)), index=2, key="kelas")  # default kelas 3
    
    col1, col2, col3 = st.columns(3)
    with col1:
        semester = st.selectbox("Semester", ["1 (Ganjil)", "2 (Genap)"], index=0, key="semester")
    with col2:
        tahun_ajaran = st.text_input("Tahun Ajaran", value="2025/2026", key="tahun_ajaran")
    with col3:
        agama = st.selectbox("Agama (untuk mapel Agama)", 
                             ["Islam", "Kristen Protestan", "Katolik", "Hindu", "Buddha", "Konghucu"], 
                             index=0, key="agama")
    
    fase = get_fase(kelas)
    st.caption(f"Fase otomatis: **Fase {fase}** (sesuai kelas {kelas})")
    
    # === MATA PELAJARAN & NILAI ===
    st.markdown('<div class="section-header">📊 NILAI AKHIR & CAPAIAN KOMPETENSI</div>', unsafe_allow_html=True)
    
    agama_label = agama
    subjects = get_subjects(kelas, agama_label)
    
    st.caption(f"📌 {len(subjects)} mata pelajaran untuk Kelas {kelas} (Fase {fase})")
    
    # Inisialisasi session state untuk nilai & deskripsi
    if 'nilai_data' not in st.session_state or st.session_state.get('last_kelas') != kelas or st.session_state.get('last_agama') != agama:
        st.session_state.nilai_data = {}
        st.session_state.last_kelas = kelas
        st.session_state.last_agama = agama
    
    # Inisialisasi mata pelajaran tambahan (bisa ditambah/hapus oleh user)
    if 'extra_subjects' not in st.session_state:
        st.session_state.extra_subjects = []
    if 'extra_nilai' not in st.session_state:
        st.session_state.extra_nilai = {}
    if 'extra_deskripsi' not in st.session_state:
        st.session_state.extra_deskripsi = {}
    
    nilai_list = []
    deskripsi_list = []
    
    # Input per mata pelajaran
    for i, mapel in enumerate(subjects):
        with st.container(border=True):
            col_nilai, col_desk = st.columns([0.22, 0.78])
            
            with col_nilai:
                default_nilai = st.session_state.nilai_data.get(mapel, 82)
                nilai = st.number_input(
                    f"Nilai Akhir", 
                    min_value=0, max_value=100, value=default_nilai, 
                    key=f"nilai_{i}",
                    help="Masukkan nilai 0-100"
                )
                nilai_list.append(nilai)
                st.session_state.nilai_data[mapel] = nilai
                
                # Predikat otomatis
                if nilai >= 90:
                    predikat = "A (Sangat Baik)"
                elif nilai >= 80:
                    predikat = "B (Baik)"
                elif nilai >= 70:
                    predikat = "C (Cukup)"
                else:
                    predikat = "D (Perlu Bimbingan)"
                st.caption(f"Predikat: **{predikat}**")
            
            with col_desk:
                # Generate atau edit deskripsi
                default_desk = generate_deskripsi_otomatis(nama_siswa, mapel, nilai)
                if mapel not in st.session_state.nilai_data:
                    st.session_state.nilai_data[mapel] = nilai
                
                deskripsi = st.text_area(
                    f"Deskripsi Capaian Kompetensi", 
                    value=default_desk, 
                    height=75, 
                    key=f"desk_{i}",
                    help="Edit sesuai observasi nyata siswa. Bahasa sudah positif & memotivasi."
                )
                deskripsi_list.append(deskripsi)
            
            # Keterangan singkat mata pelajaran (di bawah kolom)
            mapel_key = next((k for k in SUBJECT_KETERANGAN.keys() if k in mapel), None)
            if mapel_key:
                st.caption(f"ℹ️ **{mapel_key}**: {SUBJECT_KETERANGAN[mapel_key]}")
    
    # Tombol regenerate semua deskripsi
    if st.button("🔄 Regenerate Semua Deskripsi Otomatis", type="secondary"):
        for i, mapel in enumerate(subjects):
            new_desk = generate_deskripsi_otomatis(nama_siswa, mapel, nilai_list[i])
            st.session_state[f"desk_{i}"] = new_desk
        st.rerun()
    
    st.info("💡 Deskripsi sudah disesuaikan dengan level nilai dan mata pelajaran. Anda bebas mengedit untuk lebih akurat.")
    
    # ========== FITUR TAMBAH / HAPUS MATA PELAJARAN (BARU) ==========
    with st.expander("➕ Tambah / Hapus Mata Pelajaran Tambahan (Opsional)", expanded=False):
        st.caption("Gunakan fitur ini jika ada mata pelajaran tambahan di luar standar Kurikulum Merdeka untuk kelas tersebut (misalnya Muatan Lokal khusus atau kegiatan unggulan sekolah).")
        
        col_add1, col_add2 = st.columns([0.7, 0.3])
        with col_add1:
            new_mapel_name = st.text_input("Nama Mata Pelajaran Baru", placeholder="Contoh: Bahasa Daerah / Keterampilan", key="new_mapel_input")
        with col_add2:
            if st.button("➕ Tambah Mapel", type="secondary", use_container_width=True):
                if new_mapel_name.strip() and new_mapel_name.strip() not in st.session_state.extra_subjects:
                    st.session_state.extra_subjects.append(new_mapel_name.strip())
                    st.session_state.extra_nilai[new_mapel_name.strip()] = 75
                    st.session_state.extra_deskripsi[new_mapel_name.strip()] = generate_deskripsi_otomatis(nama_siswa, new_mapel_name.strip(), 75)
                    st.rerun()
                elif new_mapel_name.strip() in st.session_state.extra_subjects:
                    st.warning("Mata pelajaran sudah ada!")
        
        if st.session_state.extra_subjects:
            st.markdown("**Mata Pelajaran Tambahan yang Sudah Ditambahkan:**")
            for idx, extra_mapel in enumerate(st.session_state.extra_subjects):
                with st.container(border=True):
                    col_e1, col_e2, col_e3 = st.columns([0.55, 0.25, 0.2])
                    with col_e1:
                        extra_nilai_val = st.number_input(
                            f"Nilai {extra_mapel}", 
                            min_value=0, max_value=100, 
                            value=st.session_state.extra_nilai.get(extra_mapel, 75),
                            key=f"extra_nilai_{idx}"
                        )
                        st.session_state.extra_nilai[extra_mapel] = extra_nilai_val
                        
                        # Predikat otomatis untuk extra
                        if extra_nilai_val >= 90:
                            extra_pred = "A"
                        elif extra_nilai_val >= 80:
                            extra_pred = "B"
                        elif extra_nilai_val >= 70:
                            extra_pred = "C"
                        else:
                            extra_pred = "D"
                        st.caption(f"Predikat: **{extra_pred}**")
                    
                    with col_e2:
                        extra_desk = st.text_area(
                            f"Deskripsi {extra_mapel}", 
                            value=st.session_state.extra_deskripsi.get(extra_mapel, ""),
                            height=60,
                            key=f"extra_desk_{idx}"
                        )
                        st.session_state.extra_deskripsi[extra_mapel] = extra_desk
                    
                    with col_e3:
                        if st.button("🗑️ Hapus", key=f"del_extra_{idx}", type="secondary"):
                            # Hapus dari semua storage
                            if extra_mapel in st.session_state.extra_subjects:
                                st.session_state.extra_subjects.remove(extra_mapel)
                            if extra_mapel in st.session_state.extra_nilai:
                                del st.session_state.extra_nilai[extra_mapel]
                            if extra_mapel in st.session_state.extra_deskripsi:
                                del st.session_state.extra_deskripsi[extra_mapel]
                            st.rerun()
    
    # Gabungkan mata pelajaran tambahan ke daftar utama (untuk PDF & generate)
    final_subjects = subjects.copy()
    for extra_mapel in st.session_state.extra_subjects:
        if extra_mapel not in final_subjects:  # hindari duplikat
            final_subjects.append(extra_mapel)
            nilai_list.append(st.session_state.extra_nilai.get(extra_mapel, 75))
            deskripsi_list.append(st.session_state.extra_deskripsi.get(extra_mapel, ""))
    
    # === CAPAIAN PEMBELAJARAN (CP) & TUJUAN PEMBELAJARAN (TP) - BARU ===
    st.markdown('<div class="section-header">🎯 CAPAIAN PEMBELAJARAN (CP) & TUJUAN PEMBELAJARAN (TP)</div>', unsafe_allow_html=True)
    
    st.caption("Opsional: Ringkasan CP & TP yang menjadi fokus semester ini (bisa diambil dari ATP / Modul Ajar). Akan muncul sebagai bagian terpisah di PDF rapor untuk memperkuat kesesuaian dengan Kurikulum Merdeka.")
    
    cp_tp_ringkasan = st.text_area(
        "Ringkasan Capaian Pembelajaran & Tujuan Pembelajaran yang Dicapai Semester Ini",
        value=f"Pada semester ini peserta didik telah mencapai Capaian Pembelajaran Fase {fase} pada mata pelajaran yang diajarkan. Tujuan Pembelajaran difokuskan pada penguatan literasi, numerasi, karakter, dan Profil Pelajar Pancasila melalui pembelajaran yang bermakna dan berbasis proyek.",
        height=75,
        key="cp_tp_ringkasan",
        help="Anda dapat menyesuaikan dengan CP & TP spesifik dari dokumen ATP sekolah Anda."
    )
    
    # === DATA TAMBAHAN ===
    st.markdown('<div class="section-header">📋 DATA TAMBAHAN</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        sakit = st.number_input("Sakit (hari)", min_value=0, value=0, key="sakit")
    with col2:
        izin = st.number_input("Izin (hari)", min_value=0, value=1, key="izin")
    with col3:
        tanpa_keterangan = st.number_input("Tanpa Keterangan (hari)", min_value=0, value=0, key="tanpa")
    
    catatan_wali = st.text_area(
        "Catatan Wali Kelas", 
        value="Ananda menunjukkan sikap yang baik dan aktif dalam pembelajaran. Pertahankan semangat dan tingkatkan kedisiplinan serta tanggung jawab dalam mengerjakan tugas.",
        height=60,
        key="catatan_wali"
    )
    
    # Tanda tangan
    st.markdown('<div class="section-header">✍️ TANDA TANGAN</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        nama_wali_kelas = st.text_input("Nama Wali Kelas", value="Siti Aminah, S.Pd.", key="wali_kelas")
        nip_wali = st.text_input("NIP Wali Kelas", value="198505152010012003", key="nip_wali")
    with col2:
        nama_kepala = st.text_input("Nama Kepala Sekolah", value="Drs. Budi Santoso, M.Pd.", key="kepala")
        nip_kepala = st.text_input("NIP Kepala Sekolah", value="196812101990011001", key="nip_kepala")
    
    tempat_tanggal = st.text_input("Tempat, Tanggal (contoh: Jakarta, 18 Juni 2026)", 
                                   value=f"{kota}, 18 Juni 2026", key="tempat_tanggal")
    
    # === GENERATE BUTTON ===
    st.markdown("---")
    
    if st.button("📥 GENERATE & DOWNLOAD RAPOR PDF", type="primary", use_container_width=True):
        if not nama_siswa or not nama_sekolah:
            st.error("Nama siswa dan nama sekolah wajib diisi!")
        else:
            with st.spinner("Membuat dokumen PDF rapor..."):
                data = {
                    'nama_sekolah': nama_sekolah,
                    'npsn': npsn,
                    'alamat_sekolah': alamat_sekolah,
                    'kota': kota,
                    'nama_siswa': nama_siswa,
                    'nisn': nisn,
                    'kelas': kelas,
                    'fase': fase,
                    'semester': semester,
                    'tahun_ajaran': tahun_ajaran,
                    'agama': agama,
                    'subjects': final_subjects,
                    'nilai_list': nilai_list,
                    'deskripsi_list': deskripsi_list,
                    'sakit': sakit,
                    'izin': izin,
                    'tanpa_keterangan': tanpa_keterangan,
                    'catatan_wali': catatan_wali,
                    'nama_wali_kelas': nama_wali_kelas,
                    'nip_wali': nip_wali,
                    'nama_kepala': nama_kepala,
                    'nip_kepala': nip_kepala,
                    'tempat_tanggal': tempat_tanggal,
                    'cp_tp_ringkasan': cp_tp_ringkasan,
                    'logo_bytes': logo_file.getvalue() if logo_file else None,
                }
                
                pdf_bytes = create_rapor_pdf(data)
                
                filename = f"Rapor_{nama_siswa.replace(' ', '_')}_Kelas{kelas}_{semester.replace(' ', '')}_{tahun_ajaran.replace('/', '-')}.pdf"
                
                st.success("✅ Rapor berhasil dibuat! Silakan unduh di bawah ini.")
                st.download_button(
                    label="⬇️ UNDUH RAPOR PDF",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True
                )
                
                st.balloons()

# ==================== TAB 2: BATCH EXCEL ====================
with tab2:
    st.header("📦 Batch Processing - Banyak Siswa Sekaligus")
    st.markdown("""
    Fitur ini memungkinkan Anda membuat **puluhan atau ratusan rapor PDF** hanya dengan mengisi satu file Excel.
    Sangat cocok untuk sekolah yang memiliki banyak siswa per kelas.
    """)
    
    st.info("💡 **Alur kerja**: Download template Excel → Isi data & nilai siswa → Upload kembali → Generate ZIP berisi semua PDF")
    
    # Input info sekolah bersama untuk batch
    with st.expander("🏫 Isi Identitas Sekolah (berlaku untuk semua siswa dalam batch ini)", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            batch_nama_sekolah = st.text_input("Nama Sekolah Dasar", value="SD Negeri 01 Contoh", key="batch_nama_sekolah")
            batch_npsn = st.text_input("NPSN", value="12345678", key="batch_npsn")
        with col2:
            batch_alamat = st.text_input("Alamat Sekolah", value="Jl. Pendidikan No. 1", key="batch_alamat")
            batch_kota = st.text_input("Kota/Kabupaten", value="Kota Contoh", key="batch_kota")
        
        # Logo untuk batch (sama untuk semua siswa)
        batch_logo_file = st.file_uploader(
            "🖼️ Upload Logo Sekolah untuk semua rapor dalam batch ini (opsional)", 
            type=["png", "jpg", "jpeg"],
            key="batch_logo",
            help="Logo yang sama akan muncul di setiap PDF rapor yang dihasilkan."
        )
    
    # Tombol download template
    if st.button("📥 DOWNLOAD TEMPLATE EXCEL SIAP PAKAI", type="primary", use_container_width=True):
        template_bytes = create_excel_template()
        st.download_button(
            label="⬇️ Unduh Template_Rapor_Batch.xlsx",
            data=template_bytes,
            file_name="Template_Rapor_Batch_SD_Kelas1-6.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("Template berhasil dibuat! Buka file Excel tersebut dan isi data siswa.")
    
    st.markdown("---")
    
    # Upload Excel
    uploaded_file = st.file_uploader(
        "Upload File Excel yang sudah diisi",
        type=["xlsx", "xls"],
        help="File harus mengikuti format template di atas"
    )
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name="Data_Siswa")
            df = df.dropna(how='all')  # hapus baris kosong
            
            st.success(f"✅ File berhasil dibaca! Ditemukan **{len(df)} baris data**.")
            
            # Preview
            with st.expander("👀 Lihat Preview Data (10 baris pertama)", expanded=True):
                st.dataframe(df.head(10), use_container_width=True)
            
            # Tombol generate
            if st.button("🚀 GENERATE SEMUA PDF & DOWNLOAD ZIP", type="primary", use_container_width=True):
                with st.spinner(f"Memproses {len(df)} siswa... Mohon tunggu beberapa saat"):
                    zip_bytes, success_count, error_list = process_batch_excel(
                        df,
                        nama_sekolah=batch_nama_sekolah,
                        npsn=batch_npsn,
                        alamat_sekolah=batch_alamat,
                        kota=batch_kota,
                        logo_bytes=batch_logo_file.getvalue() if batch_logo_file else None
                    )
                
                st.success(f"🎉 Selesai! **{success_count}** rapor berhasil dibuat.")
                
                if error_list:
                    st.warning(f"Ada {len(error_list)} error. Lihat detail di file RINGKASAN_GENERATE.txt di dalam ZIP.")
                
                # Download ZIP
                zip_filename = f"Rapor_Batch_{datetime.now().strftime('%Y%m%d_%H%M')}.zip"
                st.download_button(
                    label=f"⬇️ UNDUH {success_count} FILE PDF (ZIP)",
                    data=zip_bytes,
                    file_name=zip_filename,
                    mime="application/zip",
                    use_container_width=True
                )
                
                # Tampilkan ringkasan error jika ada
                if error_list:
                    with st.expander("Lihat Daftar Error"):
                        for err in error_list:
                            st.error(err)
        
        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file Excel: {str(e)}")
            st.info("Pastikan file yang diupload adalah template yang benar dan sheet bernama 'Data_Siswa'.")

# ==================== TAB 3: PANDUAN ====================
with tab3:
    st.header("📖 Panduan Penggunaan & Catatan Penting")
    
    st.markdown("""
    ### Fitur Utama Aplikasi
    - **Input mudah** untuk semua mata pelajaran SD (Kelas 1-6)
    - **Deskripsi otomatis** yang positif, memotivasi, dan sesuai bahasa Kurikulum Merdeka
    - **Predikat otomatis** (A/B/C/D) berdasarkan nilai
    - **PDF profesional** siap cetak dengan layout rapi
    - Mendukung **Fase A (Kelas 1-2)**, **Fase B (Kelas 3-4)**, **Fase C (Kelas 5-6)**
    
    ### Mata Pelajaran yang Didukung
    | Kelas | Mata Pelajaran Utama |
    |-------|----------------------|
    | 1-2   | Agama & Budi Pekerti, Pendidikan Pancasila, Bahasa Indonesia, Matematika, PJOK, Seni Budaya, Muatan Lokal |
    | 3-4   | + IPAS |
    | 5-6   | + Bahasa Inggris |
    
    ### Tips Pengisian yang Akurat
    1. Nilai Akhir adalah nilai akhir/holistik (bisa rata-rata Pengetahuan + Keterampilan atau penilaian menyeluruh)
    2. Edit deskripsi sesuai observasi nyata siswa (deskripsi otomatis hanya panduan)
    3. Gunakan bahasa positif dan memotivasi seperti contoh resmi Kemendikdasmen
    4. Untuk siswa dengan nilai rendah, deskripsi tetap memberikan harapan dan arahan perbaikan
    
    ### Format PDF
    - Ukuran A4 Portrait
    - Layout mengikuti contoh resmi Panduan Pembelajaran dan Asesmen
    - Siap untuk dicetak atau diarsipkan secara digital
    - Tanda tangan digital-ready (bisa ditambah tanda tangan scan jika diedit lebih lanjut)
    
    ### Batasan & Saran
    - Aplikasi ini untuk **satu siswa per generate** (ulangi untuk siswa lain)
    - Untuk sekolah besar, Anda bisa membuat script batch tambahan atau gunakan Excel sebagai sumber data
    - Logo sekolah belum disertakan di PDF versi ini (bisa ditambahkan di pengembangan selanjutnya)
    - Simpan file PDF dengan nama yang jelas untuk arsip
    
    ### Legal & Akurasi
    Format ini mengacu pada:
    - Panduan Pembelajaran dan Asesmen Pendidikan Dasar (Kemendikdasmen)
    - Contoh Format Laporan Hasil Belajar (Rapor) Jenjang SD Fase A, B, dan C
    - Struktur Kurikulum SD Permendikdasmen terbaru
    
    Anda tetap bertanggung jawab atas keakuratan data dan deskripsi yang diinput.
    """)
    
    st.success("Aplikasi ini dibuat untuk memudahkan guru SD dalam menyusun rapor yang akurat, rapi, dan sesuai standar Kurikulum Merdeka.")

# Footer
st.markdown("---")
st.caption("© 2026 Aplikasi Rapor SD | Dibuat dengan ❤️ untuk kemudahan administrasi guru Indonesia")