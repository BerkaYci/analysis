# -*- coding: utf-8 -*-
"""
Kesinti Analiz - Konfigürasyon Dosyası
Tüm sütun indeksleri ve ayarlar burada tanımlanır.
"""

# ============================================================================
# KAYNAK DOSYA SÜTUN İNDEKSLERİ (0-tabanlı)
# ============================================================================

# Ana kesinti Excel dosyası (table.xlsx / kesinti dosyası)
KESINTI_SUTUNLARI = {
    'INOUT': 'Tablo-1 IN-OUT FLG',
    'KESINTI_NO': 'Kesinti No',
    'KADEME': 'Kademe',
    'SEBEKE_UNSURU': 'Şebeke Unsuru',
    'BASLAMA': 'Kesinti Başlama Zamanı',
    'BITIS': 'Kesinti/Kademe Bitiş Zamanı',
    'SCADA': 'Scada Kesintisi ',
    'SON_CAGRI': 'Son Çağrı Zamanı',
    'ILK_MUSTERI_DISI': 'İlk Müşteri Dışı Çağrı Zamanı',
    'ILK_MUSTERI': 'İlk Müşteri Çağrı Zamanı',
    'CBS_TM_NO': 'CBS TM No',          # BA sütunu - TM No (Dağıtım-AG için)
    'KAYNAGA_GORE': 'Kaynağa Göre',    # Dağıtım-AG kontrolü için
    'TOPLAM_CAGRI': 'Toplam Çağrı Sayısı',
    'KESINTI_SEVIYESI': 'Kesinti Seviyesi'
}

# Dağıtım-AG ayarları
DAGITIM_AG_AYARLARI = {
    'KAYNAGA_GORE_DEGER': 'Dağıtım-AG',  # "Kaynağa Göre" sütunundaki değer
    'TARAMA_SAAT': 12,                    # TM bazlı tarama için saat (±12 saat)
    'JTK_OLUSTUR': False                  # Dağıtım-AG için JTK oluşturulsun mu?
}

# TM No Ard Arda ayarları (Dağıtım-AG için)
TM_ARDARDA_AYARLARI = {
    'KRITIK_SAAT': 9,           # Kritik süre (saat) - bu değere göre tolerans belirlenir
    'TOLERANS_USTU_DK': 60,     # Kritik süre üstü için tolerans (dakika)
    'TOLERANS_ALTI_DK': 15      # Kritik süre altı için tolerans (dakika)
}

# CM.xlsx sütun indeksleri
CM_SUTUN_INDEKSLERI = {
    'HIZMET_NO': 2,           # C sütunu - Hizmet No
    'OMS_TICKET_ID': 22,      # W sütunu - OMS Ticket ID
    'KESINTI_ID': 27,         # AB sütunu - Kesinti ID (arama sütunu)
    'OLUSTURMA_TARIHI': 29    # AD sütunu - Oluşturma Tarihi
}

# table.xlsx sütun indeksleri (OTG için)
TABLE_SUTUN_INDEKSLERI = {
    'KESINTI_ID': 3,          # D sütunu - Kesinti No
    'SECILEN_SUTUNLAR': [0, 2, 3, 4, 6, 7, 10, 12, 13, 14, 15, 16, 17, 18],
    'ETKILENEN_KULLANICI_T': 19,
    'ETKILENEN_KULLANICI_U': 20,
    'ETKILENEN_KULLANICI_V': 21,
    'ETKILENEN_KULLANICI_W': 22,
    'EK_SUTUNLAR': [29, 30, 41, 48, 49, 50],
    'OMS_YORUM': 44,
    'ILK_CAGRI_AT': 45,
    'ILK_CAGRI_AU': 46,
    'SON_SUTUN': 59
}

# jtk.xlsx sütun indeksleri
JTK_SUTUN_INDEKSLERI = {
    'KESINTI_ID': 2           # C sütunu - Kesinti ID
}

# veri.xlsx (Birlesik_Analiz) sütun indeksleri
VERI_SUTUN_INDEKSLERI = {
    'ILGILI_KESINTILER': 7,   # H sütunu - İlgili Kesintiler (;)
    'ORTAK_W_DEGERLERI': 9    # J sütunu - Ortak W Değerleri (eski: K sütunu)
}

# ============================================================================
# EXCEL OKUMA AYARLARI
# ============================================================================

EXCEL_AYARLARI = {
    'KESINTI_HEADER_ROW': 3,     # Kesinti dosyası başlık satırı (0-tabanlı: 4. satır)
    'CM_HEADER_ROW': None,       # CM dosyası başlık yok (header=None)
    'TABLE_SKIP_ROWS': 3,        # table.xlsx atlanacak satır sayısı
    'JTK_HEADER_ROW': 0,         # jtk.xlsx başlık satırı
    'VERI_HEADER_ROW': 0         # veri.xlsx başlık satırı
}

# ============================================================================
# GÖRSEL AYARLAR
# ============================================================================

PNG_AYARLARI = {
    'OTG_FIG_WIDTH': 40,          # OTG için çok geniş (büyük yazı + çok sütun)
    'JTK_FIG_WIDTH': 26,          
    'DPI': 250,                    # Biraz düşük DPI (dosya boyutu için)
    'OTG_FONT_SIZE': 9,            # BÜYÜK YAZI
    'JTK_FONT_SIZE': 10,           # BÜYÜK YAZI
    'HEADER_FONT_SIZE': 10,        # Başlık için BÜYÜK
    'HEADER_COLOR': '#4472C4',
    'ALTERNATE_ROW_COLOR': '#E7E6E6',
    'CELL_PAD': 0.025,             # Hücre padding (artırıldı)
    'CELL_HEIGHT': 0.025,          # Hücre yüksekliği (artırıldı)
    'MAX_CELL_WIDTH': 25,          # Maksimum hücre genişliği
    'MAX_TEXT_LENGTH': 50,         # Maksimum metin uzunluğu
    'MIN_COL_WIDTH': 4,            # Minimum sütun genişliği
    'HEADER_WRAP_LENGTH': 12,      # Başlık wrap (kısa - alta geçsin)
    'CELL_WRAP_LENGTH': 15,        # Veri hücresi wrap (kısa - alta geçsin)
    'OTG_CELL_WRAP': 12,           # OTG için özel wrap (çok kısa - alta geçsin)
    'ROW_HEIGHT_FACTOR': 0.5       # Satır yükseklik çarpanı (artırıldı)
}

# ============================================================================
# EXCEL STIL AYARLARI
# ============================================================================

EXCEL_STIL = {
    'HEADER_BG_COLOR': '4472C4',
    'HEADER_FONT_COLOR': 'FFFFFF',
    'ALTERNATE_ROW_COLOR': 'E7E6E6',
    'MAX_COLUMN_WIDTH': 60
}

# ============================================================================
# VARSAYILAN DEĞERLER
# ============================================================================

VARSAYILAN = {
    'TOLERANS_DAKIKA': 30,
    'OUTPUT_FOLDER': 'outputs',
    'ANALIZ_DOSYA_ADI': 'Birlesik_Analiz.xlsx'
}

