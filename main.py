# -*- coding: utf-8 -*-
"""
Kesinti Analiz - Birleşik Panel
Kesinti analizi ve dosyalama işlemlerini tek arayüzden yönetir.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys

# Modülleri import et
from config import VARSAYILAN, VERI_SUTUN_INDEKSLERI, TM_ARDARDA_AYARLARI
from modules.kesinti_analiz import KesintiAnaliz
from modules.dosyalama import Dosyalama
from modules.excel_yardimci import ExcelYardimci


class ModernButton(tk.Canvas):
    """Yuvarlak köşeli modern buton"""
    
    def __init__(self, parent, text, command, bg_color, hover_color, fg='white', 
                 width=200, height=45, corner_radius=10, font_size=11, bold=True):
        super().__init__(parent, width=width, height=height, 
                        bg=parent.cget('bg'), highlightthickness=0)
        
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.fg = fg
        self.corner_radius = corner_radius
        self.btn_width = width
        self.btn_height = height
        self.text = text
        self.font_weight = 'bold' if bold else 'normal'
        self.font_size = font_size
        
        self._draw_button(bg_color)
        
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<Button-1>', self._on_click)
    
    def _draw_button(self, color):
        self.delete('all')
        r = self.corner_radius
        w, h = self.btn_width, self.btn_height
        
        # Gölge (koyu gri renk)
        shadow_color = '#1a1a1a'
        self.create_oval(3, 3, r*2+3, r*2+3, fill=shadow_color, outline='')
        self.create_oval(w-r*2+3, 3, w+3, r*2+3, fill=shadow_color, outline='')
        self.create_oval(3, h-r*2+3, r*2+3, h+3, fill=shadow_color, outline='')
        self.create_oval(w-r*2+3, h-r*2+3, w+3, h+3, fill=shadow_color, outline='')
        self.create_rectangle(r+3, 3, w-r+3, h+3, fill=shadow_color, outline='')
        self.create_rectangle(3, r+3, w+3, h-r+3, fill=shadow_color, outline='')
        
        # Ana buton şekli
        self.create_oval(0, 0, r*2, r*2, fill=color, outline='')
        self.create_oval(w-r*2, 0, w, r*2, fill=color, outline='')
        self.create_oval(0, h-r*2, r*2, h, fill=color, outline='')
        self.create_oval(w-r*2, h-r*2, w, h, fill=color, outline='')
        self.create_rectangle(r, 0, w-r, h, fill=color, outline='')
        self.create_rectangle(0, r, w, h-r, fill=color, outline='')
        
        # Metin
        self.create_text(w/2, h/2, text=self.text, fill=self.fg, 
                        font=('Segoe UI', self.font_size, self.font_weight))
    
    def _on_enter(self, event):
        self._draw_button(self.hover_color)
        self.config(cursor='hand2')
    
    def _on_leave(self, event):
        self._draw_button(self.bg_color)
    
    def _on_click(self, event):
        if self.command:
            self.command()
    
    def config(self, **kwargs):
        if 'state' in kwargs:
            if kwargs['state'] == 'disabled':
                self._draw_button('#cccccc')
                self.unbind('<Button-1>')
            else:
                self._draw_button(self.bg_color)
                self.bind('<Button-1>', self._on_click)
        super().config(**kwargs)


class BirlesikPanel:
    """Birleşik kesinti analiz ve dosyalama paneli"""
    
    # Renk paleti
    COLORS = {
        'bg_dark': '#1a1a2e',
        'bg_medium': '#16213e',
        'bg_light': '#0f3460',
        'accent_blue': '#4361ee',
        'accent_purple': '#7209b7',
        'accent_pink': '#f72585',
        'accent_green': '#06d6a0',
        'accent_orange': '#fb8500',
        'accent_cyan': '#00b4d8',
        'text_light': '#ffffff',
        'text_muted': '#a0a0a0',
        'card_bg': '#1e2746',
        'input_bg': '#2a3f5f',
        'success': '#10b981',
        'warning': '#f59e0b',
        'error': '#ef4444'
    }
    
    def __init__(self, root):
        self.root = root
        self.root.title("⚡ Kesinti Analiz Pro")
        self.root.configure(bg=self.COLORS['bg_dark'])
        self.root.resizable(True, True)
        
        # Pencere boyutunu ekrana göre ayarla
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = 780
        window_height = min(720, int(screen_height * 0.85))
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2 - 30
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Minimum boyut
        self.root.minsize(750, 650)
        
        # Değişkenler
        self.kesinti_dosyasi = None
        self.analiz_sonuc_yolu = None
        self.cikti_klasoru = None
        self.analiz_engine = None
        self.dosyalama_engine = None
        self.grup_list = []
        
        # ttk stil ayarları
        self._setup_styles()
        self._create_ui()
    
    def _setup_styles(self):
        """ttk stil ayarları"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Progress bar stili
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=self.COLORS['bg_medium'],
            background=self.COLORS['accent_cyan'],
            thickness=8,
            borderwidth=0
        )
        
        # Entry stili
        style.configure(
            "Custom.TEntry",
            fieldbackground=self.COLORS['input_bg'],
            foreground=self.COLORS['text_light'],
            insertcolor=self.COLORS['text_light']
        )
    
    def _create_header(self, parent):
        """Başlık bölümü"""
        header_frame = tk.Frame(parent, bg=self.COLORS['bg_dark'])
        header_frame.pack(fill='x', pady=(0, 10))
        
        # Logo ve başlık
        title_frame = tk.Frame(header_frame, bg=self.COLORS['bg_dark'])
        title_frame.pack()
        
        # Elektrik ikonu
        icon_canvas = tk.Canvas(title_frame, width=40, height=40, 
                               bg=self.COLORS['bg_dark'], highlightthickness=0)
        icon_canvas.pack(side='left', padx=(0, 10))
        
        # Şimşek ikonu çiz
        points = [20, 4, 8, 20, 16, 20, 12, 36, 24, 18, 16, 18]
        icon_canvas.create_polygon(points, fill=self.COLORS['accent_orange'], 
                                   outline=self.COLORS['accent_pink'], width=2)
        
        title_text = tk.Frame(title_frame, bg=self.COLORS['bg_dark'])
        title_text.pack(side='left')
        
        tk.Label(
            title_text,
            text="KESİNTİ ANALİZ",
            font=('Segoe UI', 22, 'bold'),
            bg=self.COLORS['bg_dark'],
            fg=self.COLORS['text_light']
        ).pack(anchor='w')
        
        tk.Label(
            title_text,
            text="Profesyonel Analiz & Raporlama Sistemi",
            font=('Segoe UI', 9),
            bg=self.COLORS['bg_dark'],
            fg=self.COLORS['text_muted']
        ).pack(anchor='w')
    
    def _create_card(self, parent, title, icon, accent_color):
        """Kart bileşeni oluştur"""
        card = tk.Frame(parent, bg=self.COLORS['card_bg'], padx=15, pady=10)
        card.pack(fill='x', pady=5)
        
        # Başlık satırı
        header = tk.Frame(card, bg=self.COLORS['card_bg'])
        header.pack(fill='x', pady=(0, 8))
        
        # Accent çizgisi
        accent_bar = tk.Frame(header, bg=accent_color, width=4, height=20)
        accent_bar.pack(side='left', padx=(0, 10))
        
        tk.Label(
            header,
            text=f"{icon} {title}",
            font=('Segoe UI', 11, 'bold'),
            bg=self.COLORS['card_bg'],
            fg=self.COLORS['text_light']
        ).pack(side='left')
        
        return card
    
    def _create_input_row(self, parent, label_text, entry_var_name, button_text, 
                          button_command, button_color, entry_width=45):
        """Input satırı oluştur"""
        row = tk.Frame(parent, bg=self.COLORS['card_bg'])
        row.pack(fill='x', pady=4)
        
        tk.Label(
            row,
            text=label_text,
            font=('Segoe UI', 9),
            bg=self.COLORS['card_bg'],
            fg=self.COLORS['text_muted'],
            width=12,
            anchor='w'
        ).pack(side='left')
        
        entry = tk.Entry(
            row,
            width=entry_width,
            font=('Segoe UI', 9),
            bg=self.COLORS['input_bg'],
            fg=self.COLORS['text_light'],
            insertbackground=self.COLORS['text_light'],
            relief='flat',
            highlightthickness=1,
            highlightcolor=self.COLORS['accent_blue'],
            highlightbackground=self.COLORS['bg_light']
        )
        entry.pack(side='left', padx=6, ipady=4)
        setattr(self, entry_var_name, entry)
        
        btn = tk.Button(
            row,
            text=button_text,
            command=button_command,
            font=('Segoe UI', 8, 'bold'),
            bg=button_color,
            fg='white',
            relief='flat',
            cursor='hand2',
            padx=10,
            pady=3,
            activebackground=self.COLORS['accent_purple'],
            activeforeground='white'
        )
        btn.pack(side='left')
        
        return entry
    
    def _create_ui(self):
        """Arayüzü oluştur"""
        # Ana kaydırılabilir alan
        main_canvas = tk.Canvas(self.root, bg=self.COLORS['bg_dark'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient='vertical', command=main_canvas.yview)
        
        main_frame = tk.Frame(main_canvas, bg=self.COLORS['bg_dark'], padx=20, pady=12)
        
        main_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        main_canvas.pack(side='left', fill='both', expand=True)
        
        canvas_frame = main_canvas.create_window((0, 0), window=main_frame, anchor='nw')
        
        def configure_scroll(event):
            main_canvas.configure(scrollregion=main_canvas.bbox('all'))
            main_canvas.itemconfig(canvas_frame, width=event.width)
        
        main_frame.bind('<Configure>', lambda e: main_canvas.configure(scrollregion=main_canvas.bbox('all')))
        main_canvas.bind('<Configure>', configure_scroll)
        
        # Mouse wheel scroll
        def on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        main_canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # ═══════════════════════════════════════════════════════════════
        # BAŞLIK
        # ═══════════════════════════════════════════════════════════════
        self._create_header(main_frame)
        
        # ═══════════════════════════════════════════════════════════════
        # BÖLÜM 1: KESİNTİ ANALİZİ
        # ═══════════════════════════════════════════════════════════════
        analiz_card = self._create_card(
            main_frame, 
            "ADIM 1: Kesinti Analizi", 
            "📊", 
            self.COLORS['accent_blue']
        )
        
        # Kesinti dosyası
        self._create_input_row(
            analiz_card, "Kesinti Excel:", "entry_kesinti",
            "📁 Dosya Seç", self._kesinti_dosyasi_sec,
            self.COLORS['accent_blue']
        )
        
        # Tolerans Ayarları Kartı
        tolerans_inner = tk.Frame(analiz_card, bg=self.COLORS['bg_medium'], padx=12, pady=8)
        tolerans_inner.pack(fill='x', pady=6)
        
        tk.Label(
            tolerans_inner,
            text="⚙️ Ard Arda Tolerans Ayarları",
            font=('Segoe UI', 9, 'bold'),
            bg=self.COLORS['bg_medium'],
            fg=self.COLORS['accent_cyan']
        ).pack(anchor='w', pady=(0, 6))
        
        tolerans_row = tk.Frame(tolerans_inner, bg=self.COLORS['bg_medium'])
        tolerans_row.pack(fill='x')
        
        # Kritik Süre
        self._create_param_input(tolerans_row, "Kritik Süre (saat)", "entry_kritik_saat", 
                                TM_ARDARDA_AYARLARI['KRITIK_SAAT'])
        
        # ≥ Tolerans
        self._create_param_input(tolerans_row, "≥ Tolerans (dk)", "entry_tolerans_ustu",
                                TM_ARDARDA_AYARLARI['TOLERANS_USTU_DK'])
        
        # < Tolerans
        self._create_param_input(tolerans_row, "< Tolerans (dk)", "entry_tolerans_alti",
                                TM_ARDARDA_AYARLARI['TOLERANS_ALTI_DK'])
        
        # Analiz butonu
        btn_frame = tk.Frame(analiz_card, bg=self.COLORS['card_bg'])
        btn_frame.pack(fill='x', pady=(10, 3))
        
        self.btn_analiz = ModernButton(
            btn_frame, "▶  ANALİZİ BAŞLAT", self._analiz_baslat,
            self.COLORS['accent_green'], '#0ea472',
            width=180, height=38, font_size=10
        )
        self.btn_analiz.pack()
        
        # Sonuç etiketi
        self.lbl_analiz_sonuc = tk.Label(
            analiz_card,
            text="",
            font=('Segoe UI', 9),
            bg=self.COLORS['card_bg'],
            fg=self.COLORS['success']
        )
        self.lbl_analiz_sonuc.pack(pady=3)
        
        # ═══════════════════════════════════════════════════════════════
        # BÖLÜM 2: DOSYALAMA / RAPORLAMA
        # ═══════════════════════════════════════════════════════════════
        dosyalama_card = self._create_card(
            main_frame,
            "ADIM 2: Raporlama (PNG/Excel)",
            "📁",
            self.COLORS['accent_purple']
        )
        
        # Hazır analiz
        self._create_input_row(
            dosyalama_card, "Hazır Analiz:", "entry_hazir_analiz",
            "📄 Analiz Seç", self._hazir_analiz_sec,
            self.COLORS['accent_cyan']
        )
        
        # Çıktı klasörü
        self._create_input_row(
            dosyalama_card, "Çıktı Klasörü:", "entry_cikti",
            "📁 Klasör Seç", self._cikti_klasoru_sec,
            self.COLORS['accent_purple']
        )
        
        # Bilgi etiketi
        info_frame = tk.Frame(dosyalama_card, bg=self.COLORS['bg_medium'], padx=10, pady=5)
        info_frame.pack(fill='x', pady=4)
        
        tk.Label(
            info_frame,
            text="💡 Klasörde bulunması gereken: table.xlsx, jtk.xlsx, cm.xlsx",
            font=('Segoe UI', 8),
            bg=self.COLORS['bg_medium'],
            fg=self.COLORS['text_muted']
        ).pack(anchor='w')
        
        # Grup listesi
        grup_frame = tk.Frame(dosyalama_card, bg=self.COLORS['bg_medium'], padx=8, pady=6)
        grup_frame.pack(fill='both', expand=True, pady=4)
        
        tk.Label(
            grup_frame,
            text="📋 Bulunan Kesinti Grupları",
            font=('Segoe UI', 9, 'bold'),
            bg=self.COLORS['bg_medium'],
            fg=self.COLORS['text_light']
        ).pack(anchor='w', pady=(0, 5))
        
        # Listbox container
        list_container = tk.Frame(grup_frame, bg=self.COLORS['input_bg'])
        list_container.pack(fill='both', expand=True)
        
        list_scroll = tk.Scrollbar(list_container)
        list_scroll.pack(side='right', fill='y')
        
        self.grup_listbox = tk.Listbox(
            list_container,
            height=6,
            font=('Consolas', 9),
            yscrollcommand=list_scroll.set,
            bg=self.COLORS['input_bg'],
            fg=self.COLORS['text_light'],
            selectbackground=self.COLORS['accent_purple'],
            selectforeground='white',
            relief='flat',
            highlightthickness=0,
            activestyle='none'
        )
        self.grup_listbox.pack(fill='both', expand=True, padx=2, pady=2)
        list_scroll.config(command=self.grup_listbox.yview)
        
        # Raporlama butonları
        btn_frame2 = tk.Frame(dosyalama_card, bg=self.COLORS['card_bg'])
        btn_frame2.pack(fill='x', pady=(10, 3))
        
        self.btn_gruplar = ModernButton(
            btn_frame2, "🔄 Grupları Yükle", self._gruplari_yukle,
            self.COLORS['accent_orange'], '#e07800',
            width=140, height=35, font_size=9
        )
        self.btn_gruplar.pack(side='left', padx=(0, 12))
        
        self.btn_raporla = ModernButton(
            btn_frame2, "📊 RAPORLARI OLUŞTUR", self._raporlari_olustur,
            self.COLORS['accent_purple'], '#5c078f',
            width=180, height=38, font_size=10
        )
        self.btn_raporla.pack(side='left')
        
        # ═══════════════════════════════════════════════════════════════
        # DURUM ÇUBUĞU
        # ═══════════════════════════════════════════════════════════════
        status_card = tk.Frame(main_frame, bg=self.COLORS['bg_medium'], padx=15, pady=8)
        status_card.pack(fill='x', pady=(10, 0))
        
        status_header = tk.Frame(status_card, bg=self.COLORS['bg_medium'])
        status_header.pack(fill='x')
        
        self.status_indicator = tk.Canvas(
            status_header, width=12, height=12,
            bg=self.COLORS['bg_medium'], highlightthickness=0
        )
        self.status_indicator.pack(side='left', padx=(0, 10))
        self.status_indicator.create_oval(2, 2, 10, 10, fill=self.COLORS['success'], outline='')
        
        self.lbl_status = tk.Label(
            status_header,
            text="✓ Sistem hazır",
            font=('Segoe UI', 10),
            bg=self.COLORS['bg_medium'],
            fg=self.COLORS['text_light'],
            anchor='w'
        )
        self.lbl_status.pack(side='left', fill='x', expand=True)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            status_card,
            style="Custom.Horizontal.TProgressbar",
            mode='determinate',
            length=300
        )
        self.progress.pack(fill='x', pady=(6, 0))
    
    def _create_param_input(self, parent, label, var_name, default_value):
        """Parametre input bileşeni"""
        frame = tk.Frame(parent, bg=self.COLORS['bg_medium'])
        frame.pack(side='left', padx=(0, 25))
        
        tk.Label(
            frame,
            text=label,
            font=('Segoe UI', 9),
            bg=self.COLORS['bg_medium'],
            fg=self.COLORS['text_muted']
        ).pack(anchor='w')
        
        entry = tk.Entry(
            frame,
            width=10,
            font=('Segoe UI', 11),
            bg=self.COLORS['input_bg'],
            fg=self.COLORS['text_light'],
            insertbackground=self.COLORS['text_light'],
            relief='flat',
            justify='center',
            highlightthickness=1,
            highlightcolor=self.COLORS['accent_cyan'],
            highlightbackground=self.COLORS['bg_light']
        )
        entry.pack(ipady=5)
        entry.insert(0, str(default_value))
        setattr(self, var_name, entry)
    
    def _set_status(self, text, status_type='info'):
        """Durum çubuğunu güncelle"""
        colors = {
            'info': self.COLORS['accent_blue'],
            'success': self.COLORS['success'],
            'warning': self.COLORS['warning'],
            'error': self.COLORS['error'],
            'processing': self.COLORS['accent_cyan']
        }
        color = colors.get(status_type, self.COLORS['text_light'])
        
        self.lbl_status.config(text=text, fg=color)
        self.status_indicator.delete('all')
        self.status_indicator.create_oval(2, 2, 10, 10, fill=color, outline='')
        self.root.update()
    
    def _kesinti_dosyasi_sec(self):
        """Kesinti Excel dosyasını seç"""
        dosya = filedialog.askopenfilename(
            title="Kesinti Excel Dosyasını Seç",
            filetypes=[("Excel Dosyası", "*.xlsx *.xls")]
        )
        if dosya:
            self.kesinti_dosyasi = dosya
            self.entry_kesinti.delete(0, tk.END)
            self.entry_kesinti.insert(0, dosya)
            
            klasor = os.path.dirname(dosya)
            self.entry_cikti.delete(0, tk.END)
            self.entry_cikti.insert(0, klasor)
            self.cikti_klasoru = klasor
            
            self._set_status(f"✓ Dosya seçildi: {os.path.basename(dosya)}", 'success')
    
    def _cikti_klasoru_sec(self):
        """Çıktı klasörünü seç"""
        klasor = filedialog.askdirectory(title="Çıktı Klasörünü Seç")
        if klasor:
            self.cikti_klasoru = klasor
            self.entry_cikti.delete(0, tk.END)
            self.entry_cikti.insert(0, klasor)
            self._set_status(f"✓ Klasör seçildi: {os.path.basename(klasor)}", 'success')
    
    def _hazir_analiz_sec(self):
        """Hazır analiz dosyasını seç"""
        dosya = filedialog.askopenfilename(
            title="Hazır Analiz Dosyasını Seç",
            filetypes=[("Excel Dosyası", "*.xlsx *.xls")]
        )
        if dosya:
            self.analiz_sonuc_yolu = dosya
            self.entry_hazir_analiz.delete(0, tk.END)
            self.entry_hazir_analiz.insert(0, dosya)
            
            klasor = os.path.dirname(dosya)
            self.entry_cikti.delete(0, tk.END)
            self.entry_cikti.insert(0, klasor)
            self.cikti_klasoru = klasor
            
            self._set_status(f"✓ Hazır analiz seçildi: {os.path.basename(dosya)}", 'success')
            self._gruplari_yukle()
    
    def _analiz_baslat(self):
        """Kesinti analizini başlat"""
        dosya = self.entry_kesinti.get().strip()
        kritik_saat_text = self.entry_kritik_saat.get().strip()
        tolerans_ustu_text = self.entry_tolerans_ustu.get().strip()
        tolerans_alti_text = self.entry_tolerans_alti.get().strip()
        
        if not os.path.exists(dosya):
            messagebox.showerror("Hata", "Lütfen geçerli bir Excel dosyası seçin.")
            return
        
        if not kritik_saat_text.isdigit():
            messagebox.showerror("Hata", "Kritik süre saat cinsinden sayısal olmalı.")
            return
        
        if not tolerans_ustu_text.isdigit() or not tolerans_alti_text.isdigit():
            messagebox.showerror("Hata", "Tolerans değerleri dakika cinsinden sayısal olmalı.")
            return
        
        tolerans_ayarlari = {
            'kritik_saat': int(kritik_saat_text),
            'tolerans_ustu_dk': int(tolerans_ustu_text),
            'tolerans_alti_dk': int(tolerans_alti_text)
        }
        
        self._set_status("⏳ Analiz yapılıyor, lütfen bekleyin...", 'processing')
        self.btn_analiz.config(state='disabled')
        self.progress['value'] = 20
        self.root.update()
        
        try:
            self.analiz_engine = KesintiAnaliz()
            df_sonuc = self.analiz_engine.analiz_yap(dosya, tolerans_ayarlari)
            
            self.progress['value'] = 60
            self.root.update()
            
            if df_sonuc.empty:
                messagebox.showinfo("Sonuç Yok", "Belirtilen kriterlere göre ard arda veya iç içe kesinti bulunamadı.")
                self._set_status("⚠️ Sonuç bulunamadı", 'warning')
                self.btn_analiz.config(state='normal')
                self.progress['value'] = 0
                return
            
            self.analiz_sonuc_yolu = os.path.join(os.path.dirname(dosya), VARSAYILAN['ANALIZ_DOSYA_ADI'])
            self.analiz_engine.kaydet(self.analiz_sonuc_yolu)
            
            self.progress['value'] = 100
            self.root.update()
            
            self.lbl_analiz_sonuc.config(
                text=f"✓ {len(df_sonuc)} grup bulundu → {VARSAYILAN['ANALIZ_DOSYA_ADI']}",
                fg=self.COLORS['success']
            )
            self._set_status(f"✓ Analiz tamamlandı! {len(df_sonuc)} grup bulundu", 'success')
            
            self._gruplari_yukle()
            
            messagebox.showinfo(
                "Analiz Tamamlandı",
                f"✅ Analiz başarıyla tamamlandı!\n\n"
                f"📊 Bulunan grup sayısı: {len(df_sonuc)}\n"
                f"📁 Kaydedilen dosya: {self.analiz_sonuc_yolu}\n\n"
                f"Şimdi 'Raporları Oluştur' butonuna tıklayarak PNG/Excel raporlarını oluşturabilirsiniz."
            )
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Hata", f"Analiz sırasında hata oluştu:\n\n{str(e)}")
            self._set_status(f"✗ Hata: {str(e)[:50]}...", 'error')
        
        finally:
            self.btn_analiz.config(state='normal')
            self.progress['value'] = 0
    
    def _gruplari_yukle(self):
        """Analiz sonucundan grupları yükle"""
        self.grup_listbox.delete(0, tk.END)
        self.grup_list = []
        
        if self.analiz_sonuc_yolu and os.path.exists(self.analiz_sonuc_yolu):
            try:
                self.dosyalama_engine = Dosyalama(self.cikti_klasoru or os.path.dirname(self.analiz_sonuc_yolu))
                self.dosyalama_engine.analiz_sonucunu_yukle(self.analiz_sonuc_yolu)
                self.grup_list = self.dosyalama_engine.gruplari_yukle()
                
                for idx, grup in enumerate(self.grup_list, 1):
                    id_sayisi = len(grup.split(';'))
                    self.grup_listbox.insert(tk.END, f"  {idx:02d}. {grup} ({id_sayisi} ID)")
                
                self._set_status(f"✓ {len(self.grup_list)} grup yüklendi", 'success')
                return
            except Exception as e:
                print(f"Analiz sonucundan yükleme hatası: {e}")
        
        if self.cikti_klasoru:
            veri_path = os.path.join(self.cikti_klasoru, 'veri.xlsx')
            if os.path.exists(veri_path):
                try:
                    self.dosyalama_engine = Dosyalama(self.cikti_klasoru)
                    self.grup_list = self.dosyalama_engine.gruplari_yukle(veri_path)
                    
                    for idx, grup in enumerate(self.grup_list, 1):
                        id_sayisi = len(grup.split(';'))
                        self.grup_listbox.insert(tk.END, f"  {idx:02d}. {grup} ({id_sayisi} ID)")
                    
                    self._set_status(f"✓ veri.xlsx'den {len(self.grup_list)} grup yüklendi", 'success')
                    return
                except Exception as e:
                    print(f"veri.xlsx'den yükleme hatası: {e}")
        
        messagebox.showwarning(
            "Uyarı", 
            "Grup bulunamadı.\n\nSeçenekler:\n1. Yeni analiz yapın\n2. Hazır analiz dosyası seçin\n3. veri.xlsx içeren klasör seçin"
        )
    
    def _raporlari_olustur(self):
        """PNG ve Excel raporlarını oluştur"""
        if not self.cikti_klasoru:
            messagebox.showerror("Hata", "Lütfen çıktı klasörünü seçin.")
            return
        
        if not self.grup_list:
            messagebox.showerror("Hata", "Grup listesi boş. Önce analiz yapın veya grupları yükleyin.")
            return
        
        if not self.dosyalama_engine:
            self.dosyalama_engine = Dosyalama(self.cikti_klasoru)
        else:
            self.dosyalama_engine.klasor_yolu = self.cikti_klasoru
        
        basarili, eksik = self.dosyalama_engine.dosyalari_yukle()
        if not basarili:
            messagebox.showerror("Hata", f"Eksik dosyalar:\n" + "\n".join(eksik))
            return
        
        self._set_status("⏳ Raporlar oluşturuluyor...", 'processing')
        self.btn_raporla.config(state='disabled')
        self.progress['value'] = 0
        self.root.update()
        
        try:
            def progress_callback(idx, total, grup):
                progress_pct = int((idx / total) * 100)
                self.progress['value'] = progress_pct
                self._set_status(f"⏳ Grup {idx}/{total}: {grup[:30]}...", 'processing')
                self.root.update()
            
            islenen = self.dosyalama_engine.tum_gruplari_isle(progress_callback)
            
            self.progress['value'] = 100
            self._set_status(f"✓ {islenen} grup başarıyla raporlandı!", 'success')
            
            output_path = os.path.join(self.cikti_klasoru, VARSAYILAN['OUTPUT_FOLDER'])
            messagebox.showinfo(
                "Tamamlandı",
                f"✅ Raporlama tamamlandı!\n\n"
                f"📊 İşlenen grup: {islenen}\n"
                f"📁 Çıktı klasörü: {output_path}"
            )
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Hata", f"Raporlama sırasında hata:\n\n{str(e)}")
            self._set_status(f"✗ Hata: {str(e)[:50]}...", 'error')
        
        finally:
            self.btn_raporla.config(state='normal')
            self.progress['value'] = 0


def main():
    """Ana fonksiyon"""
    print("=" * 60)
    print("⚡ Kesinti Analiz Pro")
    print("=" * 60)
    
    root = tk.Tk()
    app = BirlesikPanel(root)
    root.mainloop()


if __name__ == "__main__":
    main()
