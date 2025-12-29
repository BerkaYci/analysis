# -*- coding: utf-8 -*-
"""
CM (Customer Management) İşlemleri Modülü
CM.xlsx dosyası ile ilgili tüm işlemler burada yapılır.
"""

import pandas as pd
import os
import sys

# Config'i import et
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import CM_SUTUN_INDEKSLERI, EXCEL_AYARLARI


class CMIslemleri:
    """CM.xlsx dosyası işlemleri için sınıf"""
    
    def __init__(self, cm_dosya_yolu=None):
        """
        CM işlemleri sınıfını başlat.
        
        Args:
            cm_dosya_yolu: CM.xlsx dosyasının tam yolu
        """
        self.cm_dosya_yolu = cm_dosya_yolu
        self.df_cm = None
        
        if cm_dosya_yolu and os.path.exists(cm_dosya_yolu):
            self.yukle(cm_dosya_yolu)
    
    def yukle(self, cm_dosya_yolu):
        """
        CM.xlsx dosyasını yükle.
        
        Args:
            cm_dosya_yolu: CM.xlsx dosyasının tam yolu
            
        Returns:
            bool: Başarılı ise True
        """
        try:
            self.cm_dosya_yolu = cm_dosya_yolu
            self.df_cm = pd.read_excel(cm_dosya_yolu, header=EXCEL_AYARLARI['CM_HEADER_ROW'])
            print(f"✓ CM.xlsx yüklendi: {len(self.df_cm)} satır")
            return True
        except Exception as e:
            print(f"✗ CM.xlsx yüklenemedi: {e}")
            self.df_cm = None
            return False
    
    def yuklu_mu(self):
        """CM dosyası yüklü mü kontrol et"""
        return self.df_cm is not None
    
    def kesinti_ara(self, kesinti_id):
        """
        CM'de kesinti ID'si ile arama yap.
        
        Args:
            kesinti_id: Aranacak kesinti numarası
            
        Returns:
            DataFrame: Eşleşen satırlar
        """
        if not self.yuklu_mu():
            return pd.DataFrame()
        
        try:
            kesinti_id_str = str(kesinti_id).strip()
            col_index = CM_SUTUN_INDEKSLERI['KESINTI_ID']
            mask = self.df_cm.iloc[:, col_index].astype(str).str.strip() == kesinti_id_str
            return self.df_cm[mask]
        except Exception as e:
            print(f"Arama hatası: {e}")
            return pd.DataFrame()
    
    def ortak_w_degerlerini_bul(self, kesinti_id_listesi, elemanlar=None):
        """
        Birden fazla kesinti için ortak müşterilerin (C sütunu) 
        W değerlerini (OMS Ticket ID) bul.
        
        Args:
            kesinti_id_listesi: Kesinti ID'lerinin listesi
            elemanlar: Opsiyonel - kesinti detayları (Baslama, Bitis bilgisi için)
            
        Returns:
            str: Formatlanmış ortak W değerleri
        """
        if not self.yuklu_mu():
            return ""
        
        if len(kesinti_id_listesi) <= 1:
            return ""
        
        # Her kesinti ID için C-W eşleşmelerini sakla
        id_c_w_map = {}
        
        hizmet_no_idx = CM_SUTUN_INDEKSLERI['HIZMET_NO']
        ticket_id_idx = CM_SUTUN_INDEKSLERI['OMS_TICKET_ID']
        
        for kesinti_id in kesinti_id_listesi:
            data = self.kesinti_ara(kesinti_id)
            
            if len(data) == 0:
                continue
            
            c_w_dict = {}
            for _, row in data.iterrows():
                c_val = str(row.iloc[hizmet_no_idx]).strip() if len(row) > hizmet_no_idx else ""
                w_val = str(row.iloc[ticket_id_idx]).strip() if len(row) > ticket_id_idx else ""
                
                if c_val and c_val != '' and c_val.lower() != 'nan':
                    if c_val not in c_w_dict:
                        c_w_dict[c_val] = []
                    if w_val and w_val != '' and w_val.lower() != 'nan':
                        c_w_dict[c_val].append((kesinti_id, w_val))
            
            id_c_w_map[kesinti_id] = c_w_dict
        
        if len(id_c_w_map) <= 1:
            return ""
        
        # Ortak C değerlerini bul
        first_id = list(id_c_w_map.keys())[0]
        common_c_values = set(id_c_w_map[first_id].keys())
        
        for kesinti_id in list(id_c_w_map.keys())[1:]:
            common_c_values = common_c_values.intersection(set(id_c_w_map[kesinti_id].keys()))
        
        if len(common_c_values) == 0:
            return ""
        
        # Her ortak C (Hizmet No) için formatla
        result_parts = []
        for c_val in sorted(common_c_values):
            kesinti_ticket_pairs = []
            seen_pairs = set()
            
            for kesinti_id in id_c_w_map:
                if c_val in id_c_w_map[kesinti_id]:
                    for kesinti_no, ticket_id in id_c_w_map[kesinti_id][c_val]:
                        pair = (kesinti_no, ticket_id)
                        if pair not in seen_pairs:
                            kesinti_ticket_pairs.append(f"{kesinti_no} [{ticket_id}]")
                            seen_pairs.add(pair)
            
            if kesinti_ticket_pairs:
                formatted = f"{c_val} → {', '.join(kesinti_ticket_pairs)}"
                result_parts.append(formatted)
        
        return '\n'.join(result_parts)
    
    def cagri_ticket_idlerini_bul(self, oncesi_kesintiler, sonrasi_kesintiler, 
                                   kesinti_zamanlar_dict):
        """
        Kesinti öncesi/sonrası çağrıların ticket ID'lerini bul.
        
        Args:
            oncesi_kesintiler: Öncesi çağrı olan kesinti ID listesi
            sonrasi_kesintiler: Sonrası çağrı olan kesinti ID listesi
            kesinti_zamanlar_dict: {kesinti_id: (baslama, bitis)} sözlüğü
            
        Returns:
            str: Formatlanmış ticket ID'leri
        """
        if not self.yuklu_mu():
            return ""
        
        ticket_id_idx = CM_SUTUN_INDEKSLERI['OMS_TICKET_ID']
        olusturma_tarihi_idx = CM_SUTUN_INDEKSLERI['OLUSTURMA_TARIHI']
        
        oncesi_tickets = []
        sonrasi_tickets = []
        
        # Öncesi çağrı olan kesintiler
        for kesinti_id in oncesi_kesintiler:
            if kesinti_id not in kesinti_zamanlar_dict:
                continue
            
            kesinti_baslama, kesinti_bitis = kesinti_zamanlar_dict[kesinti_id]
            data = self.kesinti_ara(kesinti_id)
            
            if len(data) == 0:
                continue
            
            ticket_ids = set()  # Mükerrerleri önlemek için set
            for _, row in data.iterrows():
                ticket_id = str(row.iloc[ticket_id_idx]).strip() if len(row) > ticket_id_idx else ""
                olusturma_tarihi_str = str(row.iloc[olusturma_tarihi_idx]).strip() if len(row) > olusturma_tarihi_idx else ""
                
                if not ticket_id or ticket_id == '' or ticket_id.lower() == 'nan':
                    continue
                
                try:
                    olusturma_tarihi = pd.to_datetime(olusturma_tarihi_str, errors='coerce')
                    
                    if pd.isna(olusturma_tarihi):
                        continue
                    
                    if olusturma_tarihi < kesinti_baslama:
                        ticket_ids.add(ticket_id)  # set.add() kullan
                except:
                    continue
            
            if ticket_ids:
                tickets_str = ', '.join(sorted(ticket_ids))  # Sıralı ve tekil
                oncesi_tickets.append(f"{kesinti_id} ({tickets_str})")
        
        # Sonrası çağrı olan kesintiler
        for kesinti_id in sonrasi_kesintiler:
            if kesinti_id not in kesinti_zamanlar_dict:
                continue
            
            kesinti_baslama, kesinti_bitis = kesinti_zamanlar_dict[kesinti_id]
            data = self.kesinti_ara(kesinti_id)
            
            if len(data) == 0:
                continue
            
            ticket_ids = set()  # Mükerrerleri önlemek için set
            for _, row in data.iterrows():
                ticket_id = str(row.iloc[ticket_id_idx]).strip() if len(row) > ticket_id_idx else ""
                olusturma_tarihi_str = str(row.iloc[olusturma_tarihi_idx]).strip() if len(row) > olusturma_tarihi_idx else ""
                
                if not ticket_id or ticket_id == '' or ticket_id.lower() == 'nan':
                    continue
                
                try:
                    olusturma_tarihi = pd.to_datetime(olusturma_tarihi_str, errors='coerce')
                    
                    if pd.isna(olusturma_tarihi):
                        continue
                    
                    if olusturma_tarihi > kesinti_bitis:
                        ticket_ids.add(ticket_id)  # set.add() kullan
                except:
                    continue
            
            if ticket_ids:
                tickets_str = ', '.join(sorted(ticket_ids))  # Sıralı ve tekil
                sonrasi_tickets.append(f"{kesinti_id} ({tickets_str})")
        
        # Format: Öncesi: ... | Sonrası: ...
        result_parts = []
        if oncesi_tickets:
            result_parts.append(f"Öncesi: {'; '.join(oncesi_tickets)}")
        if sonrasi_tickets:
            result_parts.append(f"Sonrası: {'; '.join(sonrasi_tickets)}")
        
        return ' | '.join(result_parts) if result_parts else ""
    
    def get_dataframe(self):
        """CM DataFrame'ini döndür"""
        return self.df_cm

