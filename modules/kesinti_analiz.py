# -*- coding: utf-8 -*-
"""
Kesinti Analiz Modülü
Birleşik kesinti analizi işlemleri.
"""

import pandas as pd
from datetime import timedelta
import os
import sys

# Config ve diğer modülleri import et
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import KESINTI_SUTUNLARI, CM_SUTUN_INDEKSLERI, VERI_SUTUN_INDEKSLERI, VARSAYILAN, DAGITIM_AG_AYARLARI
from modules.cm_islemleri import CMIslemleri
from modules.excel_yardimci import ExcelYardimci


class KesintiAnaliz:
    """Kesinti analizi için ana sınıf"""
    
    def __init__(self):
        """Kesinti analiz sınıfını başlat"""
        self.cm_islemleri = None
        self.df_sonuc = None
        self.kesinti_max_bitis = None
        self.df_tum_kesintiler = None  # TM bazlı tarama için tüm kesintiler
    
    def analiz_yap(self, excel_yolu, tolerans_ayarlari=None):
        """
        Birleşik kesinti analizini gerçekleştir.
        
        Args:
            excel_yolu: Kesinti Excel dosyasının yolu
            tolerans_ayarlari: Ard arda tolerans ayarları (dict: kritik_saat, tolerans_ustu_dk, tolerans_alti_dk)
            
        Returns:
            DataFrame: Analiz sonuçları
        """
        # Tolerans ayarlarını sakla (tüm ard arda analizler için)
        self.tolerans_ayarlari = tolerans_ayarlari or {
            'kritik_saat': 9,
            'tolerans_ustu_dk': 60,
            'tolerans_alti_dk': 15
        }
        # Excel dosyasını oku (tüm sütunlar dahil - TM taraması için)
        df_full = pd.read_excel(excel_yolu, header=3)
        
        
        # CM.xlsx dosyasını yükle
        cm_dosya_yolu = os.path.join(os.path.dirname(excel_yolu), "CM.xlsx")
        if os.path.exists(cm_dosya_yolu):
            self.cm_islemleri = CMIslemleri(cm_dosya_yolu)
        else:
            print(f"✗ CM.xlsx dosyası bulunamadı: {cm_dosya_yolu}")
            self.cm_islemleri = None
        
        # Gerekli sütunları seç
        sutun_adlari = list(KESINTI_SUTUNLARI.values())
        df = df_full[sutun_adlari].copy()
        df.columns = ['INOUT', 'KesintiNo', 'Kademe', 'SebekeUnsuru', 'Baslama', 'Bitis',
                      'ScadaKesintisi', 'SonCagri', 'IlkMusteriDisiCagri', 'IlkMusteriCagri', 
                      'CBSTMNo', 'KaynagaGore', 'ToplamCagri', 'KesijtiSeviyesi']
        
        # Veri temizleme
        df = df.dropna(subset=['KesintiNo', 'SebekeUnsuru', 'Baslama', 'Bitis'])
        df['Baslama'] = pd.to_datetime(df['Baslama'], errors='coerce', dayfirst=True)
        df['Bitis'] = pd.to_datetime(df['Bitis'], errors='coerce', dayfirst=True)
        df['SonCagri'] = pd.to_datetime(df['SonCagri'], errors='coerce', dayfirst=True)
        df['IlkMusteriDisiCagri'] = pd.to_datetime(df['IlkMusteriDisiCagri'], errors='coerce', dayfirst=True)
        df['IlkMusteriCagri'] = pd.to_datetime(df['IlkMusteriCagri'], errors='coerce', dayfirst=True)
        df['ScadaKesintisi'] = df['ScadaKesintisi'].fillna('')
        df['CBSTMNo'] = df['CBSTMNo'].fillna('').astype(str)
        df['ToplamCagri'] = pd.to_numeric(df['ToplamCagri'], errors='coerce').fillna(0).astype(int)
        df['KesijtiSeviyesi'] = df['KesijtiSeviyesi'].fillna('').astype(str)
        
        # TM bazlı tarama için tüm kesintileri sakla
        self.df_tum_kesintiler = df.copy()
        
        # TM bazlı index oluştur (hızlı arama için)
        self._tm_index_olustur()
        
        # Her kesinti no için maksimum bitiş zamanını hesapla
        self.kesinti_max_bitis = df.groupby('KesintiNo')['Bitis'].max().to_dict()
        
        sonuc_list = []
        
        # Şebeke unsuruna göre grupla
        for unsur, grup in df.groupby('SebekeUnsuru'):
            grup = grup.sort_values('Baslama').reset_index(drop=True)
            temp = [grup.iloc[0].to_dict()]
            
            for i in range(1, len(grup)):
                onceki = temp[-1]
                simdiki = grup.iloc[i].to_dict()
                
                # Aynı kesinti no ard arda / iç içe sayılmaz
                if simdiki['KesintiNo'] == onceki['KesintiNo']:
                    sonuc_list.extend(self._zincir_olustur(temp, unsur))
                    temp = [simdiki]
                    continue
                
                # İç içe
                if simdiki['Baslama'] <= onceki['Bitis']:
                    temp.append(simdiki)
                # Ard arda (dinamik tolerans)
                else:
                    # Önceki kesintinin süresine göre tolerans belirle
                    tolerans_dk = self._tolerans_hesapla(onceki)
                    
                    if (simdiki['Baslama'] - onceki['Bitis']) <= timedelta(minutes=tolerans_dk):
                        temp.append(simdiki)
                    else:
                        # Yeni grup
                        sonuc_list.extend(self._zincir_olustur(temp, unsur))
                        temp = [simdiki]
            
            sonuc_list.extend(self._zincir_olustur(temp, unsur))
        
        # ═══════════════════════════════════════════════════════════════
        # TM No Ard Arda Analizi (Dağıtım-AG için)
        # ═══════════════════════════════════════════════════════════════
        tm_ardarda_sonuc = self._tm_no_ardarda_analiz(df)
        sonuc_list.extend(tm_ardarda_sonuc)
        
        # Sonuçları DataFrame'e çevir
        df_sonuc = pd.DataFrame(sonuc_list)
        df_sonuc = df_sonuc[df_sonuc['Tur'] != 'Tekil']
        df_sonuc = df_sonuc.sort_values(['SebekeUnsuru', 'BirlesikBaslama'])
        
        self.df_sonuc = df_sonuc
        return df_sonuc
    
    def _tolerans_hesapla(self, kesinti):
        """Kesinti süresine göre tolerans hesapla."""
        kritik_saat = self.tolerans_ayarlari.get('kritik_saat', 9)
        tolerans_ustu = self.tolerans_ayarlari.get('tolerans_ustu_dk', 60)
        tolerans_alti = self.tolerans_ayarlari.get('tolerans_alti_dk', 15)
        
        # Kesinti süresini saat cinsinden hesapla
        sure_saat = (kesinti['Bitis'] - kesinti['Baslama']).total_seconds() / 3600
        
        if sure_saat >= kritik_saat:
            return tolerans_ustu
        else:
            return tolerans_alti
    
    def _zincir_olustur(self, temp, unsur):
        """Karma zincirleri ikiye ayırır."""
        farklar = []
        icice_list, ardarda_list = [], []
        ic_ice, ard_arda = False, False
        
        for i in range(1, len(temp)):
            fark = (temp[i]['Baslama'] - temp[i - 1]['Bitis']).total_seconds() / 60
            # Önceki kesintinin toleransını hesapla
            tolerans_dk = self._tolerans_hesapla(temp[i - 1])
            
            if fark <= 0:
                ic_ice = True
                icice_list.append((temp[i - 1], temp[i]))
            elif fark > 0 and fark <= tolerans_dk:
                ard_arda = True
                ardarda_list.append((temp[i - 1], temp[i]))
                farklar.append(round(fark, 1))
        
        sonuc = []
        if ic_ice and ard_arda:
            if icice_list:
                sonuc.append(self._tek_zincir(icice_list, unsur, "İç içe", None))
            if ardarda_list:
                sonuc.append(self._tek_zincir(ardarda_list, unsur, "Ard arda", farklar))
            return sonuc
        elif ic_ice:
            return [self._tek_zincir(icice_list, unsur, "İç içe", None)]
        elif ard_arda:
            return [self._tek_zincir(ardarda_list, unsur, "Ard arda", farklar)]
        else:
            pairs = [(x,) for x in temp]
            return [self._tek_zincir(pairs, unsur, "Tekil", None)]
    
    def _tek_zincir(self, pairs, unsur, tur, farklar=None):
        """Tek bir zincir için detayları hesapla."""
        elemanlar = []
        seen = set()
        for pair in pairs:
            for x in pair:
                if x['KesintiNo'] not in seen:
                    elemanlar.append(x)
                    seen.add(x['KesintiNo'])
        
        elemanlar.sort(key=lambda x: x['Baslama'])
        
        # En erken başlama ve en geç bitiş
        ilk = min(x['Baslama'] for x in elemanlar)
        son = max(x['Bitis'] for x in elemanlar)
        
        # Her kesinti için maksimum bitiş zamanını kullan
        if self.kesinti_max_bitis:
            max_bitis_list = [self.kesinti_max_bitis.get(x['KesintiNo'], x['Bitis']) for x in elemanlar]
            son = max(max_bitis_list)
        
        sure_str = ExcelYardimci.format_sure(son - ilk)
        
        # IN/OUT türü belirleme
        inout_degerleri = {x['INOUT'] for x in elemanlar if isinstance(x['INOUT'], str)}
        if len(inout_degerleri) == 1:
            inout_durum = list(inout_degerleri)[0]
        elif len(inout_degerleri) > 1:
            inout_durum = "IN/OUT"
        else:
            inout_durum = "-"
        
        # Scada Kesintisi oranı hesaplama
        toplam_kesinti = len(elemanlar)
        x_olan_sayisi = sum(1 for x in elemanlar if str(x.get('ScadaKesintisi', '')).strip().upper() == 'X')
        scada_orani = f"{x_olan_sayisi}/{toplam_kesinti}"
        
        # Çağrı durumu analizi
        cagri_durumu, oncesi_kesintiler, sonrasi_kesintiler = self._analyze_cagri_durumu(elemanlar)
        
        # Tür bilgisi
        tur_final = f"{tur} - {cagri_durumu}" if cagri_durumu else tur
        
        # OMS Ticket ID'leri
        oms_ticket_ids = ""
        if (oncesi_kesintiler or sonrasi_kesintiler) and self.cm_islemleri:
            kesinti_zamanlar = {
                elem['KesintiNo']: (
                    elem['Baslama'],
                    self.kesinti_max_bitis.get(elem['KesintiNo'], elem['Bitis']) if self.kesinti_max_bitis else elem['Bitis']
                )
                for elem in elemanlar
            }
            oms_ticket_ids = self.cm_islemleri.cagri_ticket_idlerini_bul(
                oncesi_kesintiler, sonrasi_kesintiler, kesinti_zamanlar
            )
        
        kesinti_noktalivirgul = ";".join(str(x['KesintiNo']) for x in elemanlar)
        zamanlar = "\n".join([
            f"{i + 1}) {x['KesintiNo']} [{x['Kademe']}] "
            f"{x['Baslama'].strftime('%d.%m.%Y %H:%M:%S')} → {x['Bitis'].strftime('%d.%m.%Y %H:%M:%S')}"
            for i, x in enumerate(elemanlar)
        ])
        
        # Ortak W değerlerini hesapla (CM varsa)
        ortak_w_degerleri = ""
        if self.cm_islemleri:
            kesinti_id_listesi = [x['KesintiNo'] for x in elemanlar]
            ortak_w_degerleri = self.cm_islemleri.ortak_w_degerlerini_bul(kesinti_id_listesi)
        
        # Kademe ve Kaynağa Göre bilgisini al (ilk elemandan)
        kademe = elemanlar[0].get('Kademe', '') if elemanlar else ''
        kaynaga_gore = elemanlar[0].get('KaynagaGore', '') if elemanlar else ''
        
        # Toplam Çağrı Sayısı (her kesintinin çağrı sayısı ; ile ayrılmış)
        cagri_sayilari = [str(int(x.get('ToplamCagri', 0))) for x in elemanlar]
        toplam_cagri_sayisi = "; ".join(cagri_sayilari)
        
        # Kesinti Seviyesi (benzersiz seviyeleri birleştir)
        seviyeler = set()
        for x in elemanlar:
            seviye = str(x.get('KesijtiSeviyesi', '')).strip()
            if seviye and seviye.lower() != 'nan':
                seviyeler.add(seviye)
        kesinti_seviyesi = "; ".join(sorted(seviyeler)) if seviyeler else ""
        
        # Dağıtım-AG için TM bazlı kesinti taraması
        tm_kesintileri = ""
        dagitim_ag_deger = DAGITIM_AG_AYARLARI.get('KAYNAGA_GORE_DEGER', 'Dağıtım-AG')
        is_dagitim_ag = str(kaynaga_gore).strip() == dagitim_ag_deger
        
        if is_dagitim_ag:
            tm_kesintileri = self._tm_bazli_kesinti_tara(elemanlar, ilk, son)
        
        return {
            'SebekeUnsuru': unsur,
            'IN-OUT Durumu': inout_durum,
            'BirlesikBaslama': ilk.strftime('%d.%m.%Y %H:%M:%S'),
            'BirlesikBitis': son.strftime('%d.%m.%Y %H:%M:%S'),
            'Süre (hh:mm:ss)': sure_str,
            'Tur': tur_final,
            'Ardışık Farklar (dk)': '; '.join(map(str, farklar)) if farklar else '',
            'İlgiliKesintiler(;)': kesinti_noktalivirgul,
            'KesintiZamanlari': zamanlar,
            'Scada Kesintisi Oranı': scada_orani,
            'Toplam Çağrı Sayısı': toplam_cagri_sayisi,
            'Kesinti Seviyesi': kesinti_seviyesi,
            'OMS Ticket IDs': oms_ticket_ids,
            'Ortak W Değerleri': ortak_w_degerleri,
            'TM Kesintileri': tm_kesintileri
        }
    
    def _tm_no_ardarda_analiz(self, df):
        """
        TM No bazlı ard arda analizi (Dağıtım-AG için).
        
        Şebeke Unsuru farklı olsa bile aynı CBS TM No'ya sahip
        Dağıtım-AG kesintilerini gruplar.
        
        Args:
            df: Kesinti verilerini içeren DataFrame
            
        Returns:
            list: TM No Ard Arda sonuçları
        """
        sonuc_list = []
        
        # Dağıtım-AG olanları filtrele
        dagitim_ag_deger = DAGITIM_AG_AYARLARI.get('KAYNAGA_GORE_DEGER', 'Dağıtım-AG')
        df_dagitim = df[df['KaynagaGore'].astype(str).str.strip() == dagitim_ag_deger].copy()
        
        if df_dagitim.empty:
            return sonuc_list
        
        # TM numaralarını temizle
        df_dagitim['CBSTMNoTemiz'] = df_dagitim['CBSTMNo'].apply(self._tm_no_temizle)
        
        # Boş TM'leri çıkar
        df_dagitim = df_dagitim[df_dagitim['CBSTMNoTemiz'] != '']
        
        if df_dagitim.empty:
            return sonuc_list
        
        print(f"✓ TM No Ard Arda analizi: {len(df_dagitim)} Dağıtım-AG kesintisi")
        
        # Tolerans ayarlarını al
        kritik_saat = self.tolerans_ayarlari.get('kritik_saat', 9)
        tolerans_ustu = self.tolerans_ayarlari.get('tolerans_ustu_dk', 60)
        tolerans_alti = self.tolerans_ayarlari.get('tolerans_alti_dk', 15)
        
        # CBS TM No'ya göre grupla
        for tm_no, tm_grup in df_dagitim.groupby('CBSTMNoTemiz'):
            if len(tm_grup) < 2:
                continue  # En az 2 kesinti olmalı
            
            # Başlama zamanına göre sırala
            tm_grup = tm_grup.sort_values('Baslama').reset_index(drop=True)
            
            # Ard arda grupları bul
            temp = [tm_grup.iloc[0].to_dict()]
            
            for i in range(1, len(tm_grup)):
                onceki = temp[-1]
                simdiki = tm_grup.iloc[i].to_dict()
                
                # Aynı kesinti no'yu atla
                if simdiki['KesintiNo'] == onceki['KesintiNo']:
                    sonuc_list.extend(self._tm_zincir_olustur(temp, tm_no))
                    temp = [simdiki]
                    continue
                
                # Aynı Şebeke Unsuru varsa atla (normal analiz kapsar)
                if simdiki['SebekeUnsuru'] == onceki['SebekeUnsuru']:
                    sonuc_list.extend(self._tm_zincir_olustur(temp, tm_no))
                    temp = [simdiki]
                    continue
                
                # Kesinti süresini hesapla (önceki)
                onceki_sure_saat = (onceki['Bitis'] - onceki['Baslama']).total_seconds() / 3600
                
                # Toleransı belirle
                if onceki_sure_saat >= kritik_saat:
                    tolerans = tolerans_ustu
                else:
                    tolerans = tolerans_alti
                
                # Ard arda kontrolü
                fark_dakika = (simdiki['Baslama'] - onceki['Bitis']).total_seconds() / 60
                
                if 0 < fark_dakika <= tolerans:
                    temp.append(simdiki)
                else:
                    sonuc_list.extend(self._tm_zincir_olustur(temp, tm_no))
                    temp = [simdiki]
            
            # Son grubu ekle
            sonuc_list.extend(self._tm_zincir_olustur(temp, tm_no))
        
        return sonuc_list
    
    def _tm_zincir_olustur(self, temp, tm_no):
        """TM No Ard Arda zinciri oluştur."""
        if len(temp) < 2:
            return []  # Tek eleman zincir sayılmaz
        
        elemanlar = temp
        elemanlar.sort(key=lambda x: x['Baslama'])
        
        # Farkları hesapla
        farklar = []
        for i in range(1, len(elemanlar)):
            fark = (elemanlar[i]['Baslama'] - elemanlar[i-1]['Bitis']).total_seconds() / 60
            farklar.append(round(fark, 1))
        
        # En erken başlama ve en geç bitiş
        ilk = min(x['Baslama'] for x in elemanlar)
        son = max(x['Bitis'] for x in elemanlar)
        
        # Her kesinti için maksimum bitiş zamanını kullan
        if self.kesinti_max_bitis:
            max_bitis_list = [self.kesinti_max_bitis.get(x['KesintiNo'], x['Bitis']) for x in elemanlar]
            son = max(max_bitis_list)
        
        sure_str = ExcelYardimci.format_sure(son - ilk)
        
        # Şebeke Unsurları (farklı olabilir)
        sebeke_unsurlari = list(set(x['SebekeUnsuru'] for x in elemanlar))
        sebeke_unsuru_str = " | ".join(sebeke_unsurlari)
        
        # IN/OUT türü
        inout_degerleri = {x['INOUT'] for x in elemanlar if isinstance(x['INOUT'], str)}
        if len(inout_degerleri) == 1:
            inout_durum = list(inout_degerleri)[0]
        elif len(inout_degerleri) > 1:
            inout_durum = "IN/OUT"
        else:
            inout_durum = "-"
        
        # Scada oranı
        toplam = len(elemanlar)
        x_sayisi = sum(1 for x in elemanlar if str(x.get('ScadaKesintisi', '')).strip().upper() == 'X')
        scada_orani = f"{x_sayisi}/{toplam}"
        
        # Çağrı durumu analizi
        cagri_durumu, oncesi_kesintiler, sonrasi_kesintiler = self._analyze_cagri_durumu(elemanlar)
        
        # Tür bilgisi
        tur = "TM No Ard Arda"
        tur_final = f"{tur} - {cagri_durumu}" if cagri_durumu else tur
        
        # OMS Ticket ID'leri
        oms_ticket_ids = ""
        if (oncesi_kesintiler or sonrasi_kesintiler) and self.cm_islemleri:
            kesinti_zamanlar = {
                elem['KesintiNo']: (
                    elem['Baslama'],
                    self.kesinti_max_bitis.get(elem['KesintiNo'], elem['Bitis']) if self.kesinti_max_bitis else elem['Bitis']
                )
                for elem in elemanlar
            }
            oms_ticket_ids = self.cm_islemleri.cagri_ticket_idlerini_bul(
                oncesi_kesintiler, sonrasi_kesintiler, kesinti_zamanlar
            )
        
        kesinti_noktalivirgul = ";".join(str(x['KesintiNo']) for x in elemanlar)
        zamanlar = "\n".join([
            f"{i + 1}) {x['KesintiNo']} [{x['SebekeUnsuru']}] "
            f"{x['Baslama'].strftime('%d.%m.%Y %H:%M:%S')} → {x['Bitis'].strftime('%d.%m.%Y %H:%M:%S')}"
            for i, x in enumerate(elemanlar)
        ])
        
        # Ortak W değerleri
        ortak_w_degerleri = ""
        if self.cm_islemleri:
            kesinti_id_listesi = [x['KesintiNo'] for x in elemanlar]
            ortak_w_degerleri = self.cm_islemleri.ortak_w_degerlerini_bul(kesinti_id_listesi)
        
        # Toplam Çağrı Sayısı (her kesintinin çağrı sayısı ; ile ayrılmış)
        cagri_sayilari = [str(int(x.get('ToplamCagri', 0))) for x in elemanlar]
        toplam_cagri_sayisi = "; ".join(cagri_sayilari)
        
        # Kesinti Seviyesi (benzersiz seviyeleri birleştir)
        seviyeler = set()
        for x in elemanlar:
            seviye = str(x.get('KesijtiSeviyesi', '')).strip()
            if seviye and seviye.lower() != 'nan':
                seviyeler.add(seviye)
        kesinti_seviyesi = "; ".join(sorted(seviyeler)) if seviyeler else ""
        
        return [{
            'SebekeUnsuru': f"TM:{tm_no} ({sebeke_unsuru_str})",
            'IN-OUT Durumu': inout_durum,
            'BirlesikBaslama': ilk.strftime('%d.%m.%Y %H:%M:%S'),
            'BirlesikBitis': son.strftime('%d.%m.%Y %H:%M:%S'),
            'Süre (hh:mm:ss)': sure_str,
            'Tur': tur_final,
            'Ardışık Farklar (dk)': '; '.join(map(str, farklar)) if farklar else '',
            'İlgiliKesintiler(;)': kesinti_noktalivirgul,
            'KesintiZamanlari': zamanlar,
            'Scada Kesintisi Oranı': scada_orani,
            'Toplam Çağrı Sayısı': toplam_cagri_sayisi,
            'Kesinti Seviyesi': kesinti_seviyesi,
            'OMS Ticket IDs': oms_ticket_ids,
            'Ortak W Değerleri': ortak_w_degerleri,
            'TM Kesintileri': ''  # TM No Ard Arda için bu alan boş
        }]
    
    def _tm_no_temizle(self, tm_no):
        """TM numarasını temizle (float'tan int'e çevir)"""
        if not tm_no or pd.isna(tm_no):
            return ""
        tm_str = str(tm_no).strip()
        if tm_str.lower() == 'nan' or tm_str == '':
            return ""
        # Float'tan int'e çevir (1003007.0 -> 1003007)
        try:
            return str(int(float(tm_str)))
        except:
            return tm_str
    
    def _tm_index_olustur(self):
        """TM bazlı index oluştur (hızlı arama için)"""
        self.tm_kesinti_index = {}
        
        if self.df_tum_kesintiler is None:
            return
        
        for _, row in self.df_tum_kesintiler.iterrows():
            tm_no = self._tm_no_temizle(row.get('CBSTMNo', ''))
            if not tm_no:
                continue
            
            kesinti_bilgi = {
                'KesintiNo': row['KesintiNo'],
                'Baslama': row['Baslama']
            }
            
            if tm_no not in self.tm_kesinti_index:
                self.tm_kesinti_index[tm_no] = []
            self.tm_kesinti_index[tm_no].append(kesinti_bilgi)
        
        print(f"✓ TM index oluşturuldu: {len(self.tm_kesinti_index)} farklı TM")
    
    def _tm_bazli_kesinti_tara(self, elemanlar, baslama, bitis):
        """
        Dağıtım-AG kesintileri için TM bazlı kesinti taraması (optimize edilmiş).
        
        Args:
            elemanlar: Gruptaki kesinti elemanları
            baslama: Grubun başlangıç zamanı
            bitis: Grubun bitiş zamanı
            
        Returns:
            str: Aynı TM'deki diğer kesinti numaraları (;'li)
        """
        if not hasattr(self, 'tm_kesinti_index') or not self.tm_kesinti_index:
            return ""
        
        if len(elemanlar) == 0:
            return ""
        
        # Gruptaki kesinti numaralarını al
        grup_kesinti_nolari = {x['KesintiNo'] for x in elemanlar}
        
        # Gruptaki TM numaralarını al (temizlenmiş)
        tm_nolari = set()
        for elem in elemanlar:
            tm_no = self._tm_no_temizle(elem.get('CBSTMNo', ''))
            if tm_no:
                tm_nolari.add(tm_no)
        
        if not tm_nolari:
            return ""
        
        # Tarama aralığını hesapla (±12 saat)
        tarama_saat = DAGITIM_AG_AYARLARI.get('TARAMA_SAAT', 12)
        tarama_baslama = baslama - timedelta(hours=tarama_saat)
        tarama_bitis = bitis + timedelta(hours=tarama_saat)
        
        # Index kullanarak hızlı arama
        bulunan_kesintiler = set()
        
        for tm_no in tm_nolari:
            if tm_no not in self.tm_kesinti_index:
                continue
            
            # Sadece bu TM'deki kesintileri kontrol et
            for kesinti_bilgi in self.tm_kesinti_index[tm_no]:
                kesinti_no = kesinti_bilgi['KesintiNo']
                row_baslama = kesinti_bilgi['Baslama']
                
                # Kendi grubundaki kesintileri atla
                if kesinti_no in grup_kesinti_nolari:
                    continue
                
                # Zaman aralığında mı?
                if pd.isna(row_baslama):
                    continue
                
                if tarama_baslama <= row_baslama <= tarama_bitis:
                    bulunan_kesintiler.add(int(kesinti_no))
        
        if bulunan_kesintiler:
            return ";".join(str(k) for k in sorted(bulunan_kesintiler))
        
        return ""
    
    def _analyze_cagri_durumu(self, elemanlar):
        """Çağrıların her kesintinin kendi sınırlarına göre durumunu analiz et."""
        oncesi_cagri = False
        sonrasi_cagri = False
        oncesi_kesintiler = set()
        sonrasi_kesintiler = set()
        
        for elem in elemanlar:
            kesinti_no = elem['KesintiNo']
            kesinti_kendi_baslama = elem['Baslama']
            
            # Maksimum bitiş zamanını al
            if self.kesinti_max_bitis and kesinti_no in self.kesinti_max_bitis:
                kesinti_kendi_bitis = self.kesinti_max_bitis[kesinti_no]
            else:
                kesinti_kendi_bitis = elem['Bitis']
            
            # Çağrı zamanlarını kontrol et
            cagri_zamanlari = [
                elem.get('SonCagri'),
                elem.get('IlkMusteriDisiCagri'),
                elem.get('IlkMusteriCagri')
            ]
            
            for cagri_zaman in cagri_zamanlari:
                if pd.isna(cagri_zaman):
                    continue
                
                if cagri_zaman < kesinti_kendi_baslama:
                    oncesi_cagri = True
                    oncesi_kesintiler.add(kesinti_no)
                elif cagri_zaman > kesinti_kendi_bitis:
                    sonrasi_cagri = True
                    sonrasi_kesintiler.add(kesinti_no)
        
        # Durum belirle
        durum = ""
        if oncesi_cagri and sonrasi_cagri:
            durum = "Kesinti Öncesi ve Sonrası Çağrı"
        elif oncesi_cagri:
            durum = "Kesinti Öncesi Çağrı"
        elif sonrasi_cagri:
            durum = "Kesinti Sonrası Çağrı"
        
        return durum, list(oncesi_kesintiler), list(sonrasi_kesintiler)
    
    def kaydet(self, dosya_yolu):
        """
        Analiz sonuçlarını Excel'e kaydet.
        
        Args:
            dosya_yolu: Hedef dosya yolu
        """
        if self.df_sonuc is None or self.df_sonuc.empty:
            print("✗ Kaydedilecek sonuç yok!")
            return False
        
        ExcelYardimci.kaydet_bicimli(self.df_sonuc, dosya_yolu)
        return True
    
    def sonuclari_al(self):
        """Analiz sonuçlarını döndür"""
        return self.df_sonuc
    
    def cm_islemlerini_al(self):
        """CM işlemleri nesnesini döndür"""
        return self.cm_islemleri

