import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String, Float, ForeignKey, Boolean, LargeBinary
from sqlalchemy.orm import declarative_base, sessionmaker, relationship
from datetime import datetime, date, timedelta
import io
import re
import plotly.express as px # YENİ: Profesyonel grafikler için eklendi

# 1. SAYFA YAPILANDIRMASI
st.set_page_config(
    page_title="Tepe Servis | Kurumsal Denetim & Harita Platformu", 
    page_icon="🗺️", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. VERİTABANI VE GELİŞMİŞ MODELLER
# Veritabanı URL'sini Streamlit Secrets (Gizli Değişkenler) içinden güvenli bir şekilde çekiyoruz
try:
    db_url = st.secrets["DB_URL"]
except:
    st.error("⚠️ Veritabanı bağlantı adresi (DB_URL) bulunamadı. Lütfen Streamlit Secrets ayarlarını kontrol edin.")
    st.stop()

# SQLAlchemy bazı durumlarda 'postgres://' yerine 'postgresql://' formatını zorunlu kılar, bunu garantiye alıyoruz
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

# PostgreSQL için check_same_thread parametresine gerek yoktur
engine = create_engine(db_url)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class ProjectDB(Base):
    __tablename__ = "projects"
    id = Column(Integer, primary_key=True, index=True)
    proje_kodu = Column(String, unique=True, index=True)
    proje_adi = Column(String)
    adres = Column(String)
    il = Column(String)
    bolge = Column(String)
    direktor = Column(String)
    bolge_md = Column(String)
    opr_muduru = Column(String)
    takim_lideri = Column(String)
    opr_uzmani = Column(String)
    denetim_gorevlisi = Column(String)
    durum = Column(String)
    notlar = Column(String)
    per_sayisi = Column(Integer, nullable=True)
    is_buyuk_proje = Column(Boolean, default=False)
    koordinat = Column(String, nullable=True)
    audits = relationship("AuditDB", back_populates="project", cascade="all, delete-orphan")
    dof_list = relationship("DofDB", back_populates="project", cascade="all, delete-orphan")
    documents = relationship("DocumentDB", back_populates="project", cascade="all, delete-orphan")

class AuditDB(Base):
    __tablename__ = "audits"
    id = Column(Integer, primary_key=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"))
    denetim_turu = Column(String)
    tarih = Column(String)
    puan = Column(Float, nullable=True)
    gorevlisi = Column(String, nullable=True)
    notlar = Column(String, nullable=True)
    project = relationship("ProjectDB", back_populates="audits")

class DofDB(Base):
    __tablename__ = "dofs"
    id = Column(Integer, primary_key=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"))
    dof_no = Column(String)
    uygunsuzluk_alan = Column(String)
    kok_neden = Column(String)
    uygunsuzluk_tanimi = Column(String)
    aksiyon_plani = Column(String, nullable=True)
    sorumlu_kisi = Column(String, nullable=True)
    acilis_tarihi = Column(String)
    termin_tarihi = Column(String, nullable=True)
    durum = Column(String)
    kapatma_notu = Column(String, nullable=True)
    project = relationship("ProjectDB", back_populates="dof_list")

class DocumentDB(Base):
    __tablename__ = "documents"
    id = Column(Integer, primary_key=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"))
    dosya_adi = Column(String)
    dosya_turu = Column(String)
    yukleme_tarihi = Column(String)
    yukleyen_kisi = Column(String, nullable=True)
    dosya_icerik = Column(LargeBinary)
    project = relationship("ProjectDB", back_populates="documents")

Base.metadata.create_all(bind=engine)

def get_db():
    db = SessionLocal()
    try:
        return db
    finally:
        db.close()

def clean_val(val, default=""):
    if pd.isnull(val) or str(val).strip() in ['NaT', 'nan', '?', 'None', '']:
        return default
    return str(val).strip()

def clean_date(val):
    res = clean_val(val, default=None)
    if not res:
        return None
    try:
        return str(pd.to_datetime(res)).split(' ')[0]
    except:
        return res

# 3. YIKICI OLMAYAN ANA EXCEL SENKRONİZASYONU (YÜKSEK HIZLI VERSİYON)
def sync_excel_without_losing_dofs(uploaded_file):
    db = SessionLocal()
    try:
        Base.metadata.create_all(bind=engine)
        df = pd.read_excel(uploaded_file, sheet_name=0)
        df = df.where(pd.notnull(df), None)
        progress_bar = st.progress(0)
        total_rows = len(df)
        yeni_p = 0
        guncel_p = 0
        yeni_d = 0
        
        for idx, row in df.iterrows():
            pkodu = clean_val(row.get('PROJE KODU'), default=f"KODSUZ-{idx}")
            padi = clean_val(row.get('PROJE'), "İsimsiz Proje")
            pnot = clean_val(row.get('NOTLAR'), "")
            
            is_buyuk = False
            if "BÜYÜK" in padi.upper() or "BÜYÜK" in pnot.upper():
                is_buyuk = True
                
            try:
                per_sayi = int(float(row.get('PER. SAYISI'))) if pd.notnull(row.get('PER. SAYISI')) else 0
            except:
                per_sayi = 0
                
            if per_sayi >= 30:
                is_buyuk = True

            project = db.query(ProjectDB).filter(ProjectDB.proje_kodu == pkodu).first()
            if not project:
                project = ProjectDB(
                    proje_kodu=pkodu,
                    proje_adi=padi,
                    adres=clean_val(row.get('ADRESİ')),
                    il=clean_val(row.get('İL')),
                    bolge=clean_val(row.get('BÖLGE')),
                    direktor=clean_val(row.get('DİREKTÖR')),
                    bolge_md=clean_val(row.get('BÖLGE MD.')),
                    opr_muduru=clean_val(row.get('OPR.MÜDÜRÜ')),
                    takim_lideri=clean_val(row.get('TAKIM LİDERİ')),
                    opr_uzmani=clean_val(row.get('OPR. UZMANI')),
                    denetim_gorevlisi=clean_val(row.get('DENETİM GÖREVLİSİ')),
                    durum=clean_val(row.get('YAPILDI'), "-"),
                    notlar=pnot,
                    per_sayisi=per_sayi,
                    is_buyuk_proje=is_buyuk
                )
                db.add(project)
                # OPTİMİZASYON 1: commit() yerine flush() kullanıyoruz. 
                # Diske yazmayı erteler ama bize o projenin ID'sini anında verir.
                db.flush() 
                yeni_p += 1
            else:
                project.proje_adi = padi
                project.adres = clean_val(row.get('ADRESİ'), project.adres)
                project.il = clean_val(row.get('İL'), project.il)
                project.bolge = clean_val(row.get('BÖLGE'), project.bolge)
                project.direktor = clean_val(row.get('DİREKTÖR'), project.direktor)
                project.bolge_md = clean_val(row.get('BÖLGE MD.'), project.bolge_md)
                project.opr_muduru = clean_val(row.get('OPR.MÜDÜRÜ'), project.opr_muduru)
                project.durum = clean_val(row.get('YAPILDI'), project.durum)
                project.notlar = pnot
                project.per_sayisi = per_sayi
                project.is_buyuk_proje = is_buyuk
                guncel_p += 1
                # OPTİMİZASYON 2: Buradaki döngü içi db.commit()'i sildik.
            
            audits_to_check = [
                ('1. Denetim', 'DENETİM TARİHİ', 'DEN PUAN', 'DENETİM GÖREVLİSİ'),
                ('2. Denetim', '2. DENETİM TARİHİ', '2. DENETİM PUANI', '2. DENETİM GÖREVLİSİ'),
                ('Tekrar Denetimi', 'TEKRAR DENETİM TARİHİ', 'TEKRAR DENETİM PUANI', 'TEKRAR DENETİM GÖREVLİSİ'),
                ('Dış Denetim', 'DIŞ DENETİM TARİHİ', 'DIŞ DEN PUAN', None)
            ]
            
            for tur, tar_col, puan_col, gor_col in audits_to_check:
                tarih_val = clean_date(row.get(tar_col))
                puan_val = None
                
                try:
                    if pd.notnull(row.get(puan_col)):
                        puan_val = float(row.get(puan_col))
                except:
                    puan_val = None
                
                gor_val = clean_val(row.get(gor_col)) if gor_col else clean_val(row.get('DENETİM GÖREVLİSİ'))
                
                if puan_val is not None or tarih_val is not None:
                    mevcut = db.query(AuditDB).filter(
                        AuditDB.project_id == project.id,
                        AuditDB.denetim_turu == tur
                    ).first()
                    
                    if not mevcut:
                        audit = AuditDB(
                            project_id=project.id,
                            denetim_turu=tur,
                            tarih=tarih_val if tarih_val else "",
                            puan=puan_val,
                            gorevlisi=gor_val
                        )
                        db.add(audit)
                        yeni_d += 1
                    else:
                        mevcut.tarih = tarih_val if tarih_val else mevcut.tarih
                        mevcut.puan = puan_val
                        mevcut.gorevlisi = gor_val
                        # OPTİMİZASYON 3: Buradaki döngü içi db.commit()'i sildik.
                    
            if idx % 50 == 0:
                progress_bar.progress(min(int((idx / total_rows) * 100), 100))
                
        # OPTİMİZASYON 4: Bütün Excel okunduktan sonra veritabanına TEK BİR ağ isteği atarak hepsini topluca kaydediyoruz.
        db.commit()
        progress_bar.progress(100)
        return total_rows, yeni_p, guncel_p, yeni_d
    except Exception as e:
        # Hata anında yarım kalan işlemleri temizle
        db.rollback() 
        st.error(f"Senkronizasyon Hatası: {e}")
        return 0, 0, 0, 0
    finally:
        db.close()

# 4. KOORDİNAT EXCEL'İ SENKRONİZASYONU
def sync_coordinates(uploaded_file):
    db = SessionLocal()
    try:
        df = pd.read_excel(uploaded_file)
        pk_col = next((c for c in df.columns if "KOD" in str(c).upper()), None)
        koor_col = next((c for c in df.columns if "KOOR" in str(c).upper()), None)
        
        if not pk_col or not koor_col:
            return False, "Excel dosyanızda 'PROJE KODU' ve 'KOORDİNAT' sütunları bulunamadı. Lütfen başlıkları kontrol edin."
            
        count = 0
        for idx, row in df.iterrows():
            pk = str(row[pk_col]).strip()
            coord = str(row[koor_col]).strip() if pd.notnull(row[koor_col]) else None
            
            if pk and coord and coord.lower() != "nan":
                proj = db.query(ProjectDB).filter(ProjectDB.proje_kodu == pk).first()
                if proj:
                    proj.koordinat = coord
                    count += 1
        db.commit()
        return True, count
    except Exception as e:
        return False, str(e)
    finally:
        db.close()

# 5. YAN MENÜ (SIDEBAR)
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2912/2912773.png", width=60)
    st.title("Tepe Servis Denetim")
    st.caption("Denetim Paneli")
    st.markdown("---")
    sayfa = st.radio(
        "🗂️ Modül Seçin",
        [
            "📊 Yönetsel Kontrol Paneli", 
            "🚨 Akıllı Ziyaret Radarı & Alarmlar",
            "📋 Gelişmiş Proje Havuzu", 
            "🎯 Proje Detay, DÖF & 🗺️ Harita", 
            "🛠️ DÖF Takip & Kök Neden Analizi",
            "➕ Denetim & Ziyaret Girişi", 
            "📈 Denetçi Performans Karnesi",
            "🔄 Excel & Koordinat Yükleme"
        ]
    )

db = get_db()
all_projects = db.query(ProjectDB).all()
all_audits = db.query(AuditDB).all()
all_dofs = db.query(DofDB).all()

# ==========================================
# SAYFA 1: YÖNETSEL KONTROL PANELİ
# ==========================================
if sayfa == "📊 Yönetsel Kontrol Paneli":
    st.header("📊 Operasyonel Denetim ve Kalite Yönetim Paneli")
    st.markdown("---")
    if not all_projects:
        st.warning("⚠️ Sistemde veri bulunmamaktadır. Lütfen yan menüden 'Excel & Koordinat Yükleme' sekmesine girerek ilk listenizi yükleyin.")
    else:
        toplam_proje = len(all_projects)
        buyuk_projeler = len([p for p in all_projects if p.is_buyuk_proje])
        puanli_denetimler = [a.puan for a in all_audits if a.puan is not None]
        ort_puan = sum(puanli_denetimler) / len(puanli_denetimler) if puanli_denetimler else 0
        acik_dof_sayisi = len([d for d in all_dofs if d.durum != "KAPANDI"])
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("🏢 Toplam Proje", f"{toplam_proje:,}", f"{buyuk_projeler} Büyük Proje ⭐")
        col2.metric("⭐ Türkiye Yaşayan Puan", f"{ort_puan:.1f} / 100", delta="Hedef: 90.0")
        col3.metric("🚨 Açık / Bekleyen DÖF", f"{acik_dof_sayisi}", delta="-Çözüm Bekleyen" if acik_dof_sayisi > 0 else "Temiz", delta_color="inverse")
        col4.metric("📍 Haritalanmış Proje", f"{len([p for p in all_projects if p.koordinat])}", "Sistemdeki GPS Verisi")
        st.markdown("---")
        
        st.subheader("🚨 Kırmızı Bayrak: Riskli Projeler (89 Puan ve Altı)")
        riskli_projeler = []
        for p in all_projects:
            for a in p.audits:
                if a.puan is not None and a.puan <= 89.0:
                    acik_dof = len([d for d in p.dof_list if d.durum != "KAPANDI"])
                    riskli_projeler.append({
                        "Proje Kodu": p.proje_kodu,
                        "Proje Adı": p.proje_adi,
                        "İl / Bölge": f"{p.il} / {p.bolge}",
                        "Operasyon Md.": p.opr_muduru,
                        "Denetim": a.denetim_turu,
                        "Tarih": a.tarih,
                        "Puan ⚠️": a.puan,
                        "Açık DÖF": f"{acik_dof} Adet"
                    })
        
        if riskli_projeler:
            st.dataframe(pd.DataFrame(riskli_projeler), use_container_width=True, hide_index=True)
        else:
            st.success("🎉 Sistemde 89 puanın altında riskli proje bulunmamaktadır.")

        st.markdown("---")
        st.subheader("👔 Direktör Bazlı Performans & DÖF Kapatma Hızı")
        
        direktor_stats = {}
        for p in all_projects:
            dir_adi = str(p.direktor).strip()
            if dir_adi == "" or dir_adi == "None":
                dir_adi = "Belirtilmemiş"
                
            if dir_adi not in direktor_stats:
                direktor_stats[dir_adi] = {
                    "Proje Sayısı": 0, 
                    "Toplam Puan": 0, 
                    "Puanlı Denetim": 0,
                    "Toplam DÖF": 0,
                    "Kapanan DÖF": 0
                }
            
            direktor_stats[dir_adi]["Proje Sayısı"] += 1
            
            for a in p.audits:
                if a.puan is not None:
                    direktor_stats[dir_adi]["Toplam Puan"] += a.puan
                    direktor_stats[dir_adi]["Puanlı Denetim"] += 1
                    
            for d in p.dof_list:
                direktor_stats[dir_adi]["Toplam DÖF"] += 1
                if str(d.durum).upper() == "KAPANDI":
                    direktor_stats[dir_adi]["Kapanan DÖF"] += 1
                    
        dir_rows = []
        for d_adi, stats in direktor_stats.items():
            ort_puan = (stats["Toplam Puan"] / stats["Puanlı Denetim"]) if stats["Puanlı Denetim"] > 0 else 0
            kapatma_orani = (stats["Kapanan DÖF"] / stats["Toplam DÖF"] * 100) if stats["Toplam DÖF"] > 0 else 100
            
            dir_rows.append({
                "Direktör": d_adi,
                "Proje Sayısı": stats["Proje Sayısı"],
                "Ortalama Denetim Puanı": f"{ort_puan:.1f}",
                "Toplam DÖF (Sorun)": stats["Toplam DÖF"],
                "Kapanan DÖF": stats["Kapanan DÖF"],
                "DÖF Çözüm Başarısı": f"%{kapatma_orani:.1f}"
            })
            
        if dir_rows:
            df_dir = pd.DataFrame(dir_rows).sort_values(by="Ortalama Denetim Puanı", ascending=False)
            st.dataframe(df_dir, use_container_width=True, hide_index=True)
            
            c_bar1, c_bar2 = st.columns(2)
            with c_bar1:
                st.markdown("**Direktör Bazlı Puan Ortalaması**")
                df_dir["Puan_Float"] = df_dir["Ortalama Denetim Puanı"].astype(float)
                st.bar_chart(df_dir.set_index("Direktör")["Puan_Float"])
            with c_bar2:
                st.markdown("**DÖF Çözüm (Kapatma) Oranları (%)**")
                df_dir["Oran_Float"] = df_dir["DÖF Çözüm Başarısı"].str.replace("%", "").astype(float)
                st.bar_chart(df_dir.set_index("Direktör")["Oran_Float"])

# ==========================================
# SAYFA 2: AKILLI ZİYARET RADARI
# ==========================================
elif sayfa == "🚨 Akıllı Ziyaret Radarı & Alarmlar":
    st.header("🚨 Akıllı Denetim Frekans Radarı")
    if all_projects:
        today = date.today()
        eski_gizle = st.checkbox("✅ 2024-2025 Yıllarından Kalan Eski Projeleri Gizle", value=True)
        muaf_durumlar = ["DENETİM DIŞI", "KAPANDI", "BİTTİ", "İLAÇLAMA", "İSTENMEDİ"]
        
        alarm_listesi = []
        for p in all_projects:
            durum_upper = str(p.durum).strip().upper()
            if any(md in durum_upper for md in muaf_durumlar):
                continue
                
            son_tarih_str = ""
            son_puan = None
            if p.audits:
                tarihli_audits = [a for a in p.audits if a.tarih and len(str(a.tarih)) >= 10]
                if tarihli_audits:
                    tarihli_audits.sort(key=lambda x: str(x.tarih), reverse=True)
                    son_tarih_str = tarihli_audits[0].tarih
                    son_puan = tarihli_audits[0].puan
            
            if eski_gizle and son_tarih_str and (son_tarih_str.startswith("2024") or son_tarih_str.startswith("2025") or son_tarih_str.startswith("2023")):
                continue
            
            gecen_gun = 9999
            if son_tarih_str:
                try:
                    gecen_gun = (today - datetime.strptime(son_tarih_str[:10], "%Y-%m-%d").date()).days
                except:
                    pass
                
            alarm_sebebi, alarm_seviyesi = "", ""
            if son_puan is not None and son_puan <= 89.0:
                if gecen_gun > 30:
                    alarm_sebebi, alarm_seviyesi = f"🚨 ACİL: Puanı Düşük ({son_puan}), {gecen_gun} gündür gidilmedi!", "KIRMIZI"
                else:
                    alarm_sebebi, alarm_seviyesi = f"🟡 TAKİP: Puanı Düşük ({son_puan}). Tekrar denetimine {30-gecen_gun} gün kaldı.", "SARI"
            elif p.is_buyuk_proje:
                if gecen_gun > 180:
                    alarm_sebebi, alarm_seviyesi = f"⏰ UYARI: BÜYÜK PROJE ⭐ - En son {gecen_gun} gün önce gidildi!", "TURUNCU"
            elif gecen_gun > 365:
                alarm_sebebi, alarm_seviyesi = f"💤 UNUTULMUŞ: 1 Yıldan uzun süredir denetim yok!", "GRİ"
                
            if alarm_sebebi:
                alarm_listesi.append({
                    "Seviye": "🔴" if alarm_seviyesi=="KIRMIZI" else ("🟠" if alarm_seviyesi=="TURUNCU" else ("🟡" if alarm_seviyesi=="SARI" else "⚪")),
                    "Proje": f"{p.proje_adi} {'⭐' if p.is_buyuk_proje else ''}",
                    "Durum": p.durum,
                    "Operasyon Md.": p.opr_muduru,
                    "Son Tarih": son_tarih_str or "-",
                    "Puan": son_puan or "-",
                    "Kural": alarm_sebebi
                })
                
        if alarm_listesi:
            st.dataframe(pd.DataFrame(alarm_listesi), use_container_width=True, hide_index=True)
        else:
            st.success("Sahanızda geciken proje bulunmuyor.")

        st.markdown("---")
        st.subheader("🚨 Termin Tarihi Geçen Açık DÖF'ler (KIRMIZI ALARM)")
        
        geciken_doflar = []
        bugun_tarih = date.today()
        
        for d in all_dofs:
            if str(d.durum).upper() != "KAPANDI" and d.termin_tarihi:
                try:
                    termin = datetime.strptime(str(d.termin_tarihi).strip()[:10], "%Y-%m-%d").date()
                    gecen_gun = (bugun_tarih - termin).days
                    
                    if gecen_gun > 0:
                        p = d.project
                        geciken_doflar.append({
                            "Gecikme": f"🔴 {gecen_gun} Gün Gecikti",
                            "Proje": p.proje_adi if p else "-",
                            "Direktör": p.direktor if p else "-",
                            "Opr. Müdürü": p.opr_muduru if p else "-",
                            "DÖF No": d.dof_no,
                            "Uygunsuzluk": d.uygunsuzluk_tanimi,
                            "Sorumlu": d.sorumlu_kisi,
                            "Termin Tarihi": d.termin_tarihi
                        })
                except:
                    pass
                    
        if geciken_doflar:
            st.error(f"Dikkat! Kapatılması gereken tarihi geçmiş {len(geciken_doflar)} adet açık DÖF bulunuyor.")
            st.dataframe(pd.DataFrame(geciken_doflar).sort_values(by="Gecikme", ascending=False), use_container_width=True, hide_index=True)
        else:
            st.success("Tüm projelerde termin tarihi geçmiş açık DÖF bulunmamaktadır. Harika iş!")

# ==========================================
# SAYFA 3: GELİŞMİŞ PROJE HAVUZU
# ==========================================
elif sayfa == "📋 Gelişmiş Proje Havuzu":
    st.header("📋 Gelişmiş Proje ve Operasyon Havuzu")
    if all_projects:
        fc1, fc2, fc3, fc4 = st.columns(4)
        with fc1: f_ara = st.text_input("🔍 Ara...")
        with fc2: f_il = st.selectbox("📍 İl", ["Tümü"] + sorted(list(set([p.il for p in all_projects if p.il]))))
        with fc3: f_opr = st.selectbox("👔 Operasyon Müdürü", ["Tümü"] + sorted(list(set([p.opr_muduru for p in all_projects if p.opr_muduru]))))
        with fc4: f_durum = st.selectbox("📌 Durum", ["Tümü"] + sorted(list(set([p.durum for p in all_projects if p.durum]))))
            
        filtered = all_projects
        if f_ara:
            filtered = [p for p in filtered if f_ara.lower() in str(p.proje_adi).lower() or f_ara.lower() in str(p.proje_kodu).lower()]
        if f_il != "Tümü":
            filtered = [p for p in filtered if p.il == f_il]
        if f_opr != "Tümü":
            filtered = [p for p in filtered if p.opr_muduru == f_opr]
        if f_durum != "Tümü":
            filtered = [p for p in filtered if p.durum == f_durum]
        
        st.markdown(f"**Bulunan Sonuç:** `{len(filtered)}`")
        data_rows = []
        for p in filtered:
            data_rows.append({
                "Proje Kodu": p.proje_kodu,
                "Proje Adı": p.proje_adi,
                "Durum": p.durum,
                "İl": p.il,
                "Bölge Md.": p.bolge_md,
                "Operasyon Md.": p.opr_muduru,
                "Takım Lideri": p.takim_lideri,
                "Açık DÖF": len([d for d in p.dof_list if d.durum != "KAPANDI"]),
                "Harita GPS": "Var" if p.koordinat else "Yok",
                "Personel": p.per_sayisi
            })
        st.dataframe(pd.DataFrame(data_rows), use_container_width=True, hide_index=True)

# ==========================================
# SAYFA 4: PROJE DETAY, DÖF & 🗺️ HARİTA
# ==========================================
elif sayfa == "🎯 Proje Detay, DÖF & 🗺️ Harita":
    st.header("🎯 Proje 360° İnceleme, Harita ve DÖF Yönetimi")
    if all_projects:
        secilen_str = st.selectbox("🔍 Proje Seçin:", [f"{p.proje_kodu} | {p.proje_adi} {'⭐' if p.is_buyuk_proje else ''}" for p in all_projects])
        kodu = secilen_str.split(" | ")[0]
        p = db.query(ProjectDB).filter(ProjectDB.proje_kodu == kodu).first()
        
        if p:
            st.markdown("---")
            c1, c2 = st.columns([2, 1])
            with c1:
                st.subheader(f"🏢 {p.proje_adi} {'⭐ (BÜYÜK PROJE)' if p.is_buyuk_proje else ''}")
                st.markdown(f"**📍 Adres:** {p.adres} | **📍 İl:** {p.il} | **📌 Durum:** `{p.durum}`")
            with c2:
                if p.koordinat:
                    st.success("📍 **GPS Lokasyonu Sistemde Kayıtlı**")
                    maps_url = f"https://www.google.com/maps/search/?api=1&query={p.koordinat.replace(' ', '')}"
                    st.markdown(f"[🗺️ **Tıkla ve Google Haritalar'da Yol Tarifi Al**]({maps_url})")
                else:
                    st.warning("Bu projeye ait GPS koordinatı girilmemiştir.")
                    
            st.markdown("---")
            p_tab1, p_tab2, p_tab3 = st.tabs(["🗺️ Canlı Harita & Lokasyon", "📂 DÖF & Belgeler", "📊 Geçmiş Analizi"])
            
            with p_tab1:
                if p.koordinat:
                    try:
                        parts = p.koordinat.replace(";", ",").split(",")
                        if len(parts) >= 2:
                            lat = float(parts[0].strip())
                            lon = float(parts[1].strip())
                            map_df = pd.DataFrame({"lat": [lat], "lon": [lon]})
                            st.subheader(f"📍 {p.proje_adi} - Konum")
                            st.map(map_df, zoom=14, use_container_width=True)
                        else:
                            st.error("Koordinat formatı hatalı. Lütfen 'Enlem, Boylam' şeklinde yükleyin.")
                    except:
                        st.error(f"Koordinat okunurken bir hata oluştu: {p.koordinat}")
                else:
                    st.info("Haritayı görebilmek için yan menüdeki 'Excel & Koordinat Yükleme' kısmından koordinat listesini yükleyiniz.")
            
            with p_tab2:
                st.subheader("🚨 Açık DÖF (Uygunsuzluk) Kayıtları")
                if p.dof_list:
                    for dof in p.dof_list:
                        if dof.durum != "KAPANDI":
                            st.markdown(f"**🔴 {dof.dof_no}** | Bulgu: {dof.uygunsuzluk_tanimi} | Sorumlu: {dof.sorumlu_kisi}")
                else:
                    st.success("Açık DÖF bulunmuyor.")
                
                st.markdown("---")
                st.subheader("📂 Proje Kanıt & PDF Arşivi")
                with st.expander("➕ Belge Yükle", expanded=False):
                    dosya = st.file_uploader("Dosya Seç:", type=["pdf", "jpg", "png", "jpeg"], key="doc_up")
                    if st.button("🚀 Yükle") and dosya:
                        yeni_doc = DocumentDB(
                            project_id=p.id,
                            dosya_adi=dosya.name,
                            dosya_turu="PDF/Kanıt",
                            yukleme_tarihi=date.today().strftime("%Y-%m-%d"),
                            dosya_icerik=dosya.read()
                        )
                        db.add(yeni_doc)
                        db.commit()
                        st.success("Yüklendi!")
                        st.rerun()
                        
                if p.documents:
                    for doc in p.documents:
                        st.download_button(
                            label=f"📥 İndir: {doc.dosya_adi}",
                            data=doc.dosya_icerik,
                            file_name=doc.dosya_adi,
                            key=f"d_{doc.id}"
                        )

            # YENİ: PROFESYONEL GRAFİK (Plotly)
            with p_tab3:
                st.subheader("📈 Denetim Puanı Trendi")
                puanli = [a for a in p.audits if a.puan is not None]
                if puanli:
                    puanli.sort(key=lambda x: str(x.tarih) if x.tarih else str(x.id))
                    df_chart = pd.DataFrame({
                        "Tarih": [a.tarih for a in puanli],
                        "Puan": [a.puan for a in puanli],
                        "Denetim Türü": [a.denetim_turu for a in puanli],
                        "Denetçi": [a.gorevlisi for a in puanli]
                    })
                    
                    fig = px.line(
                        df_chart, 
                        x="Tarih", 
                        y="Puan", 
                        text="Puan",
                        markers=True,
                        title=f"{p.proje_adi} Puan Gelişimi",
                        hover_data=["Denetim Türü", "Denetçi"]
                    )
                    fig.update_traces(textposition="top center", line_color="#1f77b4", marker=dict(size=10, color="red"))
                    fig.update_layout(yaxis=dict(range=[0, 105]))
                    
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Bu proje için henüz puanlı bir denetim kaydı bulunmamaktadır.")

# ==========================================
# SAYFA 5: DÖF TAKİP VE KÖK NEDEN ANALİZİ
# ==========================================
elif sayfa == "🛠️ DÖF Takip & Kök Neden Analizi":
    st.header("🛠️ DÖF Takip & Kök Neden Analizi")
    
    if not all_dofs:
        st.warning("Sistemde henüz kayıtlı DÖF bulunmuyor.")
    else:
        st.subheader("📊 DÖF Analizleri (Makro Bakış)")
        c1, c2 = st.columns(2)
        
        dof_df = pd.DataFrame([{
            "Uygunsuzluk Alanı": d.uygunsuzluk_alan or "Belirtilmemiş",
            "Kök Neden": d.kok_neden or "Belirtilmemiş",
            "Durum": d.durum
        } for d in all_dofs])
        
        with c1:
            st.markdown("**En Çok Hata Veren Kök Nedenler**")
            kok_neden_counts = dof_df["Kök Neden"].value_counts()
            st.bar_chart(kok_neden_counts)
            
        with c2:
            st.markdown("**Alan Bazlı DÖF Dağılımı**")
            alan_counts = dof_df["Uygunsuzluk Alanı"].value_counts()
            st.bar_chart(alan_counts)
            
        st.markdown("---")
        st.subheader("📝 Açık DÖF Yönetim Paneli (İnteraktif)")
        st.info("💡 İpucu: Tablo üzerindeki Durum, Termin, Aksiyon ve Kapatma Notu hücrelerine çift tıklayarak verileri düzenleyebilir ve 'Değişiklikleri Kaydet' butonuna basabilirsiniz.")
        
        acik_doflar = [d for d in all_dofs if str(d.durum).upper() != "KAPANDI"]
        
        if not acik_doflar:
            st.success("🎉 Tebrikler! Sistemde açık DÖF bulunmamaktadır.")
        else:
            dof_listesi = []
            for d in acik_doflar:
                proje = db.query(ProjectDB).filter(ProjectDB.id == d.project_id).first()
                dof_listesi.append({
                    "DB_ID": d.id,
                    "Proje": proje.proje_adi if proje else "Bilinmiyor",
                    "DÖF No": d.dof_no,
                    "Uygunsuzluk": d.uygunsuzluk_tanimi,
                    "Sorumlu": d.sorumlu_kisi,
                    "Durum": d.durum,
                    "Termin Tarihi": d.termin_tarihi,
                    "Aksiyon Planı": d.aksiyon_plani,
                    "Kapatma Notu": d.kapatma_notu
                })
            
            df_edit = pd.DataFrame(dof_listesi)
            
            # String olarak gelen veriyi Date formatına çeviriyoruz.
            df_edit["Termin Tarihi"] = pd.to_datetime(df_edit["Termin Tarihi"], errors='coerce').dt.date
            
            edited_df = st.data_editor(
                df_edit,
                column_config={
                    "DB_ID": None,
                    "Proje": st.column_config.TextColumn(disabled=True),
                    "DÖF No": st.column_config.TextColumn(disabled=True),
                    "Uygunsuzluk": st.column_config.TextColumn(disabled=True),
                    "Sorumlu": st.column_config.TextColumn(disabled=True),
                    "Durum": st.column_config.SelectboxColumn(options=["AÇIK", "KAPANDI", "SÜREÇ DEVAM EDİYOR"]),
                    "Termin Tarihi": st.column_config.DateColumn(format="YYYY-MM-DD"),
                    "Aksiyon Planı": st.column_config.TextColumn(),
                    "Kapatma Notu": st.column_config.TextColumn()
                },
                hide_index=True,
                use_container_width=True,
                key="dof_editor"
            )
            
            if st.button("💾 Değişiklikleri Veritabanına Kaydet", type="primary"):
                degisen_sayisi = 0
                for index, row in edited_df.iterrows():
                    dof_id = row["DB_ID"]
                    guncel_durum = row["Durum"]
                    
                    # Veritabanına kaydederken Date objesini tekrar String'e çeviriyoruz.
                    if pd.notnull(row["Termin Tarihi"]):
                        guncel_termin = row["Termin Tarihi"].strftime("%Y-%m-%d")
                    else:
                        guncel_termin = ""
                        
                    guncel_aksiyon = row["Aksiyon Planı"]
                    guncel_kapatma = row["Kapatma Notu"]
                    
                    dof_kaydi = db.query(DofDB).filter(DofDB.id == dof_id).first()
                    if dof_kaydi:
                        if (dof_kaydi.durum != guncel_durum or 
                            dof_kaydi.termin_tarihi != guncel_termin or 
                            dof_kaydi.aksiyon_plani != guncel_aksiyon or 
                            dof_kaydi.kapatma_notu != guncel_kapatma):
                            
                            dof_kaydi.durum = guncel_durum
                            dof_kaydi.termin_tarihi = guncel_termin
                            dof_kaydi.aksiyon_plani = guncel_aksiyon
                            dof_kaydi.kapatma_notu = guncel_kapatma
                            degisen_sayisi += 1
                
                if degisen_sayisi > 0:
                    db.commit()
                    st.success(f"✅ {degisen_sayisi} adet DÖF başarıyla güncellendi!")
                    st.rerun()
                else:
                    st.info("Herhangi bir değişiklik yapılmadı.")

# ==========================================
# SAYFA 6: DENETİM, ZİYARET VE DÖF GİRİŞİ (YENİ)
# ==========================================
elif sayfa == "➕ Denetim & Ziyaret Girişi":
    st.header("➕ Yeni Denetim, Ziyaret ve DÖF Girişi")
    
    if not all_projects:
        st.warning("Lütfen önce sisteme proje yükleyin.")
    else:
        tab_denetim, tab_dof = st.tabs(["📝 Yeni Denetim / Ziyaret Formu", "🚨 Yeni DÖF (Uygunsuzluk) Formu"])
        
        # --- ZİYARET / DENETİM FORMU ---
        with tab_denetim:
            with st.form("form_denetim", clear_on_submit=True):
                st.subheader("Yeni Ziyaret / Denetim Kaydı Oluştur")
                
                d_proje_str = st.selectbox("Proje Seçin:", [f"{p.proje_kodu} | {p.proje_adi}" for p in all_projects])
                c1, c2 = st.columns(2)
                with c1:
                    d_tur = st.selectbox("Denetim/Ziyaret Türü", ["1. Denetim", "2. Denetim", "Tekrar Denetimi", "Dış Denetim", "Habersiz Ziyaret", "Opr. Müdürü Ziyareti", "Gece Denetimi"])
                    d_tarih = st.date_input("Denetim Tarihi")
                with c2:
                    d_puan = st.number_input("Puan (0-100)", min_value=0.0, max_value=100.0, value=0.0, step=0.1)
                    d_gor = st.text_input("Denetim Görevlisi (Ad Soyad)")
                    
                d_not = st.text_area("Ziyaret Notları / Gözlemler")
                btn_denetim = st.form_submit_button("💾 Denetimi Kaydet", type="primary")
                
                if btn_denetim:
                    p_kodu = d_proje_str.split(" | ")[0]
                    secili_p = db.query(ProjectDB).filter(ProjectDB.proje_kodu == p_kodu).first()
                    
                    yeni_denetim = AuditDB(
                        project_id=secili_p.id,
                        denetim_turu=d_tur,
                        tarih=str(d_tarih),
                        puan=d_puan if d_puan > 0 else None,
                        gorevlisi=d_gor,
                        notlar=d_not
                    )
                    db.add(yeni_denetim)
                    db.commit()
                    st.success(f"✅ '{d_tur}' kaydı {secili_p.proje_adi} projesi için başarıyla eklendi!")

        # --- DÖF AÇMA FORMU ---
        with tab_dof:
            with st.form("form_dof", clear_on_submit=True):
                st.subheader("Yeni DÖF (Düzenleyici ve Önleyici Faaliyet) Aç")
                
                dof_proje_str = st.selectbox("Proje Seçin (DÖF):", [f"{p.proje_kodu} | {p.proje_adi}" for p in all_projects])
                c3, c4 = st.columns(2)
                with c3:
                    dof_no = st.text_input("DÖF No (Örn: DOF-2026-001)")
                    dof_alan = st.selectbox("Uygunsuzluk Alanı", ["Personel", "Ekipman/Malzeme", "Evrak/Dokümantasyon", "İSG", "Müşteri Şikayeti", "Diğer"])
                with c4:
                    dof_neden = st.selectbox("Kök Neden", ["Eğitim Eksikliği", "Süreç Hatası", "Dikkatsizlik/İhmal", "Arıza/Yıpranma", "İletişim Kopukluğu", "Diğer"])
                    dof_sorumlu = st.text_input("Sorumlu Kişi")
                    
                dof_tanim = st.text_area("Uygunsuzluk Tanımı (Bulgu)")
                dof_termin = st.date_input("Termin Tarihi (Son Çözüm Tarihi)")
                
                btn_dof = st.form_submit_button("🚨 DÖF'ü Kaydet ve Aç", type="primary")
                
                if btn_dof:
                    if not dof_no or not dof_tanim:
                        st.error("Lütfen DÖF No ve Uygunsuzluk Tanımını doldurunuz.")
                    else:
                        p_kodu2 = dof_proje_str.split(" | ")[0]
                        secili_p2 = db.query(ProjectDB).filter(ProjectDB.proje_kodu == p_kodu2).first()
                        
                        yeni_dof = DofDB(
                            project_id=secili_p2.id,
                            dof_no=dof_no,
                            uygunsuzluk_alan=dof_alan,
                            kok_neden=dof_neden,
                            uygunsuzluk_tanimi=dof_tanim,
                            sorumlu_kisi=dof_sorumlu,
                            acilis_tarihi=str(date.today()),
                            termin_tarihi=str(dof_termin),
                            durum="AÇIK"
                        )
                        db.add(yeni_dof)
                        db.commit()
                        st.success(f"✅ {dof_no} numaralı DÖF başarıyla açıldı!")

# ==========================================
# SAYFA 7: DENETÇİ KARNESİ
# ==========================================
elif sayfa == "📈 Denetçi Performans Karnesi":
    st.header("📈 Denetçi Performans Karnesi")
    if all_audits:
        denetci_stats = {}
        for a in all_audits:
            gor = a.gorevlisi if a.gorevlisi else "Belirtilmemiş Denetçi"
            if gor not in denetci_stats:
                denetci_stats[gor] = {"Toplam Denetim": 0, "Puanlı Denetim": 0, "Top Puan": 0}
            denetci_stats[gor]["Toplam Denetim"] += 1
            if a.puan is not None and a.puan > 0:
                denetci_stats[gor]["Puanlı Denetim"] += 1
                denetci_stats[gor]["Top Puan"] += a.puan
                
        d_rows = []
        for k, v in denetci_stats.items():
            ort = (v["Top Puan"] / v["Puanlı Denetim"]) if v["Puanlı Denetim"] > 0 else 0
            d_rows.append({
                "Denetim Görevlisi": k,
                "Toplam İşlem": v["Toplam Denetim"],
                "Gerçek Puanlı Denetim": v["Puanlı Denetim"],
                "Ortalama Puan": f"{ort:.1f}"
            })
        st.dataframe(pd.DataFrame(d_rows).sort_values(by="Toplam İşlem", ascending=False), use_container_width=True, hide_index=True)

# ==========================================
# SAYFA 8: EXCEL VE KOORDİNAT YÜKLEME EKRANI (ŞİFRELİ)
# ==========================================
elif sayfa == "🔄 Excel & Koordinat Yükleme":
    st.header("🔄 Veritabanı, Excel ve Harita Senkronizasyonu")
    
    if "is_admin" not in st.session_state:
        st.session_state.is_admin = False

    if not st.session_state.is_admin:
        st.info("🔒 Bu sayfaya sadece sistem yöneticisi dosya yükleyebilir.")
        girilen_sifre = st.text_input("Yönetici Şifresi:", type="password")
        
        if st.button("Giriş Yap"):
            if girilen_sifre == "altug2707.": 
                st.session_state.is_admin = True
                st.rerun()
            else:
                st.error("❌ Hatalı şifre girdiniz!")

    if st.session_state.is_admin:
        col1, col2 = st.columns([4, 1])
        with col1:
            st.success("✅ Yönetici girişi başarılı. Dosya yükleyebilirsiniz.")
        with col2:
            if st.button("🚪 Çıkış Yap"):
                st.session_state.is_admin = False
                st.rerun()
                
        tab_ana, tab_koor = st.tabs(["📊 1. Ana Operasyon Excel'i Yükle", "📍 2. Harita Koordinat Excel'i Yükle"])
        
        with tab_ana:
            st.info("Ana projeleri, personelleri ve denetim puanlarını güncellemek için kullanılır. (DÖF'ler silinmez).")
            up_f = st.file_uploader("📂 Ana Excel Dosyasını Buraya Bırakın", type=["xlsx", "xls"], key="ana_excel")
            if up_f and st.button("🔄 Sistemi Senkronize Et", type="primary", key="btn_ana"):
                with st.spinner("İşleniyor..."):
                    tr, yp, gp, yd = sync_excel_without_losing_dofs(up_f)
                    st.success(f"🎉 Tamamlandı! Yeni Proje: {yp}, Güncellenen: {gp}.")
                    
        with tab_koor:
            st.success("🗺️ **Harita Entegrasyonu:** İçinde 'PROJE KODU' ve 'KOORDİNAT' sütunları olan 2. Excel dosyanızı yükleyin.")
            up_k = st.file_uploader("📍 Koordinat Excel Dosyasını Buraya Bırakın", type=["xlsx", "xls"], key="koor_excel")
            if up_k and st.button("🚀 Koordinatları Sisteme Göm", type="primary", key="btn_koor"):
                with st.spinner("Haritalar oluşturuluyor..."):
                    basari, mesaj = sync_coordinates(up_k)
                    if basari:
                        st.success(f"🎉 Harika! Toplam {mesaj} projenin koordinatı başarıyla sisteme gömüldü ve haritaları oluşturuldu!")
                    else:
                        st.error(f"⚠️ Bir hata oluştu: {mesaj}")
