import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from pathlib import Path
from analizler import abc_analizi
from openpyxl.utils import get_column_letter
st.markdown("""
<style>

/* Dört KPI kartının ortak tasarımı */
.st-key-kpi_stok,
.st-key-kpi_kritik,
.st-key-kpi_eksik,
.st-key-kpi_urun {
    border: 1px solid #D5E5E1;
    border-radius: 14px;
}

/* Toplam Stok Değeri - Pastel Mavi */
.st-key-kpi_stok {
    background-color: #D9ECF9;
}

/* Kritik Ürün Sayısı - Pastel Kırmızı */
.st-key-kpi_kritik {
    background-color: #F8DFDC;
}

/* Toplam Eksik Stok Tutarı - Pastel Bej */
.st-key-kpi_eksik {
    background-color: #F7E9D4;
}

/* Toplam Ürün Sayısı - Pastel Yeşil */
.st-key-kpi_urun {
    background-color: #DCF0E5;
}

</style>
""", unsafe_allow_html=True)

zorunlu_sutunlar = ["ad", "stok", "birim_fiyat", "kritik_stok", "yillik_tuketim"]
sablon_klasoru = Path(__file__).resolve().parent / "sablonlar"
ornek_veri_alani = st.container()
with st.expander("📂 Stok Dosyası Yükle", expanded=False):
    st.caption(
        "Boş şablonu indirip başlıkların altına her satıra bir ürün yazın. "
        "Başlıkları ve yerlerini koruyun. CSV için virgül ayırıcı ve ondalık nokta kullanın."
    )
    excel_sutunu, csv_sutunu = st.columns(2)
    for sutun, etiket, uzanti, mime in (
        (excel_sutunu, "Excel Şablonu İndir", "xlsx",
         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        (csv_sutunu, "CSV Şablonu İndir", "csv", "text/csv"),
    ):
        dosya_adi = f"matmis_stok_sablonu.{uzanti}"
        sutun.download_button(
            etiket,
            data=(sablon_klasoru / dosya_adi).read_bytes(),
            file_name=dosya_adi,
            mime=mime,
            on_click="ignore",
        )
    yuklenen_dosya = st.file_uploader("Stok Dosyanızı Yükleyiniz", type=["csv", "xlsx"])
if yuklenen_dosya is None:
    with ornek_veri_alani:
        if st.session_state.get("ornek_veri_acik", False):
            st.caption("35 ürünlük örnek veri seti gösteriliyor.")
            st.button("Örnek veriyi kaldır", on_click=st.session_state.update,
                      args=({"ornek_veri_acik": False},))
        else:
            st.button("Örnek veri setiyle deneyin", on_click=st.session_state.update,
                      args=({"ornek_veri_acik": True},))
    if st.session_state.get("ornek_veri_acik", False):
        tablo = pd.read_csv(sablon_klasoru / "matmis_ornek_stok.csv")
    else:
        tablo = pd.DataFrame(columns=zorunlu_sutunlar)
else:
    st.session_state.pop("ornek_veri_acik", None)
    print("Stok Dosyası Yüklendi")
    if yuklenen_dosya.name.endswith(".csv"):
        print("Bu bir csv dosyası")
        tablo = pd.read_csv(yuklenen_dosya)
    else:
        print("Bu bir csv dosyası degil")
        tablo = pd.read_excel(yuklenen_dosya, header=1)
        print(tablo.columns.tolist())
        print(tablo.head())
for sutun in zorunlu_sutunlar:
    if sutun not in tablo.columns:
        st.error(f"Eksik sütun: {sutun}")
        st.stop()
if yuklenen_dosya is not None and tablo.empty:
    st.error("Stok tablosu boş. Lütfen ürün bilgilerini içeren bir dosya yükleyiniz.")
    st.stop()
sayisal_sutunlar = ["stok", "birim_fiyat", "kritik_stok", "yillik_tuketim"]    
for sutun in sayisal_sutunlar:
    if pd.to_numeric(tablo[sutun], errors="coerce").isna().any(): 
        st.error(f"Geçersiz sayısal değer: {sutun}")
        st.stop()
    tablo[sutun] = pd.to_numeric(tablo[sutun], errors="coerce")
    if (tablo[sutun] < 0).any():
        st.error(f"Tabloda negatif değer tespit edildi: {sutun}")
        st.stop()
abc_tablo = abc_analizi(tablo)
print(abc_tablo.columns.tolist())
toplam_yillik_deger = abc_tablo["yillik_tuketim_degeri"].sum()
tablo["stok_degeri"] = tablo["stok"] * tablo["birim_fiyat"]
toplam_stok_degeri = tablo["stok_degeri"].sum()
tablo["eksik_miktar"] = (tablo["kritik_stok"] - tablo["stok"]).clip(lower=0)
tablo["eksik_tutar"] = tablo["eksik_miktar"] * tablo["birim_fiyat"]
toplam_eksik_tutar = tablo["eksik_tutar"].sum()
eksik_tutar_metin = f"{toplam_eksik_tutar:,.2f}"
duzenlenmis_eksik_tutar = eksik_tutar_metin.replace(",","X").replace(".",",").replace("X",".")
sayi = f"{toplam_stok_degeri:,.2f}"
duzenlenmis_sayi = sayi.replace(",","X").replace(".",",").replace("X",".")
kritik_urunler = tablo[tablo["stok"]<=tablo["kritik_stok"]]
kritik_urun_sayisi = len(kritik_urunler)
toplam_urun_sayisi = tablo["ad"].nunique()
excel_dosyasi = BytesIO()
with pd.ExcelWriter(excel_dosyasi) as writer:
    tablo.to_excel(writer, sheet_name="Stok Envanteri", index=False)
    kritik_urunler.to_excel(writer, sheet_name="Kritik Stoklar", index=False)
    abc_tablo.to_excel(writer, sheet_name="ABC Analizi", index=False)
    for sayfa in writer.sheets.values():
        for sutun in sayfa.columns:
            en_uzun = 0
            for hucre in sutun:
                en_uzun = max(en_uzun, len(str(hucre.value)))
            harf = get_column_letter(sutun[0].column)
            sayfa.column_dimensions[harf].width = en_uzun + 3

excel_verisi = excel_dosyasi.getvalue()    
abc_grafik = px.bar(abc_tablo, x="ad" , y="yillik_tuketim_degeri")
abc_grafik.update_traces(marker_color="#5C9FC5")
abc_grafik.add_trace(go.Scatter(
    line=dict(color="#20384A", width=3),
    x=abc_tablo["ad"],
    y=abc_tablo["kumulatif_pay"],
    mode="lines+markers",
    name="Kümülatif Pay",
    yaxis="y2"
))    
abc_grafik.update_layout(
    yaxis2=dict(title="Kümülatif Pay (%)", side="right", overlaying="y", range=[0, 110], dtick=20),
    yaxis_title=("Yıllık Tüketim Değeri (TL)"),
    xaxis_title=("Ürün Adı"),
    xaxis=dict(tickmode="auto", nticks=8),
    legend=dict(
    orientation="h",
    yanchor="bottom",
    y=1.08,
    xanchor="right",
    x=1
),
margin=dict(t=110),
    title="ABC Analizi | Pareto Grafiği"
    )
abc_grafik.add_shape(
    type="line", xref="paper", yref="y2", x0=0, x1=1, y0=80, y1=80, line=dict(color="#209074", width=2, dash="dash")
)
abc_grafik.add_shape(
    type="line", xref="paper", yref="y2", x0=0, x1=1, y0=95, y1=95, line=dict(color="#9493F6", width=2, dash="dash")
)
st.sidebar.title("MATMIS")
st.sidebar.subheader("Materials Management Information System")
secim = st.sidebar.radio("Menü", ["Genel Bakış", "Stok Analizi", "Kritik Stoklar", "ABC Analizi"])
if secim == "Genel Bakış":
    st.header("Genel Bakış")
    st.write("MATMIS, stokların ekonomik değerini ve kritik stok durumlarını analiz eden bir yönetim aracıdır.")
    sutun1, sutun2 = st.columns(2)
    sutun3, sutun4 = st.columns(2)
    with sutun1:
        with st.container(border=True, key="kpi_stok"):
            st.metric("Toplam Stok Değeri", f"{duzenlenmis_sayi} TL")
    with sutun2:
        with st.container(border=True, key="kpi_kritik"):
            st.metric("Kritik Ürün Sayısı:", kritik_urun_sayisi)
    with sutun3:
        with st.container(border=True, key="kpi_eksik"):
            st.metric("Toplam Eksik Stok Tutarı:", f"{duzenlenmis_eksik_tutar} TL")
    with sutun4:
        with st.container(border=True, key="kpi_urun"):
            st.metric("Toplam Ürün Sayısı:", toplam_urun_sayisi)   
    st.download_button(
    label="Excel Raporunu İndir",
    data=excel_verisi,
    file_name="MATMIS_Stok_Analiz_Raporu.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    disabled=tablo.empty,
)
elif tablo.empty:
    st.info("Analizleri görmek için stok dosyanızı yükleyin veya örnek veri setini açın.")
elif secim == "Stok Analizi":
    st.subheader("Stok Envanteri")
    secilen_urunler = st.multiselect("Ürünleri Seçiniz",
    tablo["ad"],
    default = tablo["ad"]
    )
    filtreli_tablo = tablo[tablo["ad"].isin(secilen_urunler)]
    st.dataframe(
    filtreli_tablo,
    hide_index=True,
    column_order=[
    "ad",
    "stok",
    "birim_fiyat",
    "stok_degeri",
    "kritik_stok",
    "eksik_miktar",
    "eksik_tutar",
    "yillik_tuketim",
    "yillik_tuketim_degeri"],
    column_config={
        "ad": "Ürün Adı",
        "stok": "Stok Miktarı",
        "birim_fiyat": st.column_config.NumberColumn(
            "Birim Fiyat (TL)", format="%.2f"
        ),
        "kritik_stok": "Kritik Stok",
        "yillik_tuketim": "Yıllık Tüketim",
        "yillik_tuketim_degeri": st.column_config.NumberColumn(
            "Yıllık Tüketim Değeri (TL)", format="%.2f"
        ),
        "stok_degeri": st.column_config.NumberColumn(
            "Stok Değeri (TL)", format="%.2f"
        ),
        "eksik_miktar": "Eksik Miktar",
        "eksik_tutar": st.column_config.NumberColumn(
            "Eksik Tutar (TL)", format="%.2f"
        )
    }
)
elif secim == "Kritik Stoklar":
    st.subheader("Kritik Stok Seviyesindeki Ürünler")
    st.dataframe(
    kritik_urunler,
    hide_index=True,
    column_order=["ad",
    "stok",
    "kritik_stok",
    "eksik_miktar",
    "eksik_tutar",
    "birim_fiyat",
    "stok_degeri",
    "yillik_tuketim",
    "yillik_tuketim_degeri"],
    column_config={
        "ad": "Ürün Adı",
        "stok": "Stok Miktarı",
        "birim_fiyat": st.column_config.NumberColumn(
            "Birim Fiyat (TL)", format="%.2f"
        ),
        "kritik_stok": "Kritik Stok",
        "yillik_tuketim": "Yıllık Tüketim",
        "yillik_tuketim_degeri": st.column_config.NumberColumn(
            "Yıllık Tüketim Değeri (TL)", format="%.2f"
        ),
        "stok_degeri": st.column_config.NumberColumn(
            "Stok Değeri (TL)", format="%.2f"
        ),
        "eksik_miktar": "Eksik Miktar",
        "eksik_tutar": st.column_config.NumberColumn(
            "Eksik Tutar (TL)", format="%.2f"
        )
    }
)
elif secim == "ABC Analizi":
    st.subheader("ABC Analiz Tablosu")
    st.dataframe(
    abc_tablo,
    hide_index=True,
    column_order=[
    "ad",
    "ABC_sinifi",
    "yillik_tuketim_degeri",
    "yillik_deger_payi",
    "kumulatif_pay",
    "onceki_pay",
    "yillik_tuketim",
    "stok",
    "birim_fiyat",
    "kritik_stok"],
    column_config={
        "ad": "Ürün Adı",
        "stok": "Stok Miktarı",
        "birim_fiyat": st.column_config.NumberColumn(
        "Birim Fiyat (TL)",
        format="%.2f"
    ),
        "yillik_tuketim_degeri": st.column_config.NumberColumn(
        "Yıllık Tüketim Değeri (TL)",
        format="%.2f"
    ),
        "kritik_stok": "Kritik Stok",
        "yillik_tuketim": "Yıllık Tüketim",
        "yillik_deger_payi": st.column_config.NumberColumn(
        "Yıllık Değer Payı (%)",
        format="%.2f"
    ),
    "kumulatif_pay": st.column_config.NumberColumn(
    "Kümülatif Pay (%)",
    format="%.2f"
    ),
    "onceki_pay": st.column_config.NumberColumn(
    "Önceki Pay (%)",
    format="%.2f",
        ),
    "ABC_sinifi": "ABC Sınıfı"
    })   
    if toplam_yillik_deger > 0:
        st.plotly_chart(abc_grafik)
    else:
        st.warning("Toplam yıllık tüketim değeri 0 olduğundan Pareto analizi oluşturulamadı. Ürünler varsayılan olarak C sınıfına atanmıştır.")
    