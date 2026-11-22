def sinif_belirle(onceki_pay, yillik_tuketim_degeri):
    if yillik_tuketim_degeri == 0:
        return "C"
    if onceki_pay < 80:
        return "A"
    elif onceki_pay < 95:
        return "B"
    else:
        return "C"

def abc_analizi(tablo):
    tablo["yillik_tuketim_degeri"] = tablo["yillik_tuketim"] * tablo["birim_fiyat"]
    abc_tablo = tablo.sort_values("yillik_tuketim_degeri", ascending=False)
    toplam_yillik_deger = abc_tablo["yillik_tuketim_degeri"].sum()
    if toplam_yillik_deger == 0:
        abc_tablo["ABC_sinifi"] = "C"
        abc_tablo["yillik_deger_payi"] = 0
        abc_tablo["kumulatif_pay"] = 0
        abc_tablo["onceki_pay"] = 0
        return abc_tablo

    abc_tablo["yillik_deger_payi"] = abc_tablo["yillik_tuketim_degeri"] / toplam_yillik_deger * 100
    abc_tablo["kumulatif_pay"] = abc_tablo["yillik_deger_payi"].cumsum()
    abc_tablo["onceki_pay"] = abc_tablo["kumulatif_pay"].shift(1).fillna(0)
    abc_tablo["ABC_sinifi"] = abc_tablo.apply(
    lambda satir: sinif_belirle(satir["onceki_pay"], satir["yillik_tuketim_degeri"]),
    axis=1
)
    return abc_tablo
