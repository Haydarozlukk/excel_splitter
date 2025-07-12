import os
import pandas as pd

def bol_excel_dosyalari(giris_klasoru, satir_sayisi=400_000):
    if not os.path.isdir(giris_klasoru):
        print("Geçersiz klasör yolu.")
        return

    cikis_klasoru = os.path.join(giris_klasoru, "bolunmus_exceller")
    os.makedirs(cikis_klasoru, exist_ok=True)

    dosyalar = [f for f in os.listdir(giris_klasoru) if f.endswith(".xlsx")]

    if not dosyalar:
        print("Belirtilen klasörde .xlsx uzantılı dosya bulunamadı.")
        return

    for dosya in dosyalar:
        dosya_yolu = os.path.join(giris_klasoru, dosya)
        print(f"İşleniyor: {dosya}")

        try:
            df = pd.read_excel(dosya_yolu, engine="openpyxl")
        except Exception as e:
            print(f"Hata oluştu: {dosya} dosyası okunamadı -> {e}")
            continue

        toplam = len(df)
        for i in range(0, toplam, satir_sayisi):
            parca = df.iloc[i:i + satir_sayisi]
            parca_no = i // satir_sayisi + 1
            yeni_ad = f"{os.path.splitext(dosya)[0]}_parca_{parca_no}.xlsx"
            yeni_yol = os.path.join(cikis_klasoru, yeni_ad)
            try:
                parca.to_excel(yeni_yol, index=False)
                print(f"Kaydedildi: {yeni_ad} ({len(parca)} satır)")
            except Exception as e:
                print(f"Parça kaydedilemedi: {e}")

    print("Tüm dosyalar başarıyla bölündü.")

if __name__ == "__main__":
    giris_yolu = input("Excel dosyalarının bulunduğu klasör yolunu girin: ").strip('"')
    bol_excel_dosyalari(giris_yolu)
