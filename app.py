import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# 🌍 ITU-T E.164 standardına göre dünya ülke kodları (kapsamlı liste)
country_mapping = {
    "1": "ABD / Kanada",
    "7": "Rusya / Kazakistan",
    "20": "Mısır",
    "27": "Güney Afrika",
    "30": "Yunanistan",
    "31": "Hollanda",
    "32": "Belçika",
    "33": "Fransa",
    "34": "İspanya",
    "36": "Macaristan",
    "39": "İtalya",
    "40": "Romanya",
    "41": "İsviçre",
    "43": "Avusturya",
    "44": "Birleşik Krallık",
    "45": "Danimarka",
    "46": "İsveç",
    "47": "Norveç",
    "48": "Polonya",
    "49": "Almanya",
    "52": "Meksika",
    "55": "Brezilya",
    "60": "Malezya",
    "61": "Avustralya",
    "62": "Endonezya",
    "63": "Filipinler",
    "64": "Yeni Zelanda",
    "65": "Singapur",
    "66": "Tayland",
    "81": "Japonya",
    "82": "Güney Kore",
    "84": "Vietnam",
    "86": "Çin",
    "90": "Türkiye",
    "91": "Hindistan",
    "92": "Pakistan",
    "93": "Afganistan",
    "94": "Sri Lanka",
    "95": "Myanmar",
    "98": "İran",
    "211": "Güney Sudan",
    "212": "Fas",
    "213": "Cezayir",
    "216": "Tunus",
    "218": "Libya",
    "220": "Gambiya",
    "221": "Senegal",
    "222": "Moritanya",
    "223": "Mali",
    "224": "Gine",
    "225": "Fildişi Sahili",
    "226": "Burkina Faso",
    "227": "Nijer",
    "228": "Togo",
    "229": "Benin",
    "230": "Mauritius",
    "231": "Liberya",
    "232": "Sierra Leone",
    "233": "Gana",
    "234": "Nijerya",
    "235": "Çad",
    "237": "Kamerun",
    "249": "Sudan",
    "251": "Etiyopya",
    "252": "Somali",
    "254": "Kenya",
    "255": "Tanzanya",
    "256": "Uganda",
    "260": "Zambiya",
    "261": "Madagaskar",
    "263": "Zimbabve",
    "265": "Malavi",
    "266": "Lesotho",
    "267": "Botsvana",
    "268": "Esvatini",
    "298": "Faroe Adaları",
    "351": "Portekiz",
    "352": "Lüksemburg",
    "353": "İrlanda",
    "354": "İzlanda",
    "355": "Arnavutluk",
    "356": "Malta",
    "357": "Kıbrıs",
    "358": "Finlandiya",
    "359": "Bulgaristan",
    "370": "Litvanya",
    "371": "Letonya",
    "372": "Estonya",
    "373": "Moldova",
    "374": "Ermenistan",
    "375": "Belarus",
    "376": "Andorra",
    "377": "Monako",
    "378": "San Marino",
    "380": "Ukrayna",
    "381": "Sırbistan",
    "382": "Karadağ",
    "385": "Hırvatistan",
    "386": "Slovenya",
    "387": "Bosna-Hersek",
    "389": "Kuzey Makedonya",
    "420": "Çekya",
    "421": "Slovakya",
    "423": "Lihtenştayn",
    "850": "Kuzey Kore",
    "852": "Hong Kong",
    "855": "Kamboçya",
    "880": "Bangladeş",
    "961": "Lübnan",
    "962": "Ürdün",
    "963": "Suriye",
    "964": "Irak",
    "965": "Kuveyt",
    "966": "Suudi Arabistan",
    "967": "Yemen",
    "968": "Umman",
    "970": "Filistin",
    "971": "Birleşik Arap Emirlikleri",
    "972": "İsrail",
    "973": "Bahreyn",
    "974": "Katar",
    "975": "Bhutan",
    "976": "Moğolistan",
    "977": "Nepal",
    "992": "Tacikistan",
    "993": "Türkmenistan",
    "994": "Azerbaycan",
    "995": "Gürcistan",
    "996": "Kırgızistan",
    "998": "Özbekistan"
}

# 🇹🇷 Türkiye'deki şehir alan kodları
turkey_city_codes = {
    "212", "216", "224", "232", "242", "246", "252", "256", "258", "262",
    "264", "266", "272", "274", "276", "282", "284", "286", "288", "312",
    "322", "324", "326", "342", "352", "362", "382", "388", "412", "414",
    "422", "424", "426", "432", "434", "436", "438", "442", "452", "454",
    "456", "462", "464", "466", "472", "474", "476", "478", "482", "484",
    "486", "488"
}

# 🧠 Numara temizleme ve ülke tanıma
def clean_number(num):
    if not isinstance(num, str): return "", "Bilinmiyor", ""
    raw = num.strip()
    digits = re.sub(r"[^\d]", "", raw)
    if digits.startswith("00"): digits = digits[2:]
    if len(digits) < 9: return digits, "Geçersiz Numara", ""
    for l in range(4, 0, -1):
        code = digits[:l]
        if code in country_mapping:
            return digits, country_mapping[code], code
    if digits.startswith("0") and len(digits) == 11 and digits[1] == "5":
        return "90" + digits[1:], "Türkiye", "90"
    if digits.startswith("5") and len(digits) == 10:
        return "90" + digits, "Türkiye", "90"
    if digits.startswith("0") and digits[1:4] in turkey_city_codes:
        return "90" + digits[1:], "Türkiye", "90"
    if digits.startswith("0850") and len(digits) == 11:
        return "90" + digits[1:], "Türkiye", "90"
    return digits, "Bilinmiyor", ""

# 📥 CSV'den oku, 📤 Excel'e yaz
def process_file(file_path):
    df = pd.read_csv(file_path)
    name_fields = ["First Name", "Middle Name", "Last Name", "Phonetic First Name", "Phonetic Middle Name", "Phonetic Last Name", "Name Prefix", "Name Suffix", "Nickname", "File As", "Organization Name", "Organization Title", "Organization Department", "Birthday", "Notes"]
    df["Tam İsim"] = df[[c for c in name_fields if c in df.columns]].fillna("").agg(" ".join, axis=1).str.replace(r"\s+", " ", regex=True).str.strip()
    phone_cols = [c for c in df.columns if "Phone" in c and "Value" in c]
    rows = []
    for _, row in df.iterrows():
        for col in phone_cols:
            if pd.notna(row[col]):
                for num in re.split(r"[:;,]", str(row[col])):
                    rows.append({"İsim": row["Tam İsim"], "Ham Numara": num.strip()})
    out = pd.DataFrame(rows)
    out[["Temiz Numara", "Ülke Adı", "Ülke Kodu"]] = out["Ham Numara"].apply(lambda x: pd.Series(clean_number(x)))
    out = out[out["Temiz Numara"] != ""].drop_duplicates(subset=["Temiz Numara"])
    out[["Temiz Numara", "Ülke Adı", "Ülke Kodu", "İsim"]].to_excel(
        os.path.join(os.path.dirname(file_path), "temizlenmis_rehber.xlsx"), index=False
    )

# 📂 Dosya seçimi ve başlatma
def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Google Rehber CSV Dosyasını Seçin", filetypes=[("CSV Dosyası", "*.csv")])
    if file_path:
        try:
            process_file(file_path)
            messagebox.showinfo("Başarılı", "Temiz rehber başarıyla oluşturuldu.")
        except Exception as e:
            messagebox.showerror("Hata", str(e))
    else:
        messagebox.showinfo("İptal", "Dosya seçilmedi.")

if __name__ == "__main__":
    main()
