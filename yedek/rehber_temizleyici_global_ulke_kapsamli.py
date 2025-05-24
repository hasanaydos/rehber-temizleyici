
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Tüm dünyayı kapsayan ITU bazlı ülke kodu eşlemesi (örnek 50 ülke ile, tam liste uzatılabilir)
country_mapping = {
    "1": "Amerika Birleşik Devletleri / Kanada",
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
    "51": "Peru",
    "52": "Meksika",
    "53": "Küba",
    "54": "Arjantin",
    "55": "Brezilya",
    "56": "Şili",
    "57": "Kolombiya",
    "58": "Venezuela",
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
    "236": "Orta Afrika Cumhuriyeti",
    "237": "Kamerun",
    "238": "Yeşil Burun Adaları",
    "239": "Sao Tome ve Principe",
    "240": "Ekvator Ginesi",
    "241": "Gabon",
    "242": "Kongo",
    "243": "Demokratik Kongo Cumhuriyeti",
    "244": "Angola",
    "245": "Gine-Bissau",
    "246": "Diego Garcia",
    "248": "Seyşeller",
    "249": "Sudan",
    "250": "Ruanda",
    "251": "Etiyopya",
    "252": "Somali",
    "253": "Cibuti",
    "254": "Kenya",
    "255": "Tanzanya",
    "256": "Uganda",
    "257": "Burundi",
    "258": "Mozambik",
    "260": "Zambiya",
    "261": "Madagaskar",
    "263": "Zimbabve",
    "264": "Namibya",
    "265": "Malavi",
    "266": "Lesotho",
    "267": "Botsvana",
    "268": "Esvatini",
    "269": "Komorlar",
    "290": "Saint Helena",
    "291": "Eritre",
    "297": "Aruba",
    "298": "Faroe Adaları",
    "299": "Grönland",
    "350": "Cebelitarık",
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

def guess_country_code(number):
    for length in range(4, 0, -1):
        prefix = number[:length]
        if prefix in country_mapping:
            return prefix, country_mapping[prefix]
    return "", "Bilinmiyor"

def clean_number(num):
    if not isinstance(num, str):
        return "", "Bilinmiyor", ""
    raw = num.strip()
    digits = re.sub(r"[^\d]", "", raw)
    code, country = guess_country_code(digits)
    return digits, country, code

def process_file(file_path):
    df = pd.read_csv(file_path)

    name_fields = [
        "First Name", "Middle Name", "Last Name",
        "Phonetic First Name", "Phonetic Middle Name", "Phonetic Last Name",
        "Name Prefix", "Name Suffix", "Nickname", "File As",
        "Organization Name", "Organization Title", "Organization Department",
        "Birthday", "Notes"
    ]
    existing_name_fields = [col for col in name_fields if col in df.columns]
    df["Tam İsim"] = df[existing_name_fields].fillna("").agg(" ".join, axis=1).str.replace(r"\s+", " ", regex=True).str.strip()

    phone_columns = [col for col in df.columns if "Phone" in col and "Value" in col]

    expanded_rows = []
    for idx, row in df.iterrows():
        name = row["Tam İsim"]
        for col in phone_columns:
            if pd.notna(row[col]):
                raw = str(row[col])
                parts = re.split(r"[:;,]", raw)
                for number in parts:
                    expanded_rows.append({
                        "İsim": name,
                        "Ham Numara": number.strip()
                    })

    expanded_df = pd.DataFrame(expanded_rows)
    expanded_df[["Temiz Numara", "Ülke Adı", "Ülke Kodu"]] = expanded_df["Ham Numara"].apply(
        lambda x: pd.Series(clean_number(x))
    )

    final_df = expanded_df.drop_duplicates(subset=["Temiz Numara"]).dropna(subset=["Temiz Numara"])
    final_df = final_df[final_df["Temiz Numara"] != ""]
    export_df = final_df[["Temiz Numara", "Ülke Adı", "Ülke Kodu", "İsim"]]

    output_path = os.path.join(os.path.dirname(file_path), "temizlenmis_rehber.xlsx")
    export_df.to_excel(output_path, index=False)
    return output_path

def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Google Rehber CSV Dosyasını Seçin", filetypes=[("CSV Dosyası", "*.csv")])
    if file_path:
        try:
            result_path = process_file(file_path)
            messagebox.showinfo("Başarılı", f"Temiz rehber başarıyla oluşturuldu:\n{result_path}")
        except Exception as e:
            messagebox.showerror("Hata", f"Bir hata oluştu:\n{str(e)}")
    else:
        messagebox.showinfo("İptal", "Dosya seçilmedi.")

if __name__ == "__main__":
    main()
