
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

country_mapping = {
    "1": "ABD / Kanada",
    "7": "Rusya / Kazakistan",
    "20": "Mısır",
    "31": "Hollanda",
    "33": "Fransa",
    "34": "İspanya",
    "36": "Macaristan",
    "39": "İtalya",
    "40": "Romanya",
    "41": "İsviçre",
    "43": "Avusturya",
    "44": "Birleşik Krallık",
    "46": "İsveç",
    "49": "Almanya",
    "52": "Meksika",
    "55": "Brezilya",
    "60": "Malezya",
    "61": "Avustralya",
    "62": "Endonezya",
    "63": "Filipinler",
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
    "216": "Tunus",
    "218": "Libya",
    "222": "Moritanya",
    "229": "Benin",
    "235": "Çad",
    "237": "Kamerun",
    "249": "Sudan",
    "252": "Somali",
    "254": "Kenya",
    "255": "Tanzanya",
    "256": "Uganda",
    "264": "Namibya",
    "266": "Lesotho",
    "267": "Botsvana",
    "268": "Esvatini",
    "298": "Faroe Adaları",
    "351": "Portekiz",
    "352": "Lüksemburg",
    "358": "Finlandiya",
    "359": "Bulgaristan",
    "374": "Ermenistan",
    "380": "Ukrayna",
    "381": "Sırbistan",
    "385": "Hırvatistan",
    "386": "Slovenya",
    "387": "Bosna-Hersek",
    "389": "Kuzey Makedonya",
    "420": "Çekya",
    "421": "Slovakya",
    "423": "Lihtenştayn",
    "507": "Panama",
    "673": "Brunei",
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

turkey_city_codes = {
    "212", "216", "224", "232", "242", "246", "252", "256", "258",
    "262", "264", "266", "272", "274", "276", "282", "284", "286",
    "288", "312", "322", "324", "326", "342", "352", "362", "382",
    "388", "412", "414", "422", "424", "426", "432", "434", "436",
    "438", "442", "452", "454", "456", "462", "464", "466", "472",
    "474", "476", "478", "482", "484", "486", "488"
}

def clean_number(num):
    if not isinstance(num, str):
        return "", "Bilinmiyor", ""
    raw = num.strip()
    digits = re.sub(r"[^\d]", "", raw)

    # 00 ile başlayan numaraları normalize et
    if digits.startswith("00"):
        digits = digits[2:]

    # Çok kısa numaralar
    if len(digits) < 9:
        return digits, "Geçersiz Numara", ""

    # Ülke kodu eşleşmesi
    for length in range(4, 0, -1):
        code = digits[:length]
        if code in country_mapping:
            return digits, country_mapping[code], code

    # Türkiye GSM
    if digits.startswith("0") and len(digits) == 11 and digits[1] == "5":
        return "90" + digits[1:], "Türkiye", "90"
    if digits.startswith("5") and len(digits) == 10:
        return "90" + digits, "Türkiye", "90"

    # Türkiye sabit hat
    if digits.startswith("0") and len(digits) >= 10 and digits[1:4] in turkey_city_codes:
        return "90" + digits[1:], "Türkiye", "90"
    if digits.startswith("0850") and len(digits) == 11:
        return "90" + digits[1:], "Türkiye", "90"

    return digits, "Bilinmiyor", ""

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
