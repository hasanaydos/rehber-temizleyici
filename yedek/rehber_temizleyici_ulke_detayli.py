
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Ülke kodları sözlüğü
country_mapping = {
    "90": "Türkiye",
    "49": "Almanya",
    "966": "Suudi Arabistan",
    "44": "Birleşik Krallık",
    "1": "Amerika Birleşik Devletleri",
    "33": "Fransa",
    "7": "Rusya",
    "39": "İtalya",
    "34": "İspanya",
    "971": "Birleşik Arap Emirlikleri",
    "61": "Avustralya",
    "92": "Pakistan",
    "20": "Mısır"
}

def clean_number(num):
    if not isinstance(num, str):
        return "", "", ""
    raw = num.strip()
    num_digits = re.sub(r"[^\d]", "", raw)
    country_code = ""
    country_name = "Bilinmiyor"

    if raw.startswith("+90") or raw.startswith("0090") or re.match(r"0\s*5\d{2}", raw):
        country_code = "90"
    elif raw.startswith("+49") or raw.startswith("0049"):
        country_code = "49"
    elif raw.startswith("+966") or raw.startswith("00966"):
        country_code = "966"
    elif raw.startswith("+971") or raw.startswith("00971"):
        country_code = "971"
    elif raw.startswith("+1") or raw.startswith("001"):
        country_code = "1"
    elif raw.startswith("+44") or raw.startswith("0044"):
        country_code = "44"
    elif raw.startswith("+33") or raw.startswith("0033"):
        country_code = "33"
    elif raw.startswith("+7") or raw.startswith("007"):
        country_code = "7"
    elif raw.startswith("+39") or raw.startswith("0039"):
        country_code = "39"
    elif raw.startswith("+34") or raw.startswith("0034"):
        country_code = "34"
    elif raw.startswith("+92") or raw.startswith("0092"):
        country_code = "92"
    elif raw.startswith("+20") or raw.startswith("0020"):
        country_code = "20"
    elif num_digits.startswith("0") and len(num_digits) >= 10:
        country_code = "90"
        num_digits = num_digits[1:]
    elif num_digits.startswith("5") and len(num_digits) >= 10:
        country_code = "90"

    if country_code in country_mapping:
        country_name = country_mapping[country_code]

    if country_code and not num_digits.startswith(country_code):
        num_digits = country_code + num_digits[-10:]

    return num_digits, country_name, country_code

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
