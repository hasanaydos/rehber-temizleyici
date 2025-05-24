
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def clean_number(num):
    if not isinstance(num, str):
        return "", "Bilinmiyor"
    raw = num.strip()
    num_digits = re.sub(r"[^\d]", "", raw)
    country = "Bilinmiyor"
    if raw.startswith("+90") or raw.startswith("0090") or re.match(r"0\s*5\d{2}", raw):
        country = "TR"
        if num_digits.startswith("0"):
            num_digits = "90" + num_digits[1:]
        elif not num_digits.startswith("90"):
            num_digits = "90" + num_digits
    elif raw.startswith("+49") or raw.startswith("0049"):
        country = "DE"
        num_digits = "49" + num_digits[-10:]
    elif raw.startswith("+966") or raw.startswith("00966"):
        country = "SA"
        num_digits = "966" + num_digits[-9:]
    elif raw.startswith("+"):
        country = "INT"
    elif num_digits.startswith("5") and len(num_digits) >= 10:
        country = "TR"
        num_digits = "90" + num_digits
    elif num_digits.startswith("0") and len(num_digits) >= 10:
        country = "TR"
        num_digits = "90" + num_digits[1:]
    return num_digits, country

def process_file(file_path):
    df = pd.read_csv(file_path)

    # Telefon sütunlarını tespit et
    phone_columns = [col for col in df.columns if "Phone" in col and "Value" in col]

    # Telefon sütunundan önceki tüm sütunları isim olarak kabul et
    first_phone_index = df.columns.get_loc(phone_columns[0])
    name_columns = df.columns[:first_phone_index]

    # Tüm isim sütunlarını birleştir
    df["Tam İsim"] = df[name_columns].fillna("").agg(" ".join, axis=1).str.replace(r"\s+", " ", regex=True).str.strip()

    # Numaraları ayrıştır
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
    expanded_df[["Temiz Numara", "Ülke"]] = expanded_df["Ham Numara"].apply(
        lambda x: pd.Series(clean_number(x))
    )

    final_df = expanded_df.drop_duplicates(subset=["Temiz Numara"]).dropna(subset=["Temiz Numara"])
    final_df = final_df[final_df["Temiz Numara"] != ""]
    export_df = final_df[["Temiz Numara", "Ülke", "İsim"]]

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
