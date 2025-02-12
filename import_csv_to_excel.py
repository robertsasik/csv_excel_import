#pip install pandas
#pip install openpyxl

import os
import pandas as pd

# Nastav priečinok, kde sú CSV súbory
csv_folder = "CSV/"  # Zmeň na svoj priečinok

output_excel = "Vystup.xlsx"  # Výstupný Excel súbor

# Zoznam CSV súborov v priečinku
csv_files = [f for f in os.listdir(csv_folder) if f.endswith(".csv")]

# Vytvor nový Excel writer
with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    for file in csv_files:
        file_path = os.path.join(csv_folder, file)
        
        # Načítaj CSV do DataFrame
        df = pd.read_csv(file_path)
        
        # Získaj názov súboru bez prípony ako názov záložky
        sheet_name = os.path.splitext(file)[0][:31]  # Excel podporuje max 31 znakov v názve záložky
        
        # Zapíš DataFrame do novej záložky
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Všetky CSV boli importované do {output_excel}!")

print("✅ Hotovo! Každý CSV súbor je v Exceli ako samostatná záložka.")
