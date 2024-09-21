import xlwings as xw
import re

def save_excel_links_to_txt(file_path, output_txt_path):
    # Buka workbook menggunakan xlwings
    wb = xw.Book(file_path)

    # Pola untuk mendeteksi referensi ke file Excel eksternal
    file_reference_pattern = re.compile(r'\[.*\.xlsx\]|\[.*\.xlsm\]')

    # Buka file teks untuk menyimpan output
    with open(output_txt_path, 'w') as output_file:
        # Iterasi tiap sheet
        for sheet in wb.sheets:
            sheet_has_formula = False  # Flag untuk memastikan sheet hanya ditulis jika ada formula

            # Iterasi setiap cell di sheet
            for cell in sheet.used_range:
                # Cek jika cell mengandung formula
                if cell.formula:
                    # Cek apakah formula mengandung referensi ke file Excel eksternal
                    if file_reference_pattern.search(cell.formula):
                        if not sheet_has_formula:
                            output_file.write(f"Sheet: {sheet.name}\n")
                            sheet_has_formula = True
                        output_file.write(f"Cell {cell.address} mengandung link atau formula: {cell.formula}\n")
    
    # Tutup workbook setelah selesai
    wb.close()

# Contoh penggunaan
file_path = "Automation 1.2/Liquidity Gap/Result_BSM Stress Test 30082024_Final Template/BSM Stress Test 30082024_Final Template.xlsx"
output_txt_path = "data/output_links.txt"
save_excel_links_to_txt(file_path, output_txt_path)

print(f"Output link dan formula sudah disimpan ke {output_txt_path}")
