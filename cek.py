import xlwings as xw
import re
import os
import shutil
from datetime import datetime

# Fungsi untuk mendapatkan data dari file Excel eksternal
def get_data_from_external_file(file_path, sheet_name, cell_address):
    try:
        print(f"Opening external file: {file_path}")
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]
        
        if sheet is None:
            print(f"Sheet {sheet_name} not found in {file_path}")
            return None
        
        # Ambil nilai dari cell yang dirujuk
        data = sheet.range(cell_address).value
        wb.close()
        return data
    except Exception as e:
        print(f"Error accessing {file_path}: {e}")
        return None

# Fungsi untuk mengupdate formula di workbook dengan path file baru dan mengambil data dari file eksternal
def update_workbook_with_data(source_path, new_date, path_mappings):
    # Membuat salinan dari template lama dengan tanggal baru
    new_file_name = f"BSM Stress Test {new_date.strftime('%d%m%Y')}_Final Template.xlsx"
    new_file_path = os.path.join("test", new_file_name)
    
    print(f"Copying template from {source_path} to {new_file_path}...")
    shutil.copy(source_path, new_file_path)

    print(f"Opening workbook {new_file_path}...")
    wb = xw.Book(new_file_path)

    try:
        # Proses setiap sheet di workbook
        for sheet in wb.sheets:
            print(f"Processing sheet: {sheet.name}")
            for cell in sheet.used_range:
                if cell.formula:
                    original_formula = cell.formula
                    for old_path, new_path in path_mappings.items():
                        if old_path in original_formula:
                            # Ganti path dengan path yang baru
                            new_formula = original_formula.replace(old_path, new_path)
                            
                            # Ekstrak informasi worksheet dan cell dari formula lama
                            match = re.search(r'\[(.*?)\](.*?)!(.*?)', original_formula)
                            if match:
                                external_file = match.group(1)  # Nama file
                                sheet_name = match.group(2)     # Nama sheet
                                cell_address = match.group(3)   # Alamat cell
                                
                                # Ambil data dari file eksternal
                                external_data = get_data_from_external_file(new_path, sheet_name, cell_address)
                                if external_data:
                                    cell.value = external_data  # Masukkan nilai dari file eksternal ke dalam workbook
                                    print(f"Updated cell {cell.address} with data from {external_file} ({sheet_name} {cell_address})")

                            # Update formula
                            cell.formula = new_formula
                            print(f"Old formula: {original_formula}")
                            print(f"New formula: {new_formula}")
                            original_formula = new_formula  # Update formula jika terjadi perubahan

        print(f"Saving workbook to {new_file_path}...")
        wb.save()
    finally:
        # Pastikan workbook ditutup setelah selesai
        print("Closing workbook...")
        wb.close()

# Contoh penggunaan
if __name__ == "__main__":
    # Path template lama
    source_template_path = "Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template/BSM Stress Test 29082024_Final Template.xlsx"
    
    # Tanggal dari Daily Report
    new_report_date = datetime.strptime('30082024', '%d%m%Y')

    # Mapping dari path lama ke path baru
    path_mappings = {
        '/10.27.12.90/REPORT/1-DAILY/2-REPORT/3-MCO DAILY/2024/08. Aug/29 Aug 2024/[BSM Stress Test 29082024_Final Template.xlsx]': 'Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template/BSM Stress Test 29082024_Final Template.xlsx',
        '/10.27.12.90/ALMRiskManagement/REPORT/1-DAILY/2-REPORT/3-MCO DAILY/2024/08. Aug/30 Aug 2024/[BSM Stress Test_P+I.xlsx]': 'Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template/BSM Stress Test_P+I.xlsx',
        '/10.27.12.90/ALMRiskManagement/REPORT/1-DAILY/2-REPORT/3-MCO DAILY/2024/08. Aug/30 Aug 2024/[BSM Stress Test.xlsx]': 'Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template/BSM Stress Test.xlsx',
        '/10.27.12.90/REPORT/1-DAILY/1-WORKSHEET/1-LIKUIDITAS HARIAN/Neraca Harian/2024/08. Aug/[Daily Report_30082024.xlsx]': 'Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template/Daily Report_30082024.xlsx',
        '/10.27.12.90/REPORT/1-DAILY/2-REPORT/8. Interbank Activities/[PTBN_LNBR.xlsm]': 'Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template/PTBN_LNBR.xlsm',
        '/10.27.12.90/REPORT/[Update Assumption_2019 NEW.xlsx]': 'Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template/Update Assumption_2019 NEW.xlsx'
    }

    # Jalankan fungsi untuk update link file
    print("Starting to update workbook links with data from external files...")
    update_workbook_with_data(source_template_path, new_report_date, path_mappings)

    print("Referensi file sudah diperbarui dan disimpan di direktori 'test'.")
