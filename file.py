import os
import re
import xlwings as xw
import shutil

def replace_date_in_filename(old_filename, new_date_str):
    date_pattern = re.compile(r'\d{2}\s?[A-Za-z]+\s?\d{4}')
    new_filename = date_pattern.sub(new_date_str, old_filename)
    print(f"Updating filename from {old_filename} to {new_filename}")
    return new_filename

def replace_external_references(wb, file_mapping):
    file_reference_pattern = re.compile(r'(\[.*?\])')
    changes_made = False  # Flag to check if any changes are made
    for sheet in wb.sheets:
        for cell in sheet.used_range:
            if cell.formula:
                original_formula = cell.formula
                matches = file_reference_pattern.findall(cell.formula)
                if matches:
                    new_formula = cell.formula
                    for match in matches:
                        original_file = match[1:-1]  # Remove the [ and ] brackets
                        if original_file in file_mapping:
                            new_file = file_mapping[original_file]
                            new_formula = new_formula.replace(match, f'[{new_file}]')
                            print(f"Replacing {match} with [{new_file}] in formula of cell {cell.address} in sheet {sheet.name}")
                            changes_made = True
                    if new_formula != original_formula:
                        print(f"Updating formula in cell {cell.address} from {original_formula} to {new_formula} in sheet {sheet.name}")
                        cell.formula = new_formula
                else:
                    print(f"No external file references found in cell {cell.address} formula: {cell.formula}")
    if not changes_made:
        print("No changes were made to the workbook.")

def generate_bsm_stress_test_template(source_directory, target_directory, new_date_str, file_mapping):
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)
    for filename in os.listdir(source_directory):
        if filename.endswith(('.xlsx', '.xlsm')):
            source_file_path = os.path.join(source_directory, filename)
            new_filename = replace_date_in_filename(filename, new_date_str)
            new_file_path = os.path.join(target_directory, new_filename)
            shutil.copy(source_file_path, new_file_path)
            wb = xw.Book(new_file_path)
            replace_external_references(wb, file_mapping)
            wb.save()
            wb.close()
            print(f"Processed and saved: {new_file_path}")

def create_file_mapping(source_directory):
    return {
        "BSM Stress Test 29082024_Final Template.xlsx": "BSM Stress Test 30082024_Final Template.xlsx",
        "BSM Stress Test_P+I.xlsx": "BSM Stress Test_P+I.xlsx",
        "BSM Stress Test.xlsx": "BSM Stress Test.xlsx",
        "Daily Report_30082024.xlsx": "Daily Report_30082024.xlsx",
        "PTBN_LNBR.xlsm": "PTBN_LNBR.xlsm",
        "Update Assumption_2019 NEW.xlsx": "Update Assumption_2019 NEW.xlsx"
    }

source_directory = "/Users/owwl/Downloads/Automation/Automation 1.1/Liquidity Gap/Support Files_BSM Stress Test 30082024_Final Template"
target_directory = "/Users/owwl/Downloads/Automation/Automation 1.1/Liquidity Gap/Result_BSM Stress Test Final"
new_date_str = "30 Aug 2024"

file_mapping = create_file_mapping(source_directory)
generate_bsm_stress_test_template(source_directory, target_directory, new_date_str, file_mapping)
