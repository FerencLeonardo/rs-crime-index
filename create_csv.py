import re
import os
import pandas as pd

from openpyxl import load_workbook

# Main function
def main():
    directory_path = "/home/leoli/projects/rs-crime-index/data/"
    files_list = get_all_files_list(directory_path)
    create_csv(directory_path, files_list)
    
    return 0

# <----- Single-use functions ----->
def _rename_all_files(directory_path, files_list):
    for file in files_list:
        filename = file[len(directory_path):]
        match = re.search(r"\d{4}", filename)
        if match:
            year = match.group()
            os.rename(file, os.path.join("data/", year + ".xlsx"))
        else:
            raise ValueError(f"No 4-digit year found in filename: {file}")

# This function is very slow!
def _remove_sheet(files_list, sheet):
    for file in files_list:
        wb = load_workbook(file, read_only = True)
        wb.close()
        wb = load_workbook(file)
        if sheet in wb.sheetnames:
            del wb[sheet]
            wb.save(file)
        wb.close()

# <----- Helper functions ----->
def _get_year(file):
    match = re.search(r"\d{4}", file)
    if match:
        year = match.group()
        return year
    else:
        raise ValueError(f"No 4-digit year found in filename: {file}")

def _get_file(files_list, year):
    # Loop over files_list and find the first file that contains the year.
    for file in files_list:
        if _get_year(file) == year:
            return file
    raise ValueError(f"{year} is not in the files_list")

def _get_row_of_word(file, sheet, word):
    row = None
    df = pd.read_excel(
        file,
        sheet_name = sheet,
        header = None
    )
    old_first_name = df.columns[0]
    df = df.rename(columns={old_first_name: "values"})
    for i, val in enumerate(df["values"]):
        if val == word:
            row = i
            break
    return row

# <----- Primary functions ----->
def get_all_files_list(directory_path):
    files_list = []
    for _, _, filenames in os.walk(directory_path):
        for filename in filenames:
            files_list.append(os.path.join(directory_path, filename))
    files_list.sort()
    return files_list

def get_sheet(file, sheet):
    # Open file with pandas to read data
    excel_file = pd.ExcelFile(file)
    df = pd.DataFrame()
    if sheet in excel_file.sheet_names:
        head_row = _get_row_of_word(file, sheet, "Municípios")
        tail_row = _get_row_of_word(file, sheet, "Total RS")

        df = pd.read_excel(
            file,
            sheet_name = sheet,
            usecols="A:O",
            header = head_row,
            nrows = tail_row - head_row,
        )
    return df

def create_csv(directory_path, files_list):
    relevant_sheets = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

    csv_path = os.path.join(directory_path, "csv")
    os.makedirs(csv_path, exist_ok = True)
    for file in files_list:
        current_year = _get_year(file)
        current_path = os.path.join(csv_path, current_year)
        os.makedirs(current_path, exist_ok = True)

        df = get_sheet(file, current_year)
        df.to_csv(os.path.join(current_path, current_year + ".csv"), index = False, encoding = "utf-8-sig")
        for sheet in relevant_sheets:
            print(f"printing sheet {sheet}. of file {file}")
            df = get_sheet(file, sheet)
            if not df.empty:
                df.to_csv(os.path.join(current_path, sheet + ".csv"), index = False, encoding = "utf-8-sig")

if __name__ == "__main__":
    main()


