import pandas as pd
import os

from openpyxl.utils import range_to_tuple


def excel():

    csv_name = input("Enter file name: ")
    if not os.path.exists(csv_name):
        print("File doesn't exist.")
        return

    file_name = os.path.splitext(csv_name)[0]
    excel_file = f"{file_name}.xlsx"

    try:
        print("---- Csv to Excel ----")

        df = pd.read_csv(csv_name)
        df.to_excel(excel_file, index=False)
        file_location = os.path.abspath(excel_file)

        print(f"File location: {file_location}")
    except Exception as e:
        print(e)

if __name__ == "__main__":
    excel()