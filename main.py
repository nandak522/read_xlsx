import pandas
import os
import sys

def main(xlsx_absolute_path, sheet_name):
    excel_data_df = pandas.read_excel(
        xlsx_absolute_path,
        sheet_name=sheet_name,
        usecols=[
            'Category ',
            'S.No ',
            'Parameters ',
            'Category ',
            'Filter ',
            'Guidelines ',
            'List of audit evidence/Remarks ',
            'Observation on deviation ',
            'Recommendation ']
    )

    # To get the column header names
    # print(excel_data_df.columns.ravel())

    print(excel_data_df[['List of audit evidence/Remarks\xa0', 'Observation on deviation\xa0', 'Recommendation\xa0']])


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python3 main.py <xlsx_absolute_path> <sheet_name>")
        exit(1)
    main(sys.argv[1], argv[2])
