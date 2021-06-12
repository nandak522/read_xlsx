import pandas
import os
import sys

def main(xlsx_absolute_path, sheet_name):
    excel_data_df = pandas.read_excel(
        xlsx_absolute_path,
        sheet_name=sheet_name,
        usecols=[
            'Category ',
            'Sr. No. ',
            'Parameters ',
            'Category ',
            'Filter ',
            'Guidelines ',
            'List of audit evidence/Remarks ',
            'Observation on deviation ',
            'Recommendation ']
    )
    print(type(excel_data_df))
    # for item in excel_data_df[['List of audit evidence/Remarks ']]:
    #     print(dir(item))

    # To get the column header names
    print("These are the column header names:%s" % excel_data_df.columns.ravel())
    for index, row in excel_data_df.iterrows():
        print(row['List of audit evidence/Remarks '], row['Observation on deviation '], row['Recommendation '])

    # print(excel_data_df[['List of audit evidence/Remarks\xa0', 'Observation on deviation\xa0', 'Recommendation\xa0']])


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print("Usage: python3 main.py <xlsx_absolute_path> <sheet_name>")
        exit(1)
    main(sys.argv[1], sys.argv[2])
