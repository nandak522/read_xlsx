import pandas
import os

def main():
    excel_data_df = pandas.read_excel(
        os.path.join(os.getcwd(), 'src/Test.xlsx'),
        sheet_name='OpsHi5!&DEI',
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
    main()
