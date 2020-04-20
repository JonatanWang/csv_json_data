#! python3
# A program that reads in all the Excel files in the current working directory
# and outputs them as CSV files.
import csv, os, xlrd


def excel_to_csv():
    os.chdir('./excelSpreadsheets')
    for filename in os.listdir('.'):
        wb = xlrd.open_workbook('spreadsheet-A.xlsx')
        # Loop through every sheet in the workbook
        sheet = wb.sheet_by_name('Sheet')
        csv_file = open(f'{filename}.csv', 'w', encoding='utf8')
        wr = csv.writer(csv_file, quoting=csv.QUOTE_ALL)
        for row_num in range(sheet.nrows):
            wr.writerow(sheet.row_values(row_num))
        csv_file.close()


if __name__ == '__main__':
    excel_to_csv()
