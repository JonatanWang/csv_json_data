#! python3
# A program that reads in all the Excel files in the current working directory
# and outputs them as CSV files.
import csv, os, xlrd


def excel_to_csv():
    os.chdir('./excelSpreadsheets')
    for filename in os.listdir('.'):
        print(f'{filename}')
        wb = xlrd.open_workbook(filename)
        # Loop through every sheet in the workbook
        for sh in wb.sheets():
            # Create the csv filename from the excel filename and sheet title
            csv_file = open(f'output.csv', 'wb')

            # Create the csv.writer object for the this csv file
            csv_writer = csv.writer(csv_file)

            # Write the rowData list to the csv file
            for row in range(sh.nrows):
                row_data = []
                for col in range(sh.ncols):
                    value = sh.cell(row, col).value
                    row_data.append(value)
                csv_writer.writerow(row_data)
            csv_file.close()


if __name__ == '__main__':
    excel_to_csv()
