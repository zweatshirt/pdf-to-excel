from tabula import read_pdf as rp
from pandas import DataFrame
# from shutil import copy
from openpyxl import Workbook
from openpyxl import load_workbook


def main():
    df_dict = pdf_to_dict()

    # Parse date from dict:
    date =  list(df_dict[0].values())[2]
    print(date)

    # Random testing
    for i in range(1, len(df_dict)):
        print(df_dict[i]['Employee'])

    # Grab name to save new excel file as:
    xlsx_name = str(input("What would you like to name the excel file?\n"))

    # Create workbook from template:
    wb = load_workbook(filename='template.xlsx')
    ws = wb.active


def pdf_to_dict():
    df_dict = {}
    while True:
        try:
            pdf_name = str(input("Enter full path of the PDF to read: "))
            # Convert pdf info to DataFrame, then to dict
            df = DataFrame(rp(pdf_name, pages='all')[0])
            df_dict = df.to_dict('index')
            break
        except IOError:
            print("Error parsing PDF, please try again.")

    return df_dict




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()