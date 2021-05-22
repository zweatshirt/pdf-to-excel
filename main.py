from tabula import read_pdf as rp
from pandas import DataFrame
# from shutil import copy
from openpyxl import Workbook
from openpyxl import load_workbook
import traceback as tb


def main():
    df_dict = pdf_to_dict()

    # Parse date from dict:
    date =  list(df_dict[0].values())[2]
    print(date)
    print_pretty_dict(df_dict)


    # Grab name to save new excel file as:
    xlsx_name = str(input("What would you like to name the excel file?\n"))

    # Create workbook from template:
    wb = load_workbook(filename='template.xlsx')
    ws = wb.active


def print_pretty_dict(df_dict):
    assert isinstance(df_dict, dict)
    cols = list(x for x in df_dict[0])
    # all rows
    for i in range(0, len(df_dict)):
        if i == 0:
            print(str.join('            ', (i for i in cols)))
        else:
            employee = df_dict[i][cols[0]]
            job = df_dict[i][cols[1]]
            shift = df_dict[i][cols[2]]
            print(str.join('        ', [repr(employee), repr(job), repr(shift)]))

def pdf_to_dict():
    while True:
        try:
            pdf_name = str(input("Enter full path of the PDF to read: "))
            # Convert pdf info to DataFrame, then to dict
            df = DataFrame(rp(pdf_name, pages='all')[0])
            # Essentially just a 0 thru n - 1 indexed dict, where n == num of keys
            # Each key points to a row in the pdf
            df_dict = df.to_dict('index')
            break
        except Exception:
            tb.print_exc()

            print("Error parsing PDF, please try again.")

    return df_dict


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()