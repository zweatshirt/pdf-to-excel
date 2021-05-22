# pdf-to-excel
Program I'm making to automate a portion of my mom's job. 
Still a WIP

Flow:
- uses tabula-py to extract data from a PDF
- data gets converted a pandas DataFrame, DataFrame -> Dict
- data stored in dict gets transferred over to .xlsx file copied from template
using openpyxl

