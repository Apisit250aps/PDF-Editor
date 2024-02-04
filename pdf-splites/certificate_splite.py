from PyPDF2 import PdfWriter, PdfReader
import pandas as pd

file = pd.ExcelFile('list_name.xlsx')
data = pd.read_excel(file, sheet_name='Sheet2')

pdf = PdfReader('certificate.pdf')
pages_number = len(pdf.pages)

for index in range(pages_number):
    writer = PdfWriter()
    cer = pdf.pages[index]
    
    name = data.values[index]
    writer.add_page(cer)
    
    file_output = open(file=f"{index+1} {name[1]} {name[2]}.pdf", mode='wb')
    writer.write(file_output)