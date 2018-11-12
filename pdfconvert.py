# from tabula import convert_into
# import sys
#
# # if len(sys.argv)!=3:
# #     print("Input pdf file name, csv file name!!")
# # else:
# #     convert_into("input.pdf", "out.csv", output_format="csv")
#
# convert_into("input01.pdf", "out.csv", output_format="csv", pages='8-9')



import os
import sys
import pdftables_api
from PyPDF2 import PdfFileWriter, PdfFileReader

if len(sys.argv) < 3:
    command = os.path.basename(__file__)
    sys.exit('Usage: {} pdf-file page-number, ...'.format(command))

pdf_input_file = sys.argv[1];
pages_args = ",".join(sys.argv[2:]).replace(" ","")
pages_required = [int(p) for p in filter(None, pages_args.split(","))]

print("Converting pages: {}".format(str(pages_required)[1:-1]))

excel_output_file = pdf_input_file + '.xlsx'

pages_out_of_range = []
pdf_file_reader = PdfFileReader(open(pdf_input_file, 'rb'))
pdf_file_pages = pdf_file_reader.getNumPages()

for page_number in pages_required:
    if page_number < 1 or page_number > pdf_file_pages:
        pages_out_of_range.append(page_number)

if len(pages_out_of_range) > 0:
    pages_str = str(pages_out_of_range)[1:-1]
    sys.exit('Error: page numbers out of range: {}'.format(pages_str))

pdf_writer_selected_pages = PdfFileWriter()

for page_number in pages_required:
    page = pdf_file_reader.getPage(page_number-1)
    pdf_writer_selected_pages.addPage(page)

pdf_file_selected_pages = pdf_input_file + '.tmp'

with open(pdf_file_selected_pages, 'wb') as f:
   pdf_writer_selected_pages.write(f)

c = pdftables_api.Client("rylxfsmhrpxj")
c.xlsx(pdf_file_selected_pages, excel_output_file) #use c.xlsx_single here to output all pages to a single Excel sheet
print("Complete")
os.remove(pdf_file_selected_pages)