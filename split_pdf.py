import os
from os import path
from PyPDF2 import PdfFileReader, PdfFileWriter


USER   = os.path.expanduser('~')

def pdf_splitter(start,end,filepath,selected_supplier):
    global folder
    folder = USER + "\\Desktop\\" + selected_supplier + "\\SPLIT\\"
    startpage = 0
    endpage   = 0
    startpage = int(start) - 1
    endpage   = int(end)
    pageRange = '[' + str(start) + ',' + str(end) + ']'
    if not path.isdir(folder):
        os.makedirs(folder)

    fname = path.splitext(path.basename(filepath))[0]
    pdf = PdfFileReader(filepath)
    for page in range(pdf.getNumPages()):
        pdf_writer = PdfFileWriter()
        for i in range(startpage,endpage):
            pdf_writer.addPage(pdf.getPage(i))
            output_filename = folder +  '{}_page{}.pdf'.format(
                fname, pageRange)
    with open(output_filename, 'wb') as out:
        pdf_writer.write(out)

    if path.isfile(output_filename):
        return 1
    else:
        return 0


# filepath = "E:\\Php\\cwo\\JS UNITRADE\\PARALLEL DATA\\SI273060-61.pdf"

# pdf_splitter(4,7,filepath)

