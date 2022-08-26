import pdfplumber as plum


def read_pdf(path):
    
    with plum.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() 

            print(text)














path = r"E:\Php\cwo\NUTRITIVE\SALES PROFORMA ALTURAS ML 73491-92 APRIL 29.22.pdf"

read_pdf(path)