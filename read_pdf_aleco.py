import pdfplumber as plum
import win32com.client
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path
from os import path
from db import Database



def read_pdf(file_path):
    with plum.open(file_path) as pdf:
        print(pdf)
        for page in pdf.pages:
            text = page.extract_text() 

            print(text)