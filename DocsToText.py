import pandas as pd
import pyodbc 
import json
import imaplib
import base64
import os
import email
import datetime
from email.header import decode_header
from email.message import EmailMessage
from win32com import client as wc
import re
import base64
import quopri
from  os.path import splitext
import io
import csv
#import sys
import docx2txt
import codecs
import win32com.client
import docx
import PyPDF2
import pikepdf
import chardet    
import subprocess
import traceback
import zipfile
import EmailTools

temp_path = "e:/test/att/"
temp_path_ = "e:\\test\\att\\"

def convert_xls_to_txt (FileText):
    FileTextString = ''
    Separator = '\t'
    my_excel_file_io = io.BytesIO(FileText)

    try:

        xl= pd.ExcelFile(my_excel_file_io)

        fn = pd.read_excel(xl, sheet_name = xl.sheet_names[0])  
        s_buf = io.StringIO()
        s_buf.seek(0)
        fn.to_csv(s_buf, 
                    index=False, sep=Separator, 
                    float_format='%g',
                    date_format='%d.%m.%Y')
        FileTextString = s_buf.getvalue()
    except Exception as e:

        FileTextString = 'ERROR: ' + str(e) + '\r\n' + traceback.format_exc()
    
    return FileTextString


def convert_docx_to_txt(FileText):
    FileTextString = ''
    s_buf = io.BytesIO(FileText)
    
    try :
        document = docx.Document(s_buf)
        style = document.styles['Normal']
        font = style.font
        font.size = docx.shared.Pt(4)
        document.save(s_buf)
        FileTextString = docx2txt.process(s_buf)
    except: 
        FileTextString = docx2txt.process(s_buf)
    return FileTextString

def convert_doc_to_txt(FileText):
    FileTextString = ''
    TempFile = datetime.datetime.now().strftime("temp-%Y-%m-%d-%H-%M-%S.doc")

    f = open(temp_path + TempFile, 'wb')
    f.write(FileText)
    f.close()

    w = wc.Dispatch('Word.Application')
    try:                        
        doc=w.Documents.Open(temp_path + TempFile)
        doc.SaveAs(temp_path+ TempFile+"x",16)# Must have parameter 16, otherwise an error will occur.
        doc.Close()
        #word.Quit()
    
        document = docx.Document(temp_path+ TempFile+"x")
        style = document.styles['Normal']
        font = style.font
        font.size = docx.shared.Pt(4)
        document.save(temp_path+ TempFile + "x")
        FileTextString = docx2txt.process(temp_path + TempFile + "x")
        os.remove(temp_path+ TempFile)
        os.remove(temp_path+ TempFile + "x")

    except Exception as e:

        FileTextString = 'ERROR: ' + str(e) + '\r\n' + traceback.format_exc()

    return FileTextString                        
                    

def convert_pdf_to_txt(FileText, HavePassword= ''): 
    FileTextString = ''
    pdf1 = io.BytesIO(FileText)
    pdf_no_pass = pikepdf.open(pdf1,password=HavePassword)
    pdf = io.BytesIO()
    pdf_no_pass.save(pdf)

    TempFile = datetime.datetime.now().strftime("temp-%Y-%m-%d-%H-%M-%S.pdf")
    f = open(temp_path + TempFile, 'wb')
    f.write(pdf.getvalue())
    f.close()

    cmd_text = 'E:\\ivanm\\BATH-FILES\\exe\\pdftotext.exe -table ""' + temp_path_ \
    + TempFile +'"" ""' + temp_path_ + TempFile +'.txt""'
                        
    subprocess.run(cmd_text)

    f = open(temp_path + TempFile +'.txt', 'r')
    FileTextString = f.read()
    f.close()


    os.remove(temp_path + TempFile)
    os.remove(temp_path + TempFile + ".txt")

    return FileTextString

