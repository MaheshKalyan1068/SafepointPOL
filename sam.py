# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import PyPDF2
import io
import pytesseract
import PIL.Image
import tkinter as tk
from tkinter import *
import os
import sys
from pdf2jpg import pdf2jpg
import shutil
import re
import pandas
import pyodbc

folderpath = r"C:\Users\Manomay\Desktop\Policy_Extraction\testing_folder\testing_fol"
#txt = folderpath+"\PolicyExtraction.txt"
file1 = open(r"C:\Users\Manomay\Desktop\Policy_Extraction\POC\PolicyExtraction.txt","a+")
#cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
#sql = "INSERT INTO SP_PolicyExtraction(Formstack_Link, Reported_By,PolicyNumber, Bot_Execution_Status, Bot_Start_Time, Bot_End_Time,ClaimNumber, Error_Details, Actions) VALUES (?,?,?,?,?,?,?,?,?)"
#val = (p,claimname,policy_number,"InProgress",dateint,"None"," "," "," ")
#cursor = cnxn.cursor()
#cursor.execute(sql, val)
#cnxn.commit()
import glob
filelist = []
os.chdir(folderpath)
lsttitle=[]
for file in glob.glob("*.pdf"):
    filelist.append(file)
for i in filelist:
    filepath = folderpath+"\\"+i
    print(filepath)
    dirname = os.path.dirname(filepath)      #filename = os.path.basename(self.PC_paths)

        
    outputpath = os.path.dirname(filepath)+"\Output1"
    pdffilename = os.path.basename(filepath)
#   result = pdf2jpg.convert_pdf2jpg(inputpath, outputpath, dpi=400,pages="2")
    pdf2jpg.convert_pdf2jpg(filepath, outputpath, dpi=350,pages = "ALL")
   # pytesseract.pytesseract.tesseract_cmd = r'C:\Users\mk51620\AppData\Local\Tesseract-OCR\\tesseract'
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\manomay\AppData\Local\Tesseract-OCR\tesseract'

    #extract text from jpg images
    path = os.path.dirname(filepath) +"\\Output\\"

    path = path+ "\\"+pdffilename+"_dir"
    Dpath = path

    files = []
    pdf_pages=[]
    for r, d, f in os.walk(path):
        for file in f:
            if '.jpg' in file:
                files.append(os.path.join(r, file))


    t=""
    openfilename = filepath.replace(".pdf",".txt")
    newfile = open(openfilename,"w+")
    for f in files:
                                             					
        value=PIL.Image.open(f)
        #pdf = pytesseract.image_to_osd(value)
       # if (pdf.find('270') != -1): 
        #    value = value.rotate(270, PIL.Image.NEAREST, expand = 1)
         #   print("PC_image")
       # PC_image = image
        #elif (pdf.find('180') != -1): 
         #   value = value.rotate(180, PIL.Image.NEAREST, expand = 1)
          #  print("PC_image")
        #else: 
         #   print ("Doesn't contains given substring")
        pdf = pytesseract.image_to_pdf_or_hocr(value, extension='pdf')
        pdf_pages.append(pdf)

        pdf_writer = PyPDF2.PdfFileWriter()
        for page in pdf_pages:
            pdf = PyPDF2.PdfFileReader(io.BytesIO(page))
            pdf_writer.addPage(pdf.getPage(0)) 
        path = os.path.dirname(filepath) +"\\Output\\"+pdffilename
        with open(path, 'wb') as out:
            pdf_writer.write(out)                         
        text = pytesseract.image_to_string(value, config='',lang='eng')
        t=t+text
        print(text)
        s=str(t)
       # newfile.write(t)
        print(t)
    match = re.search(r'\D\D\D\D\d\d\d\d\d\d\d.\d\d'  +"|"+ r"\D\D\D\D\d\d\d\d\d\d\d\d\d", t)
    pol = match.group()
    print("Extracted Policy number For PDF"+ i+" is : " + match.group() )
    m = re.search('Name of Applicant:(.+?)\n'+"|"+"Insured:(.+?)\n"+"|"+"Named Insured:(.+?)\n"+"|"+"Named Insured and Mailing Address:(.+?)\n", t)
    if m:
        found = m.group()
    filename12 = pol+"_"+found
    
    lst = ["Renewal Supplemental Application","Citizens Assumption Policies.","COMMON POLICY CHANGE ENDROSEMENT","ACKNOWLEDGEMENT OF CONSENT TO RATE","INSURANCE COVERAGE NOTIFICATION(S) ","COMMERCIAL PROPERTY POLICY DECLARATIONS"]
    lst1 = ["Name of Applicant:","Insured:","Named Insured:"]
    for i in lst:
        if i in s:
            print("Yokoob")
            if i=="Citizens Assumption Policies.":
                i= "Renewal Supplemental Application" + i
            filename13 = filename12+"_"+i
    print(filename13)
    for i in lst1:
        if i in filename13:
            resfile = filename13.replace(i,"")
    print(resfile)
    resfile1 = resfile.replace("\n","")
    lsttitle.append(resfile1)
    #print(outputpath+"\\"+resfile1+".pdf")
    #shutil.copy(filepath,outputpath+"\\"+resfile1+".pdf")
  #  shutil.rmtree(Dpath)
    
    #tk.messagebox.showinfo("Policy Extraction", "Extracted policy Number is "+match.group()+"for "+i)

