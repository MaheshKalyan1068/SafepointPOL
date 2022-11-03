# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pytesseract
import PIL.Image
import tkinter as tk
from tkinter import *
import os
import sys
from pdf2jpg import pdf2jpg
import shutil
import re
filepath = r"C:\Users\Manomay\Desktop\Policy_Extraction\testing_folder\testing_fol\Leonhard _ Sullivan.pdf"
filepath = r"C:\Users\MANOMAY\Desktop\Policy_Extraction\ht\Howard, renewal.pdf"
dirname = os.path.dirname(filepath)
print(dirname,"dir_Name")       #filename = os.path.basename(self.PC_paths)

file1 = open(r"C:\Users\Manomay\Desktop\Policy_Extraction\PolicyExtraction.txt","a+")
outputpath = os.path.dirname(filepath)+"\Output"
print(outputpath,"Output_Name")
pdffilename = os.path.basename(filepath)
print(pdffilename,"filepath")
#   result = pdf2jpg.convert_pdf2jpg(inputpath, outputpath, dpi=400,pages="2")
pdf2jpg.convert_pdf2jpg(filepath, outputpath, dpi=350,pages = "ALL")
   # pytesseract.pytesseract.tesseract_cmd = r'C:\Users\mk51620\AppData\Local\Tesseract-OCR\\tesseract'
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\manomay\AppData\Local\Tesseract-OCR\tesseract'

#extract text from jpg images
path = os.path.dirname(filepath) +"\\Output\\"
print(path,"Mahesh")

path = path+ pdffilename+"_dir"
print(path,"path12")
Dpath = path

files = []
for r, d, f in os.walk(path):
    for file in f:
        if '.jpg' in file:
            files.append(os.path.join(r, file))


t=""
for f in files:                                     					
    value=PIL.Image.open(f)
    #pdf = pytesseract.image_to_osd(value)
    #if (pdf.find('270') != -1): 
     #   value = value.rotate(270, PIL.Image.NEAREST, expand = 1)
       # print("PC_image")
       # PC_image = image
   # elif (pdf.find('180') != -1): 
    #    value = value.rotate(180, PIL.Image.NEAREST, expand = 1)
     #   print("PC_image")
    #else: 
     #   print ("Doesn't contains given substring")                         
    text = pytesseract.image_to_string(value, config='',lang='eng')
    t=t+text
print(t)    
match = re.search(r'\D\D\D\D\d\d\d\d\d\d\d.\d\d' +"|"+ r"\D\D\D\D\d\d\d\d\d\d\d\d\d", t)
match.group()
shutil.rmtree(Dpath)
file1.write("Extracted Policy number For PDF is" + match.group() )
file1.close()
tk.messagebox.showinfo("Policy Extraction", "Extracted policy Number is "+match.group())

