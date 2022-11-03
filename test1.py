import pyodbc
import PyPDF2
import io
import pytesseract
import PIL.Image
import tkinter as tk
from tkinter import *
import glob,os
import sys
from pdf2jpg import pdf2jpg
import shutil
import re
import pandas as pd
import datetime
from encodings import cp1252
from tkinter import filedialog
from tkinter import *
from tkinter import ttk



def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global e2
    filename = filedialog.askdirectory()
    e2.set(filename)
def poilcy_exe():
    totfiles = 0
    original_decode  = cp1252.Codec.decode
    cp1252.Codec.decode =  lambda self, input, errors="replace": original_decode(self, input, errors)
    folderpath = e2.get()
    filelist = []
    os.chdir(folderpath)
    cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
    cursor = cnxn.cursor()
    for file in glob.glob("*.pdf"):    #folderpath = r"C:\Users\MANOMAY\Desktop\Policy_Extraction\testing_folder\testing_fol\testing_fol_main"
        filelist.append(file)
    progressBar.config(maximum = len(filelist))
#    totalcount = len(filelist)
    for i in filelist:
        docname = i
        totfiles = totfiles + 1
        dateint1 = datetime.datetime.now()
        filepath = folderpath+"\\"+i       
        outputpath = os.path.dirname(filepath)+"\Output1"
        pdffilename = os.path.basename(filepath)
        pdf2jpg.convert_pdf2jpg(filepath, outputpath, dpi=350,pages = "ALL")
        pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'    
        path = os.path.dirname(filepath) +"\\Output1\\"   
        path = path+ "\\"+pdffilename+"_dir"
        t= os.path.split(os.path.dirname(filepath))
        pathout1 = t[0]+"\\"+"static\\"+pdffilename+"_dir"
        pathout = path  
        files = []
        #pdf_pages=[]
        for r, d, f in os.walk(path):
            for file in f:
                if '.jpg' in file:
                    files.append(os.path.join(r, file))
    
    
        t=""
        #openfilename = filepath.replace(".pdf",".txt")
        #newfile = open(openfilename,"w+")
        for f in files:
            
            #print(f,"f")                                     					
            value=PIL.Image.open(f)
            text = pytesseract.image_to_string(value, config='',lang='eng')
            t=t+text
        try:
            match = re.search(r'\D\D\D\D\d\d\d\d\d\d\d.\d\d'  +"|"+ r"\D\D\D\D\d\d\d\d\d\d\d\d\d"+"|"+ r"\D\D\D\D\d\d\d\d\d\d\d", t)
            pol = match.group()
            policy = pol
        except:
            policy = "--"
        
        #print("Extracted Policy number For PDF"+ i+" is : " + match.group() )
        try:
            m = re.search('Name of Applicant:(.+?)\n'+"|"+"Insured:(.+?)\n"+"|"+"Named Insured:(.+?)\n"+"|"+"Named Insured and Mailing Address:(.+?)\n", t)
            nameit = m.group()
            lst1=['Name of Applicant:',"Insured:","Named Insured:",'Named Insured and Mailing Address:']
            for j in lst1:
                if j in nameit:
                        resfile = nameit.replace(j,"")
        except:
            resfile = "--"
        lst = ["Renewal Supplemental Application","Citizens Assumption Policies.","Citizens Assumption Policies","COMMON POLICY CHANGE ENDROSEMENT","ACKNOWLEDGEMENT OF CONSENT TO RATE","INSURANCE COVERAGE NOTIFICATION(S) ","COMMERCIAL PROPERTY POLICY DECLARATIONS"]
        try:
            for i in lst:
                if i in t:

                    typedoc = i
                    if i=="Citizens Assumption Policies." or i=="Citizens Assumption Policies":
                        i= "Renewal Supplemental Application " + i
                        print(i)
                        typedoc = i
                    if i=="Renewal Supplemental Application " +"Citizens Assumption Policies." or i=="Renewal Supplemental Application " +"Citizens Assumption Policies":
                        sf ="RSACAP"
                    elif i== "Renewal Supplemental Application":
                        sf = "RSA"
                    elif i=="COMMON POLICY CHANGE ENDROSEMENT":
                        sf="CPCE"
                    elif i=="ACKNOWLEDGEMENT OF CONSENT TO RATE":
                        sf="ACR"
                    elif i=="INSURANCE COVERAGE NOTIFICATION(S) ":
                        sf="ICN"
                    elif i=="COMMERCIAL PROPERTY POLICY DECLARATIONS":
                        sf="CPPD"
                    else:
                        sf="--"
        except:
            typedoc = "--"
        try:
            lst1=['Name of Applicant:',"Insured:","Named Insured:",'Named Insured and Mailing Address:']
            for j in lst1:
                if j in nameit:
                    resfile = nameit.replace(j,"")                    
            resfile1 = resfile.replace("\n","")
            resfile1 = resfile1.replace("/","")
            resfile1 = resfile1.replace("|","")
            resfile1 = resfile1.replace("}","")
            resfile1 = resfile1.replace("[","")
            
            #cnxn = pyodbc.connect('DRIVER={SQL Server};server=183.82.0.186;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
            #sql = "UPDATE SP_PolicyExtraction SET Meta_Data_Extracted = ?,Status = ? WHERE ID = (SELECT max(ID) FROM SP_PolicyExtraction)"
           # metadata = policy+"_"+resfile1+"_"+typedoc
        except:
            resfile1 = "--"

        dateend = datetime.datetime.now()
        renamepath = outputpath+"\\"+policy +"_"+resfile1+"_"+sf+".pdf"
        metdata = policy +"_"+resfile1+"_"+sf+".pdf"
        if metdata.count("--")==1:
            rate = "70%"
        elif metdata.count("--")==2:
            rate = "50%"
        elif metdata.count("--")==3:
            rate = "0%"
        else:
            rate ="100%"
        sql = "INSERT INTO SP_PolicyExtraction(Date,Doc_Name,Origin,File_path,PolicyNumber,Name_Insured,TypeOf_Doc,Meta_Data_Extracted,Date_Completed ,Status,Success_Rate ,file_rename) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"
        val = (dateint1,docname,"Scan",pathout1,policy,resfile1,typedoc,metdata,dateend,"Sucess",rate,renamepath)
        print(renamepath,"oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo")
        cursor.execute(sql, val)
        cursor.commit()
        progressBar.config(value = totfiles)
        progressBar.update()
        shutil.copy(filepath,outputpath+"\\"+policy +"_"+resfile1+"_"+sf+".pdf")
        shutil.move(pathout, pathout1)
    cursor.close()
    cnxn.close()
    root.destroy()

root = Tk()
root.title('Safepoint Policy Extraction')
root.geometry('500x150')
root.configure(bg="#065b83")
folder_path = StringVar()
Label(root,text="Select Folder of Pdf's  ",bg="#065b83").grid(row=1,pady = 15)
e2 = StringVar()
e2_1 = Entry(root,width=45,textvariable=e2)
e2_1.grid(row=1, column=1,pady = 15,padx = 5)
#lbl1 = Label(master=root,textvariable=folder_path)
#lbl1.grid(row=0, column=1)
button2 = Button(text="Browse",padx = 5, command=browse_button)
button2.grid(row=1, column=3)
button1 = Button(text="Submit", command=poilcy_exe)
button1.grid(row=3, column=1,padx=10,pady = 8)
Label(root,text="Execution Status: ",bg="#065b83").grid(row=5,column = 0,pady = 10)
progressBar = ttk.Progressbar(root, orient="horizontal", length=300,mode="determinate")
progressBar.grid(row =5,column = 1)
#Label(root,text="Total No. Of Files : ").grid(row=6,column = 0)
mainloop()
    
