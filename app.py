# -*- coding: utf-8 -*-
"""
Created on Tue Aug  4 20:29:14 2020

@author: Manomay
"""
import os
import html5lib
import win32com.client
import pythoncom
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from time import sleep
from selenium.common.exceptions import TimeoutException
import dateutil
import datetime
import numpy as np
import logging
import pyodbc
import dateparser
from flask import Flask, render_template
import win32com.client
import datetime
from flask_material import Material
from flask import Flask, render_template, url_for, request, redirect, Response, session, flash, send_from_directory, \
    send_file, abort, jsonify, logging, make_response
app = Flask(__name__)
Material(app)




@app.route('/')
def index():
    return render_template("index.html")

@app.route("/dashboard", methods=['POST',"GET"])
def dashpage():
    cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.0.13;PORT=1433;database=Saxon_Demo')
    cursor = cnxn.cursor()
    query = """select *  from SP_PolicyExtraction"""
    #sql_query = pd.read_sql_query(query,cnxn)
    table = cursor.execute(query)
    for i in table:
        print(i[0])
    
    ID = []
    Date =[]
    Origin=[]
    Doc_Name =[]
    Meta_Data_Extracted=[]
    PolicyNumber=[]
    Name_Insured=[]
    TypeOf_Doc=[]
    Success_Rate=[]
    Status=[]
    Date_Completed = []
    Action = []
    file_path=[]
    file_rename=[]
    query1 = "select count(*) from SP_PolicyExtraction"
    table = cursor.execute(query)
    for i in table:
        ID.append(i[0])
        Date.append(i[1])
        #Name.append(i[4])
        #Date_Time_Submitted.append(i[2])
        Origin.append(i[2])
        Doc_Name.append(i[3])
        Meta_Data_Extracted.append(i[4])
        PolicyNumber.append(i[5])
        Name_Insured.append(i[6])
        TypeOf_Doc.append(i[7])
        Success_Rate.append(i[8])
        Status.append(i[9])
        Date_Completed.append(i[10])
        Action.append(i[11])
        file_path.append(i[12])
        file_rename.append(i[13])
    print(file_rename,"#############################################################")
    count = cursor.execute(query1)
    count = [int(i[0]) for i in cursor.fetchall()]
    cp = count[0]
    return render_template("dashboard.html",ID=ID,Date=Date,Origin=Origin,file_path=file_path,file_rename=file_rename,Doc_Name=Doc_Name,Meta_Data_Extracted=Meta_Data_Extracted,PolicyNumber=PolicyNumber,Name_Insured=Name_Insured,TypeOf_Doc=TypeOf_Doc,Success_Rate=Success_Rate,Status=Status,Action=Action,Date_Completed=Date_Completed,cp=cp)


@app.route("/images", methods=['POST','GET','OPTIONS'])
def images():
    print("its image path")
    filepath= request.form.get("filepath")
    print(filepath)
    idfile= request.form.get("idfile")
    print(idfile)
    rate= request.form.get("rate")
    print(idfile)
    polnum= request.form.get("polnum")
    print(filepath)
    nain= request.form.get("nain")
    print(filepath)
    tydo= request.form.get("tydo")
    print(filepath)
    filerename = request.form.get("filerename")
    print(filerename)
    t = os.path.split(filepath)
    print(t)
    filepath1 = filepath.replace("/","\\")
    filelist = os.listdir("static/")
    print(filelist)
    imageList = os.listdir('static/'+t[-1])
    imagelist = ['../static/'+t[-1]+"/"+image for image in imageList]
    print(imageList)
    #return render_template("home.html")
    return render_template("home.html", imagelist=imagelist,filepath=filepath,rate=rate,filerename=filerename,polnum=polnum,nain=nain,tydo=tydo,idfile=idfile)
@app.route("/docrename", methods=['POST','GET','OPTIONS'])
def docrename():
    print("Its doc rename dun")
    polnum= request.form.get("polnum")
    nain= request.form.get("nain")
    tydo= request.form.get("tydo")
    rate= request.form.get("rate")
    filerename = request.form.get("filerename")
    print(filerename,"+++++++++++++++++++++++++++")
    if tydo=="Renewal Supplemental Application " +"Citizens Assumption Policies." or tydo=="Renewal Supplemental Application " +"Citizens Assumption Policies":
        sf ="RSACAP"
    elif tydo== "Renewal Supplemental Application":
        sf = "RSA"
    elif tydo=="COMMON POLICY CHANGE ENDROSEMENT":
        sf="CPCE"
    elif tydo=="ACKNOWLEDGEMENT OF CONSENT TO RATE":
        sf="ACR"
    elif tydo=="INSURANCE COVERAGE NOTIFICATION(S) ":
        sf="ICN"
    elif tydo=="COMMERCIAL PROPERTY POLICY DECLARATIONS":
        sf="CPPD"
    else:
        sf="--"
    met = polnum+"_"+nain+"_"+sf
    
    idfile= request.form.get("idfile")
    filepath= request.form.get("filepath")
    t = os.path.split(filerename)
    print(t)
    filerenamed = t[0]+"\\"+met+".pdf"
    os.rename(filerename,filerenamed)
    if met.count("--")==1:
            rate = "70%"
    elif met.count("--")==2:
        rate = "50%"
    elif met.count("--")==3:
        rate = "0%"
    else:
        rate ="100%"
    print(filepath,"its idfile================================================================================================================")
    cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.0.13;PORT=1433;database=Saxon_Demo')
    cursor = cnxn.cursor()
    sql = "UPDATE SP_PolicyExtraction SET PolicyNumber = ?,Name_Insured=?,TypeOf_Doc=?,Success_Rate=?,Meta_Data_Extracted=?,file_rename=? WHERE ID = ?"
    val = (polnum,nain,tydo,rate,met,filerenamed,idfile)
    cursor.execute(sql, val)
    cursor.commit()
    return redirect(url_for('dashpage'))
@app.route("/home", methods=['POST','GET','OPTIONS'])
def home():
    return render_template("home.html")
@app.route("/dochistory", methods=['POST','GET','OPTIONS'])
def dochistory():
    return render_template("dochistory.html")
@app.route("/login", methods=['POST','GET','OPTIONS'])
def login():
    return render_template("index.html")
r"""
@app.route("/link_email", methods=['POST','GET','OPTIONS'])
def link_email():
    msdtxt = ""
    formid= request.form.get("formid")
    #Date_Time_Submitted= request.form.get("Date_Time_Submitted")
    nid= request.form.get("id")
    print(formid,nid)
    print("app.py")
    pythoncom.CoInitialize()
    ol = win32com.client.Dispatch( "Outlook.Application")
    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    for mail in inbox.Items:
        receviedtime = mail.ReceivedTime
        print(mail.SenderEmailAddress)
        sentmail = mail.SenderEmailAddress
        print(receviedtime)
        temp = str(receviedtime).rstrip("+00:00").strip()
        print(temp,"temp")
        #y= datetime.datetime.strptime(temp, "%Y-%m-%d %H:%M:%S")
        #print(y)
        if temp == formid and sentmail == "mahesh.kalyankar@manomay.biz":
            msg = mail.Body
            msghtml = mail.HTMLBody
            print(msghtml)
            dfs = pd.read_html(msghtml,flavor='html5lib')
            df1 = dfs[2]
            df1.columns = ["claim_question","claim_answer"]
            df2 = df1.replace(np.nan, '', regex=True)
            dictcon = pd.Series(df2.claim_answer.values,index=df2.claim_question).to_dict()
           # print(dfs[])
            #data_dict = dfs.to_dict() 
            #print(type(msg))
            #msdtxt = str(msg)
    return render_template('home.html',dictcon=dictcon)
@app.route("/rerun_email", methods=['POST','GET','OPTIONS'])
def rerun_email():
    formid= request.form.get("formid")
    Date_Time_Submitted= request.form.get("Date_Time_Submitted")
    nid= request.form.get("id")
    print(formid,nid,Date_Time_Submitted)
    print("app.py")
    pythoncom.CoInitialize()
    ol = win32com.client.Dispatch( "Outlook.Application")
    inbox = ol.GetNamespace("MAPI").GetDefaultFolder(6)
    for mail in inbox.Items:
        receviedtime = mail.ReceivedTime
        print(receviedtime)
        temp = str(receviedtime).rstrip("+00:00").strip()
        print(temp,"temp")
        #y= datetime.datetime.strptime(temp, "%Y-%m-%d %H:%M:%S")
        #print(y)
        if temp == formid:
            msg = mail.Body
            print(msg)
            try:        
                msg = mail.Body
                receviedtime = mail.ReceivedTime
                print(receviedtime)
                temp = str(receviedtime).rstrip("+00:00").strip()
                print(temp)
                y= datetime.datetime.strptime(temp, "%Y-%m-%d %H:%M:%S")
                print(y)
                p =y.strftime('%Y-%m-%d %H:%M:%S')
                print(y)
                print(p)
                msghtml = mail.HTMLBody
                dfs = pd.read_html(msghtml,flavor='html5lib')
                df1=dfs[2]
                df = df1.replace(np.nan, '', regex=True)
                print(df)
                df.to_csv("com.csv")
                date_time = df.iloc[17,1]
                #datest = date_time.split()
                #datef = dateutil.parser.parse(date_time)
                t = dateparser.parse(date_time)
                #oldformat = datef
                #datetimeobject = datetime.datetime.strptime(str(oldformat),'%Y-%m-%d  %H:%M:%S')
                newformat = t.strftime('%d/%m/%Y')
                print (newformat)
                newtime = t.strftime('%I:%M %p')
                print(newtime)
                policy_number = df.iloc[3,1]
                claimname = df.iloc[4,1]
                ad =df.iloc[18,1]
                t = ad.split(r"\D")
                a = re.split('(\d+)',t[0])
                street = a[0] +" "+a[1]
                zipco=a[-3]+a[-2]
                inti = zipco.split(",")[0]
                vechmaker = df.iloc[13,1]
                vechmod = df.iloc[14,1]
                reg = df.iloc[12,1]
                print(newformat,"/n",newtime,"/n",ad,"/n",street,"/n",zipco,"/n",reg,"/n",vechmaker)
                sleep(1.5)
                dateint1 = datetime.datetime.now()
                dateint = dateint1.strftime("%d/%m/%Y %H:%M:%S")
                chrome = webdriver.Chrome(executable_path='C:\Program Files\chromedriver.exe')
                #chrome = webdriver.Chrome()
                
                chrome.get("https://qasims.saxon.ky/Sims/(S(0vciitns3ckae0h4egjnq2ik))/Login.aspx")
                #chrome.maximize_window()
                
                # wait for element to appear, then hover it
                #wait = WebDriverWait(chrome, 10)
                uname = WebDriverWait(chrome, 10).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#Login1_UserName"))))
                uname.send_keys("Priya")
                pwd = chrome.find_element_by_css_selector("#Login1_Password")
                pwd.send_keys("Saxon2018!")
                chrome.find_element_by_css_selector("#Login1_Button1").click()
                
                New_claim = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl00_btnNewClaim"))))
                New_claim.click()
                
                sleep(1)
                chrome.switch_to.frame("GridEditor")
                
                sleep(3)
                
                chrome.find_element_by_id("ctl03_ctl00_wizNewClaim_chkCoverageVerification").click()
                #check1 = WebDriverWait(chrome, ).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_chkCoverageVerification"))))
                #check1.click()
                
                loss_date=chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_rdpIncidentDate_dateInput")
                loss_date.send_keys(newformat)
                
                Policy_Picker = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_sppPolicy_lnkPolicyPicker").click()
                chrome.switch_to.frame("rwPolicyPicker")
                sleep(2)
                
                inputElement = chrome.find_element_by_name("tbPolicyNumber")
                inputElement.click()
                inputElement.send_keys(policy_number)
                #Polinpt = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.NAME,("#tbPolicyNumber"))))
                #Polinpt.click()
                #Polinpt.send_keys("P05062018-KY-051268")
                
                chrome.find_element_by_css_selector("#btnOK").click()
                sleep(1)
                
                btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#rgPolicy_ctl00_ctl04_btnSelect"))))
                btnselect.click()
                #chrome.find_element_by_css_selector("#rgPolicy_ctl00_ctl04_btnSelect").click()
                chrome.switch_to.default_content()
                #chrome.switch_to.frame("contents")
                chrome.switch_to.frame("GridEditor")
                sleep(2.5)
                #btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_chkUseInsured"))))
                
                chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_chkUseInsured").click()
                #check2 = WebDriverWait(chrome, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_chkUseInsured"))))##ctl03_ctl00_wizNewClaim_chkUseInsured
                #check2.click()
                
                step2nxt = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_StartNavigationTemplateContainerID_btnNext"))))
                step2nxt.click()
                #chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StartNavigationTemplateContainerID_btnNext").click()
                sleep(2)
                step2button = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1"))))
                step2button.click()
                #chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1").click()
                sleep(4)
                okbutton = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CLASS_NAME,("rwOkBtn"))))
                okbutton.click()
                #chrome.find_element_by_class_name("rwOkBtn").click()
                
                time_picker = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInjuryTime_dateInput")
                time_picker.send_keys(newtime)
                secondary_loss = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_rdpIncidentDateSecondary_dateInput")
                secondary_loss.send_keys(newformat)
                
                loss_location = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_lnkAddress").click()
                ad1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_txtAddress1")
                ad1.send_keys(street)
                sleep(0.5)
                text = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_Input")
                text.send_keys(inti)
                sleep(5)
                btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_DropDown > div > ul > li.rcbHovered"))))
                
                
                textlist = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_cboZipCitySearch_DropDown > div")
                li = textlist.parent.find_elements_by_tag_name('li')
                text_str = textlist.text
                textsplit = text_str.split("\n")
                clikindex = textsplit.index(zipco)
                li[clikindex].click()
                chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_addrLossLocationAddress_lnkAddress").click()
                
                examin = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboExaminer_Input")
                examin.send_keys("priya")
                claimant = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInjuryDescription")
                claimant.send_keys("Vechile Damaged")
                claimloss = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtClaimLossDesc")
                claimloss.send_keys("Accident")
                step3_next = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_StepNavigationTemplateContainerID_Button1").click()
                sleep(6)
                btnselect = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input"))))
                vno = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input")
                
                op = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Arrow").click()
                sleep(0.5)
                li = vno.parent.find_elements_by_tag_name('li')
                li[1].click()
                sleep(2)
                #v1 = chrome.find_element_by_css_selector("ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input").click()
                #sleep(2)
                chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInsuredVehicleDesc").click()
                sleep(5.5)
                #namdd1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_Arrow")
                namddlst = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_Arrow")))).click()
                sleep(1)
                lst = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboInsuredDriver_DropDown"))))
                ni = lst.parent.find_elements_by_tag_name('li')
                ni[1].click()
                    
                sleep(3)
                typevech = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicleStyle_Input").get_property("value")
                vech = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleType_Input"))))
                vech.send_keys(typevech)
                sleep(0.25)
                plate =WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantVehicleLicensePlate"))))
                plate.send_keys(reg)
                sleep(0.25)
                year = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtClaimantVehicleYear"))))
                year.send_keys("2015")
                sleep(0.25)
                islandtxt = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicleState_Input").get_property("value")
                island = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleState_Input"))))
                island.send_keys(islandtxt)
                sleep(0.25)
                make = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleMake_Input"))))
                make.send_keys(vechmaker)
                sleep(0.25)
                model = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleModel_Input"))))
                model.send_keys(vechmod)
                sleep(0.25)
                colourtxt = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicleColor_Input").get_property("value")
                print(colourtxt)
                colour = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_cboClaimantVehicleColor_Input"))))
                colour.send_keys(colourtxt)
                sleep(0.25)
                licentext = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_txtInsuredDriverLicenseNumber").get_property("value")
                licen = WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CSS_SELECTOR,("#ctl03_ctl00_wizNewClaim_txtDriverLicenseNumber"))))
                licen.send_keys(licentext)
                sleep(1)
                coverage = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_Input").click()
                sleep(1.5)
                lbox = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_DropDown > div.rcbScroll.rcbWidth")
                li = lbox.parent.find_elements_by_tag_name('li')
                c1 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i0_CheckBox").click()
                sleep(0.4)
                #c2 = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i4_CheckBox").click()
                #sleep(0.4)
                #c3= chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboCoverage_i1_CheckBox").click()
                
                #sleep(0.4)
                losstype = chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboLossType_Input")
                losstype.send_keys("Auto")
                sleep(3)
                chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_cboInsuredVehicle_Input").click()
                sleep(2)
                chrome.find_element_by_css_selector("#ctl03_ctl00_wizNewClaim_FinishNavigationTemplateContainerID_btnCreate").click()
                sleep(8)
                WebDriverWait(chrome, 20).until(ec.visibility_of_element_located((By.CLASS_NAME,("rwDialogMessage"))))
                t=chrome.find_element_by_class_name("rwDialogMessage")
                poltext = t.text
                res = ''.join(filter(lambda i: i.isdigit(), poltext))
                claimno = "#"+res
                print(claimno)
                sleep(2)
                #chrome.find_element_by_class_name("rwOkBtn").click()
                chrome.close()
                #sleep(5)
                submittime1 = datetime.datetime.now()
                submittime = submittime1.strftime("%d/%m/%Y %H:%M:%S")
                #chrome.close()
                cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.1.5;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
                cursor = cnxn.cursor()
                sql = "INSERT INTO Saxon_Claims(Formstack_Link, Reported_By,PolicyNumber, Bot_Execution_Status, Bot_Start_Time, Bot_End_Time,ClaimNumber, Error_Details, Actions) VALUES (?,?,?,?,?,?,?,?,?)"
                val = (p,claimname,policy_number,"Success",dateint,submittime,claimno," ","Rerun")
                print(val)
                cursor.execute(sql, val)
                cnxn.commit()
                cursor.close()
                cnxn.close()
            except Exception as e:
                print(e)
                t=chrome.get_log('browser')
                #print(t)
                #print(e)
                chrome.save_screenshot(r"C:\Users\Manomay\\"+policy_number+'.png')
                chrome.close()
                inbox = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
                #print(dir(inbox))
                inbox = win32com.client.Dispatch("Outlook.Application")
                #print(dir(inbox))
                mail = inbox.CreateItem(0x0)
                mail.To = "mahesh.kalyankar@manomay.biz"
                m = mail.To
                mail.Attachments.Add(r"C:\Users\Manomay\\"+policy_number+'.png')
                #mail.CC = "testcc@test.com"
                mail.Subject = "Saxon Bot Failure Report"
                mail.Body = "Hi " + m.split(".")[0].capitalize() +"\n" +"Please find error report below"+"\n" +str(t)+"\n"+str(e)+"\n" +"Thank you."
                                
                mail.Send()
                cnxn = pyodbc.connect('DRIVER={SQL Server};server=192.168.1.5;PORT=1433;database=Saxon_Demo;uid=MANOMAY1;pwd=manomay1')
                cursor = cnxn.cursor()
                submittime1 = datetime.datetime.now()
                submittime = submittime1.strftime("%d/%m/%Y %H:%M:%S")
                sql = "INSERT INTO Saxon_Claims(Formstack_Link, Reported_By,PolicyNumber, Bot_Execution_Status, Bot_Start_Time, Bot_End_Time,ClaimNumber, Error_Details, Actions) VALUES (?,?,?,?,?,?,?,?,?)"
                val = (p,claimname,policy_number,"Failure",dateint,submittime," ","Error Occured","RERUN")
                print(val)
                cursor.execute(sql, val)
                cnxn.commit()
                cursor.close()
                cnxn.close()
    return redirect(url_for('dashpage'))
"""
if __name__ == '__main__':
    app.run(debug=True)
