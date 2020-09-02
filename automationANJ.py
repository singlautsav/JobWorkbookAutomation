# from xlrd import *
from datetime import date
import sqlite3
import subprocess
from openpyxl import load_workbook
today = date.today()
import schedule
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import time
import datetime
import os
import argparse



# def queryOldTable():

# conn = sqlite3.connect('cciAutomation.db')
# c = conn.cursor()

# def updateOldTable():

def editRecord(val):
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    listX = []
    listX.append(input("*Client Name: "))
    listX.append(input("Client Phone Number: ")) 
    listX.append(input("Client Email: "))
    listX.append(input("*Item Recieved: "))
    listX.append(input("*Model Number: "))
    listX.append(input("*Problem Reported: "))
    listX.append(input("Problem Diagnosed: "))
    listX.append(input("Additional Comments: "))
    listX.append(input("Minimum Charges: "))
    if listX[-1]!='':
        listX[-1]=int(listX[-1])
    x = (input("Job Detail, Pending or Completed, (P/C): "))
    if x.startswith('P') or x.startswith('p'):
        listX.append("Pending")
    elif x.startswith("C") or x.startswith("c"):
        listX.append("Completed")
    else:
        listX.append('')
    listX.append(input("Advance Recieved: "))
    if listX[-1]!='':
        listX[-1] = int(listX[-1])
    q = f'''SELECT * FROM jobTask j where j.jobNumber={val}'''
    x = c.execute(q)
    valX = x.fetchall()[0]
    valX = list(valX)
    sheet['E8'] = valX[1]
    '''ChangeLine'''
    sheet['F8'] = "ANJ-"+str(valX[0])
    valX = valX[2:]
    # valX = ['' if v is None else v for v in valX]
    listX = [None if v is '' else v for v in listX]
    print(valX)
    # print(listX)
    finalRecord = []
    for i in range(len(listX)):
        if listX[i]==valX[i]:
            finalRecord.append(listX[i])
        else:
            if listX[i]!=None:
                finalRecord.append(listX[i])
            else:
                finalRecord.append(valX[i])
    sheet['C7'] = finalRecord[0]
    sheet['C8'] = finalRecord[1]
    sheet['C9'] = finalRecord[2]
    sheet['C10'] = finalRecord[3]
    sheet['C11'] = finalRecord[4]
    sheet['C13'] = finalRecord[5]
    sheet['C14'] = finalRecord[6]
    sheet['C15'] = finalRecord[7]
    sheet['B16'] = f"We are not Responsible for any data losses. Minimum diagnostic charges are: {finalRecord[8]}"
    sheet['E10'] = finalRecord[8]
    sheet['F11'] = finalRecord[9]
    sheet['F10'] = finalRecord[10]
    q = """UPDATE jobTask SET clientName = ?, clientNumber = ?, clientEmail = ?, itemRecieved = ?, modelNumber = ?, ProblemReported = ?, ProblemDiagnosed = ?, AdditionalComments = ?, minimumCharges = ?, workDetail = ?, advanceRecieved = ? WHERE jobNumber = ?"""
    finalRecord.append(int(val))
    # finalRecord = set(finalRecord)
    print(finalRecord)
    x = c.execute(q,finalRecord)
    conn.commit()
    c.close()
    conn.close()


def updateRecord(val):
    stringX = "Completed"
    q = f'''UPDATE jobTask SET workDetail = "Completed" WHERE jobNumber = {val}'''
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    x = c.execute(q)
    print(f"Updated Record number {val} to completed")
    conn.commit()
    c.close()
    conn.close()


def deleteARecord(val):
    q= f''' DELETE FROM jobTask where jobNumber ={val}'''
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    x = c.execute(q)
    conn.commit()
    c.close()
    conn.close()


def getPendingRecords():
    f = open("pendingStuff.txt","w+")
    strX = "Pending"
    q = f'''SELECT * FROM jobTask j where j.workDetail="Pending"'''
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    x = c.execute(q)
    valX = x.fetchall()
    for i in valX:
        strX = ", ".join(str(e) for e in i)
        f.write(strX)
        f.write('\n')
    f.close()
    c.close()
    conn.close()


def getOldTable(val):
    q = f'''SELECT * FROM jobTask j where j.jobNumber={val}'''
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    x = c.execute(q)
    valX = x.fetchall()[0]
    valX = list(valX)
    print(valX)
    # conv = lambda i : i or '' 
    valX = ['' if v is None else v for v in valX]
    print(valX)
    sheet['E8'] = valX[1]
    '''ChangeLine'''
    sheet['F8'] = "ANJ-"+str(valX[0])
    sheet['C7'] = valX[2]
    sheet['C8'] = valX[3]
    sheet['C9'] = valX[4]
    sheet['C10'] = valX[5]
    sheet['C11'] = valX[6]
    sheet['C13'] = valX[7]
    sheet['C14'] = valX[8]
    sheet['C15'] = valX[9]
    sheet['B16'] = f"We are not Responsible for any data losses. Minimum diagnostic charges are: {valX[10]}"
    sheet['E10'] = valX[10]
    sheet['F11'] = valX[11]
    sheet['F10'] = valX[12]
    c.close()
    conn.close()

def createNewTable():
    orderDate = today.strftime("%d/%m/%Y")
    print(orderDate)
    # sheet['F8'] = orderDate
    q = '''SELECT MAX(jobNumber) FROM jobTask '''
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    x = c.execute(q)
    orderNumber = x.fetchall()[0][0]+1
    # ord
    print(orderNumber)
    # orderNumber = 1
    sheet['E8'] = orderDate
    '''changeLine'''
    sheet['F8'] = "ANJ-"+str(orderNumber)
    ClientPhone = input("Client Whatsapp Number: ")
    qX = f'''SELECT clientNumber,clientName,clientEmail from Profile p where p.clientNumber = {ClientPhone}'''
    x = conn.execute(qX)
    valZ = x.fetchall()
    print(valZ)
    # print(len(valZ))
    if len(valZ)==0:
        ClientName = input("*Client Name: ")
        sheet['C7'] = ClientName
        sheet['C8'] = ClientPhone
        ClientEmail = input("Client Email: ")
        sheet['C9'] = ClientEmail
        getMax = '''SELECT MAX(userNumber) FROM Profile'''
        xnew = conn.execute(getMax)
        num = int(xnew.fetchall()[0][0])+1
        q = f'''INSERT INTO Profile (userNumber,clientName,clientNumber,clientEmail) VALUES (?,?,?,?)'''
        vals = (num,ClientName,ClientPhone,ClientEmail)
        x = conn.execute(q,vals)
    else:
        valZ = list(valZ[0])
        num = valZ[0]
        ClientName = valZ[1]
        sheet['C7'] = ClientName
        ClientEmail = valZ[2]
        sheet['C9'] = ClientEmail
        sheet['C8'] = ClientPhone

    ItemRcvd = input("*Item Recieved: ")
    sheet['C10'] = ItemRcvd
    ModelNumber = input("*Model Number: ")
    sheet['C11'] = ModelNumber
    ProblemReported = input("*Problem Reported: ")
    sheet['C13'] = ProblemReported
    ProblemDiagnosed = input("Problem Diagnosed: ")
    sheet['C14'] = ProblemDiagnosed
    AdditionalComments = input("Additional Comments: ")
    sheet['C15'] = AdditionalComments
    MinimumCharges = input("Minimum Charges: ")
    sheet['E10'] = MinimumCharges
    sheet['B16'] = f"We are not Responsible for any data losses. Minimum diagnostic charges are: {MinimumCharges}"
    sheet['F11'] = "Pending"
    advanceRecieved = int(input("Advance Recieved: "))
    sheet['F10'] = advanceRecieved
    q = f'''INSERT INTO jobTask (jobNumber, orderDate, clientName, clientNumber, clientEmail, itemRecieved, modelNumber, ProblemReported, ProblemDiagnosed, AdditionalComments,minimumCharges,workDetail,advanceRecieved) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)'''
    vals = (orderNumber,orderDate,ClientName,ClientPhone,ClientEmail,ItemRcvd,ModelNumber,ProblemReported,ProblemDiagnosed,AdditionalComments,int(MinimumCharges),"Pending",advanceRecieved)
    x = c.execute(q,vals)
    relationship= f'''INSERT INTO jobTask_Profile_Relationship (jobNumber,userNumber) VALUES(?,?)'''
    vals = (orderNumber,num)
    x = conn.execute(relationship,vals)
    conn.commit()
    c.close()
    conn.close()
    phoneNum = '91'+ str(ClientPhone)
    '''sendMessage'''
    message = '''Dear {}
Thanks for visiting *Computer City India*
We have assigned Job Number: *{}* for the item: *{}* on {}
Minimum charges for the same is Rs.{} and we have recieved advance of Rs.{}
We will update you timely as per progress of this job
'''.format(ClientName,"ANJ-"+str(orderNumber),ItemRcvd,orderDate, MinimumCharges, advanceRecieved)

    link = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(phoneNum)
    #driver  = webdriver.Chrome()
    driver.get(link)
    print("Sending message to", phoneNum[2:])
    send_unsaved_contact_message(message)
    # send_unsaved_contact_message()

def send_unsaved_contact_message(message,num):
    # global message
    global driver
    try:
        time.sleep(7)
        input_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        for ch in message:
            if ch == "\n":
                ActionChains(driver).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
            else:
                input_box.send_keys(ch)
        # input_box.send_keys(Keys.ENTER)
        print("Message sent successfuly")
        return
    except:
        # print(e.)
        print("Failed to send message, retrying")
        if num==6:
            return
        else:
            send_unsaved_contact_message(message,num+1)
            return    
    return
    
file = 'template.xlsx'
rb = load_workbook(file)
# r_sheet = rb.sheet_by_index(0)
sheet = rb.active
chrome_options = Options()
chrome_options.add_argument('no-sandbox')
# chrome_options.add_argument('--headless')
chrome_options.add_argument("--user-data-dir=chrome-data")
driver = webdriver.Chrome(r'chromedriver.exe',options=chrome_options)
driver.get('https://web.whatsapp.com')
wait = WebDriverWait(driver, 600)
# wb = copy(rb)
# wb_sheet = wb.get_sheet(0)
while True:
    queryA = input('''Enter 1 for new file Query
    Enter 2 for old file Query
    Enter 3 for Deleting Record
    Enter 4 for Pending Records
    Enter 5 for Editing Record
    Enter 6 for Sending Update:   ''')

    if queryA=='1':
        createNewTable()
        rb.save("newBill.xlsx")
        subprocess.Popen("newBill.xlsx",shell=True)
    elif queryA=='2':
        a = int(input("Enter jobNumber: "))
        getOldTable(a)
        rb.save("newBill.xlsx")
        subprocess.Popen("newBill.xlsx",shell=True)
    elif queryA=='3':
        a = int(input("Enter jobNumber: "))
        deleteARecord(a)
        print("record Deleted")
    elif queryA=='4':
        getPendingRecords()
        print("All records in the text file")
        subprocess.Popen("pendingStuff.txt",shell=True)
    elif queryA=='5':
        a = int(input("Enter jobNumber: "))
        print("**Leave the field empty in case you dont want to update it.**")
        editRecord(a)
        rb.save("newBill.xlsx")
        subprocess.Popen("newBill.xlsx",shell=True)
    elif queryA=='6':
        val = int(input("Enter jobNumber: "))
        
        q = f'''SELECT * FROM jobTask j where j.jobNumber={val}'''
        conn = sqlite3.connect('cciAutomation.db')
        c = conn.cursor()
        x = c.execute(q)
        valX = x.fetchall()[0]
        valX = list(valX)
        print(valX)
        id = "ANJ-"+ str(valX[0])
        phoneNum = '91'+valX[3]
        print(id,phoneNum)
        b = input("Enter the Message to be sent: ")

        message = '''Update from *Computer City India*
Your job- {} is under Process 
{}
'''.format(id,b)

        link = "https://web.whatsapp.com/send?phone={}&text&source&data&app_absent".format(phoneNum)
        #driver  = webdriver.Chrome()
        driver.get(link)
        print("Sending message to", phoneNum[2:])
        send_unsaved_contact_message(message,0)