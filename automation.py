# from xlrd import *
from datetime import date
import sqlite3
import subprocess
from openpyxl import load_workbook
today = date.today()

# def queryOldTable():

# conn = sqlite3.connect('cciAutomation.db')
# c = conn.cursor()

# def updateOldTable():

def updateRecord(val):
    stringX = "Completed"
    q = f'''UPDATE jobTask SET workDetail = "Completed" WHERE jobNumber = {val}'''
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    x = c.execute(q)
    print(f"Updated Record number {val} to completed")
    conn.commit()



def deleteARecord(val):
    q= f''' DELETE FROM jobTask where jobNumber ={val}'''
    conn = sqlite3.connect('cciAutomation.db')
    c = conn.cursor()
    x = c.execute(q)
    conn.commit()





def getOldTable(val):
    updateRecordX = input("Do You want to update this record to complete (Y/N): ")
    if updateRecordX=='Y' or updateRecordX=='y':
        updateRecord(val)
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
    sheet['F8'] = "AKS-"+str(valX[0])
    sheet['C7'] = valX[2]
    sheet['C8'] = valX[3]
    sheet['C9'] = valX[4]
    sheet['C10'] = valX[5]
    sheet['C11'] = valX[6]
    sheet['C13'] = valX[7]
    sheet['C14'] = valX[8]
    sheet['C15'] = valX[9]
    sheet['B16'] = f"Minimum charges for the repair are: {valX[10]}"
    sheet['E10'] = valX[10]
    sheet['F11'] = valX[11]
    sheet['F10'] = valX[12]


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
    sheet['F8'] = "AKS-"+str(orderNumber)
    ClientName = input("*Client Name: ")
    sheet['C7'] = ClientName
    ClientPhone = input("Client Phone Number: ")
    sheet['C8'] = ClientPhone
    ClientEmail = input("Client Email: ")
    sheet['C9'] = ClientEmail
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
    sheet['B16'] = f"Minimum charges for the repair are: {MinimumCharges}"
    sheet['F11'] = "Pending"
    advanceRecieved = int(input("Advance Recieved: "))
    sheet['F10'] = advanceRecieved
    q = f'''INSERT INTO jobTask (jobNumber, orderDate, clientName, clientNumber, clientEmail, itemRecieved, modelNumber, ProblemReported, ProblemDiagnosed, AdditionalComments,minimumCharges,workDetail,advanceRecieved) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?1)'''
    vals = (orderNumber,orderDate,ClientName,ClientPhone,ClientEmail,ItemRcvd,ModelNumber,ProblemReported,ProblemDiagnosed,AdditionalComments,int(MinimumCharges),"Pending",advanceRecieved)
    x = c.execute(q,vals)
    conn.commit()


    
file = 'template.xlsx'
rb = load_workbook(file)
# r_sheet = rb.sheet_by_index(0)
sheet = rb.active

# wb = copy(rb)
# wb_sheet = wb.get_sheet(0)

queryA = input('''Enter 1 for new file Query
Enter 2 for old file Query
Enter 3 for Deleting Record: ''')

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

