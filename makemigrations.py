import time
import datetime
import os
import argparse
import sqlite3
from datetime import date
import schedule
today = date.today()
conn = sqlite3.connect('cciAutomation.db')

ProfileTable = '''CREATE TABLE Profile (
    userNumber   VARCHAR (200) PRIMARY KEY,
    clientName   VARCHAR       NOT NULL,
    clientNumber VARCHAR       UNIQUE
                               NOT NULL,
    clientEmail  STRING        NOT NULL
                               UNIQUE
);'''

jobTask_Profile_RelationshipTable = '''CREATE TABLE jobTask_Profile_Relationship (
    jobNumber,
    userNumber
);
'''
c = conn.execute(ProfileTable)
c2 = conn.execute(jobTask_Profile_RelationshipTable)
# conn.commit()
# conn = sqlite3.connect('cciAutomation.db')
queryGetUsers = '''SELECT jobNumber, clientName,clientNumber,clientEmail FROM jobTask'''
# q = f'''SELECT * FROM jobTask j where j.jobNumber={val}'''
x = conn.execute(queryGetUsers)
valX = x.fetchall()
for i in valX:
    # print(i)
    # print(i[1])
    qX = f'''SELECT userNumber from Profile p where p.clientNumber = {i[2]}'''
    x = conn.execute(qX)
    valZ = x.fetchall()
    print(valZ)
    # print(len(valZ))
    if len(valZ)==0:
        getMax = '''SELECT MAX(userNumber) FROM Profile'''
        xnew = conn.execute(getMax)
        try:
            num = int(xnew.fetchall()[0][0])+1
        except:
            num = 0
        q = f'''INSERT INTO Profile (userNumber,clientName,clientNumber,clientEmail) VALUES (?,?,?,?)'''
        vals = (num,i[1],i[2],i[3])
        x = conn.execute(q,vals)
        relationship= f'''INSERT INTO jobTask_Profile_Relationship (jobNumber,userNumber) VALUES(?,?)'''
        vals = (i[0],num)
        x = conn.execute(relationship,vals)
    else:
        num = valZ[0][0]
        relationship= f'''INSERT INTO jobTask_Profile_Relationship (jobNumber,userNumber) VALUES(?,?)'''
        vals = (i[0],num)
        x = conn.execute(relationship,vals)
conn.commit()
    