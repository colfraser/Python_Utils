#-- Colin Fraser 2022/04/21 Task 67451
#-- Compares XML output with SQL returned from two stored procedures.
#

import hashlib, sqlite3, pandas as pd
from xml.etree import cElementTree as ET

import io

import time
import datetime

import copy
import pyodbc
import sys
import os
import datetime
import time

from datetime import timedelta, date
from sqlalchemy import create_engine

def openSQLALCHEMYmemory():
    global cnxnAM
    print ("Open SQLALCHEMY Azure memory")
    # Code to use sqlalchemy
    engine = create_engine('sqlite://')
    cnxnAM = engine.raw_connection()   # engine.connect()
    cursorM = cnxnAM.cursor()
    cursorM.execute("create table summary (OrgID INTEGER,OrgNameWithCode TEXT, TotalOrders INTEGER, TotalTickets INTEGER, TotalFees FLOAT)")

def openERS():
    global cnxn1, cnxn2, cnxn4, cnxn5, cnxn6, cnxnX
    print ("Open Prod DB")
    server = ''
    database = ''
    authentication = 'ActiveDirectoryInteractive'  #MFA  'ActiveDirectoryPassword'
    username = ''
    driver= '{ODBC Driver 17 for SQL Server}'
    cnxn1 = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';Authentication='+authentication+';DATABASE='+database+';UID='+username)

    database = ''
    cnxn2 = pyodbc.connect('DRIVER='+driver+';SERVER='+server+';Authentication='+authentication+';DATABASE='+database+';UID='+username)
 
  

def ReadProc():
    
    cursor = cnxnX.cursor()
    cursorM = cnxnAM.cursor()
    
    cursor.execute('EXECUTE [DataExtract].[tablename]    @InvoiceStartDate = ?  ,@InvoiceEndDate = ?', (InvoiceStartDate,InvoiceEndDate))

    lineno = 0

    for row in cursor.fetchall():
        lineno = lineno + 1
        #print('Line number ',lineno)
        #if lineno > 3:   # to test
        #    print("End")
        #    break
        OrgID = row[3]
        if lineno > 1:
            OrgID = row[3]
            Org_Name = row[0]
            TotalOrders = row[4]
            TotalTickets = row[7]
            TotalFees = float(row[8])
            #print(OrgID,' - ',Org_Name,' - ',TotalOrders)
            #if  OrgID == 11507  or OrgID == 42100  or OrgID ==  2:     #to test.   End Date (2021,12,12) P2208 +  OrgID == 11507 had the difference  -- ERS1
                  # ERS2
            cursorM.execute('INSERT into summary (OrgID ,OrgNameWithCode, TotalOrders,TotalTickets,TotalFees) values (?,?,?,?,?)', (OrgID,Org_Name,TotalOrders,TotalTickets,TotalFees))
    
    cnxnAM.commit()  # Uses a temporary table in a different database as only one cursor can be opened per DB. Uses SQLALCHEMY as SQLLite can not do floats.
    cursor.close()

    cursorM.execute('select OrgID ,OrgNameWithCode, TotalOrders,TotalTickets,TotalFees from summary')
    readno = 0
 
    for row2 in cursorM.fetchall():
        readno = readno + 1
        InvoiceEndDate2 = InvoiceEndDate - timedelta(days=1) #  date(2021,12,11)
        #print('\n','InvoiceEndDate2 ',InvoiceEndDate2)
        Total_from_E025 = row2[2] 
        Total_Tickets_from_E025 = row2[3] 
        Total_Fees_from_E025 = float(row2[4])
        Description_from_E025 = row2[1]
        Agency =  row2[0] # 11507
        IncludeAll = 1
        Summary = 1
        #print('Processing ',Description_from_E025, ' ', InvoiceEndDate2, ' total:',Total_from_E025,' - ',Total_Tickets_from_E025 ,' - ', Total_Fees_from_E025 )  # To show progress
        sql_str = "EXECUTE  [dbo].[ev_GetTicketIssueFees]   @InvoiceEndDate = ?, @Agency = ?   ,@IncludeAll = ?,  @Summary = ?"
        df_ER2 = pd.read_sql(sql_str, cnxnX, params=[InvoiceEndDate2,Agency,IncludeAll,Summary])
        rtn_string = df_ER2.to_string()
        index = rtn_string.find('<') 
        rtn_string = rtn_string[index:]
        f = io.StringIO(rtn_string)
        tree = ET.parse(f)
        root = tree.getroot()

        total = 0
        total_Tickets = 0
        total_Fees = 0
        for field in root.findall('TicketIssueCountCollection'):
            for line in field:
                total = total + int(line.attrib['OrderCount'])
                total_Tickets = total_Tickets + int(line.attrib['TicketCount'])
        for field in root.findall('TicketIssueFeeCollection'):
            for line in field:
                total_Fees = total_Fees + (float(line.attrib['Fee'])/100)
        total_Fees = round(total_Fees,2)  # some issue accumulating the fees accurately 
  
        if Total_from_E025 != total or Total_Tickets_from_E025 != total_Tickets or Total_Fees_from_E025 != total_Fees:
            print('For ',Description_from_E025,' there is a difference. Total Orders', total, ' Total Tickets',total_Tickets,' Total Fees ', total_Fees)
            print('E025                                             total orders:',Total_from_E025,' - ',Total_Tickets_from_E025 ,' - ', Total_Fees_from_E025 ) 
      

        sqlite_update_query = """DELETE from summary"""
        cursorM.execute(sqlite_update_query)
        cnxnAM.commit()


def loopRailPeriods():
        global InvoiceStartDate
        global InvoiceEndDate
        cursor1 = cnxnX.cursor()
        RSPPeriod =    2212
        sql_str = "SELECT TOP 1 InvoiceEndDate, InvoiceStartDate, RSPPeriod FROM hwtInvoiceEndDate 	WHERE hwtAccountingPeriodID = 4 AND RSPPeriod > ? ORDER BY RSPPeriod "
        cursor1.execute(sql_str,RSPPeriod)
        #time.sleep(1)
        rows = cursor1.fetchall()
        for row in rows:
            InvoiceEndDate = row[0] + timedelta(days=1)
            InvoiceStartDate = row[1]
            RSPPeriod = row[2]
        while RSPPeriod <= str(2213):  
            if cursor1.rowcount == 0 :
                print("End")
                break

            print('\n','\n',InvoiceStartDate, " ",InvoiceEndDate," ",RSPPeriod)
            ReadProc()
            cursor1.execute(sql_str,RSPPeriod)
            rows = cursor1.fetchall()
            for row in rows:
                InvoiceEndDate = row[0] + timedelta(days=1)
                InvoiceStartDate = row[1]
                RSPPeriod = row[2]


print('XML vs SQL Details comparison for hard to find problem')
now = datetime.datetime.now()
print ("Current date and time : ",now.strftime("%Y-%m-%d %H:%M:%S"))


openERS()

openSQLALCHEMYmemory()

cnxnX = cnxn1
print('ERS1')
loopRailPeriods()

print('\n','\n','ERS2')
cnxnX = cnxn2
loopRailPeriods()




now = datetime.datetime.now()
print ("End date and time : ",now.strftime("%Y-%m-%d %H:%M:%S"))
