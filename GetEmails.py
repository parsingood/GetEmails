
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
#import zipfile
import EmailTools

temp_path = "e:/test/att/"
temp_path_ = "e:\\test\\att\\"
MailsBack = 10  # how many email will re-ckech before Mail_TopID - in case of move-old, some new-mails, move-back-old
PRO_connection_string = "Driver={SQL Server Native Client 11.0};Server=PARSING01\SQLEXPRESS;Database=Parsing;Trusted_Connection=yes;"

f = open("c:\parsing\connection_string.txt")
connection_string =  f.read()  
f.close
connChannels = pyodbc.connect(connection_string)
cursorChannels = connChannels.cursor()
cursorChannels.execute('SELECT * FROM Channels where IsActive > 0 order by ChannelID')

for row in cursorChannels:
    
    if bool(row.MailFolder):
        con = EmailTools.auth(row.MailServer,row.UserName,row.Password,row.MailFolder)
    else:
        con = EmailTools.auth(row.MailServer,row.UserName,row.Password)

    isOk, messages = con.search(None, 'ALL')
    ids = messages[0]  # data is a list.
    id_list = ids.split()  # ids is a space separated string
    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()

    PRO_conn = None
    PRO_cursor = None
    if row.IsActive == 2:
        PRO_conn = pyodbc.connect(PRO_connection_string)
        PRO_cursor = PRO_conn.cursor()        

 
    MailsBack = row.MailsBack if  hasattr(row, 'property') and  row.MailsBack else MailsBack
    MailsBack = row.MailsBack if  row.MailsBack else MailsBack
    start_id = max( [ min( [max(int(x) for x in id_list), row.Mail_TopID ] ) - MailsBack if row.Mail_TopID else  0  , 0 ])
    new_id_list = [x for x in id_list if int(x) >= start_id ]

    for id in new_id_list:
        isOKtype, data = con.fetch(id,'(RFC822.HEADER)')
        raw = email.message_from_bytes(data[0][1])

        isNewEmail = True
        if  raw['Message-Id']:
            check_conn = pyodbc.connect(connection_string)
            check_cursor = check_conn.cursor()
            hdr = email.header.make_header(email.header.decode_header(raw['Message-Id']))
            MessageID = str(hdr)
            check_cursor.execute("select count(*) cnt  from E_Folders where PMS_ServerID = ?  and FolderName = ? and MessageID = ?  "
                                                , row.PMS_ServerID, row.FolderName, MessageID)
            countExists=check_cursor.fetchone().cnt
            isNewEmail = countExists == 0
            
        else:

            check_conn = pyodbc.connect(connection_string)
            check_cursor = check_conn.cursor()
            this_date_tuple = email.utils.parsedate_tz(raw['Date'])
            this_FolderTime=''
            if this_date_tuple:
                this_local_date = datetime.datetime.fromtimestamp( email.utils.mktime_tz(this_date_tuple))
                this_FolderTime=this_local_date.strftime("%Y-%m-%d-%H-%M-%S")
            
            this_Subject = ''
            if raw['Subject']:
                this_hdr = email.header.make_header(email.header.decode_header(raw['Subject'])) 
                this_Subject = str(this_hdr)
            check_cursor.execute("select count(*) cnt  from E_Folders where PMS_ServerID = ?  and FolderName = ? and Subject = ? and FolderTime = ?  "
                                            , row.PMS_ServerID, row.FolderName, this_Subject, this_FolderTime )
            countExists=check_cursor.fetchone().cnt
            isNewEmail = countExists == 0

        if isNewEmail :  # if email is new:

            isOKtype, data = con.fetch(id,'RFC822')
            raw = email.message_from_bytes(data[0][1])
            charsets=raw.get_charsets() 

            body = EmailTools.get_body_decode(raw)

            date_tuple = email.utils.parsedate_tz(raw['Date'])
            FolderTime=''
            if date_tuple:
                local_date = datetime.datetime.fromtimestamp( email.utils.mktime_tz(date_tuple))
                FolderTime=local_date.strftime("%Y-%m-%d-%H-%M-%S")
            
            Subject = ''
            if raw['Subject']:
                hdr = email.header.make_header(email.header.decode_header(raw['Subject'])) 
                Subject = str(hdr)

            hdr = email.header.make_header(email.header.decode_header(raw['From']))
            FromAddress = str(hdr)

            MessageID =''
            if raw['Message-Id']:
                hdr = email.header.make_header(email.header.decode_header(raw['Message-Id']))
                MessageID = str(hdr)


            s = 'insert into E_Folders (PMS_ServerID,FolderName,FolderTime,Subject,FromAddress,MessageID,BodyTXT)' \
            + ' select ?, ?, ?, ?, ?, ?, ? ' 
            cursor.execute(s,  str(row.PMS_ServerID) , row.FolderName , FolderTime , Subject , FromAddress , MessageID, body  )
            cursor.execute("  select @@IDENTITY FolderID ")
            recs=cursor.fetchall()
            FolderID = 0
            if len(recs) > 0:
                FolderID=int(recs[0].FolderID)
            cursor.commit()

            PRO_FolderID = 0
            if  PRO_cursor :
                s = 'insert into E_Folders (PMS_ServerID,FolderName,FolderTime,Subject,FromAddress,BodyTXT)' \
                + ' select ?, ?, ?, ?, ?, ? ' 
                PRO_cursor.execute(s,  str(row.PMS_ServerID) , row.FolderName , FolderTime , Subject , FromAddress ,  body  )
                PRO_cursor.execute("  select @@IDENTITY FolderID ")
                recs=PRO_cursor.fetchall()
                PRO_FolderID = 0
                if len(recs) > 0:
                    PRO_FolderID=int(recs[0].FolderID)
                PRO_cursor.commit()


            for part in raw.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get_content_type() == 'text/html' :
                    FileText = part.get_payload(decode=True)
                    charset = part.get_content_charset('utf-8')
                    try:
                        FileTextString = FileText.decode(charset, 'replace')
                    except:
                        try:
                            FileTextString = FileText.decode('latin1')
                        except:
                            FileTextString = "Error"
                    s = 'update E_Folders set [BodyHTML] = ? where [FolderID] = ? ' 
                    if  PRO_cursor :
                        PRO_cursor.execute(s, FileTextString , PRO_FolderID )
                        PRO_cursor.commit()
 #                   EmailTools.insertFolder(cursor, PRO_cursor, FolderID, PRO_FolderID, 'body.html', '.HTML', FileTextString)
                    continue

                if part.get('Content-Disposition') is None:
                    continue
                fileName = part.get_filename()        
                if bool(fileName):

                    fileName=EmailTools.encoded_words_to_text(fileName)
                    fileName=fileName.replace("'","`")
                    FileText = part.get_payload(decode=True)

                    # Classify attachments, convert, insert in cursor
                    EmailTools.Add_FileText_to_Folder(cursor,PRO_cursor,FolderID,PRO_FolderID,fileName,FileText)


        updateChannels_conn = pyodbc.connect(connection_string)
        updateChannels_cursor = updateChannels_conn.cursor()
        updateChannels_cursor.execute("update Channels set Mail_TopID = ? where ChannelID = ? " , str(int(id)) , row.ChannelID)
        updateChannels_cursor.commit()


















