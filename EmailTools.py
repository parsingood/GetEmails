import DocsToText
import imaplib
import os
import io
import zipfile

def encoded_words_to_text(encoded_words):
    try:
        encoded_word_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='
        charset, encoding, encoded_text = re.match(encoded_word_regex, encoded_words).groups()
        if encoding is 'B':
            byte_string = base64.b64decode(encoded_text)
        elif encoding is 'Q':
            byte_string = quopri.decodestring(encoded_text)
        return byte_string.decode(charset)
    except:
        return encoded_words


def auth(mailserver, username, password, folder='Inbox'):
    #con = imaplib.IMAP4_SSL(mailserver)
    con = imaplib.IMAP4(mailserver)
    con.login(username, password)
    con.select(folder)
    return con


def get_body_decode(msg):
    if msg.is_multipart():
        return get_body_decode(msg.get_payload(0))
    else:
        body_from_payload = msg.get_payload(None,True)
        try:
            return msg.body_from_payload.decode(msg.get_charsets()[0])
        except:
            #rawdata = open(file, "r").read()
            #result = chardet.detect(body_from_payload)
            #charenc = result['encoding']

            try:
                return body_from_payload.decode(chardet.detect(body_from_payload)['encoding']) #'utf-8'
            except:
                return body_from_payload.decode('latin1')

           


            #try:
            #    return body_from_payload.decode(chardet.detect(body_from_payload)) #'utf-8'
            #except:
            #    return body_from_payload.decode('latin1')



temp_path = "e:/test/att/"
Attachments_path = 'e:/EATT/'
def Add_FileText_to_Folder(cursor,PRO_cursor,FolderID,PRO_FolderID,fileName,FileText):
    
    if PRO_FolderID !=0 :
        os.makedirs(Attachments_path + str(PRO_FolderID), exist_ok=True) 
        f = open(Attachments_path + str(PRO_FolderID) + '/' + fileName, 'wb')
        f.write(FileText)
        f.close()

    FileTextString = ""
    filenameOnly, file_extension = os.path.splitext(fileName)
    # ----> Begin of classifying the extentions
    if file_extension.lower() in ('.png','.jpg','.gif','.bmp','.jpeg'):
        return

    zipPath = ''
    if file_extension.lower()[-3:] == "zip" : zipPath = 'zip/'
    if file_extension.lower()[-2:] == "z7"  : zipPath = 'z7/'
    if file_extension.lower()[-3:] == "arj"  : zipPath = 'arj/'
    if zipPath != '' :
        my_file_io = io.BytesIO(FileText)
        with zipfile.ZipFile(my_file_io, 'r') as zip_ref:
            directory=temp_path + zipPath + str(FolderID) + '/' + fileName +  '/'
            zip_ref.extractall(directory)
            for current_filename in os.listdir(directory):
                f = open(os.path.join(directory, current_filename),'r')
                FileText = f.read()
                f.close()
                Add_FileText_to_Folder(cursor,PRO_cursor,FolderID,PRO_FolderID, filenameOnly + "~" + current_filename,FileText)
        return

    if file_extension.lower()[-3:] == "eml" :
        f = open(temp_path + 'eml/str(FolderID)/' + fileName, 'wb')
        f.write(FileText)
        f.close()
        return

    if file_extension.lower()[-3:]=="xls" or file_extension.lower()[-4:] in ("xlsm","xlsx")  :
        FileTextString = DocsToText.convert_xls_to_txt(FileText)

    if file_extension.lower()[-4:]=="docx":
        FileTextString = DocsToText.convert_docx_to_txt(FileText)   

    if file_extension.lower()[-3:]=="doc":
        FileTextString = DocsToText.convert_doc_to_txt(FileText)
 
    if file_extension.lower()[-3:]=="pdf":
        FileTextString = DocsToText.convert_pdf_to_txt(FileText)

    if FileTextString == "" and FileText :
        try: FileTextString = FileText.decode(chardet.detect(FileText)['encoding'])  #"utf-8"
        except: FileTextString = FileText.decode('latin1')

    # ----> end of classifying the extentions


    if FileTextString != "":
        s = 'insert into E_TXT (FolderID,TXT_Name,FileText) select ?, ?, ? '
        cursor.execute(s,FolderID,fileName, FileTextString)
        cursor.commit()

        if  PRO_cursor :
           s = "SELECT max(ATT_TypeID) TypeID FROM E_ATT_Types where ATT_Type = ? "
           PRO_cursor.execute(s, file_extension.upper() )
           ATT_TypeID = None
           for r in PRO_cursor: 
               ATT_TypeID = r.TypeID

           s = 'insert into E_TXT (FolderID,TXT_Name,FileText,ATT_TypeID) select ?, ?, ?, ?  '
           PRO_cursor.execute(s,PRO_FolderID,fileName, FileTextString, ATT_TypeID)
           PRO_cursor.commit()







