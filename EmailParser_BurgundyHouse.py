import imaplib
from lib2to3.pgen2.parse import ParseError
import smtplib
import email 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pyodbc

email_id = '' #emails to read from
reciever_email_id = "" #email id to send email
password = '' #password of email to read
subject = "Email Processing Hackathon" # subject line to filter
table_name = "Email_Processing_Hackathon" # database table name to operate with

def generateBody(records):
    html = ""
    for row in records:
        html += "<tr>"
        for data in row[1:]:
            if data is None:
                html += "<td> </td>"
            else:
                html += "<td>"+data+"</td>"
        html += "</tr>"
    return html

def generateHtml(records):
    return """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8" />
        <style type="text/css">
        table {
            background: white;
            border-radius:3px;
            border-collapse: collapse;
            height: auto;
            padding:5px;
            width: 100%;
            animation: float 5s infinite;
        }
        
        tr {
            border-top: 1px solid #C1C3D1;
            border-bottom: 1px solid #C1C3D1;
            border-left: 1px solid #C1C3D1;
            color:#666B85;
            font-size:16px;
            font-weight:normal;
        }
        tr:hover td {
            background:#4E5066;
            color:#FFFFFF;
            border-top: 1px solid #22262e;
        }
        td {
            background:#FFFFFF;
            padding:10px;
            text-align:left;
            vertical-align:middle;
            font-weight:300;
            font-size:13px;
            border-right: 1px solid #C1C3D1;
        }
        div{
            overflow:auto;
            max-width:900px;
        }
        </style>
    </head>
    <body>
        <div>
            <table>
            """ + generateBody(records) + """
            </table>   
        </div 
    </body>
    </html>
    """

def fetchEmailRecords():
    # Fetch records from db into parsedMails
    records = []
    conn = connectDB()
    cursor = conn.cursor()
    cursor.execute("select * from " + table_name+"")
    records.append([column[0] for column in cursor.description])
    for i in cursor:
        records.append(i)
    return records

def imapLogin():
    imapurl = 'imap.gmail.com' 
    connection = imaplib.IMAP4_SSL(imapurl)
    connection.login(email_id, password)
    connection.select('Inbox')
    return connection

def smtpLogin():
    s = smtplib.SMTP('smtp.gmail.com', 587)  
    s.starttls()
    s.login(email_id, password) 
    return s

def _generate_message(records) -> MIMEMultipart:
    message = MIMEMultipart("alternative", None, [MIMEText(generateHtml(records), 'html')])
    message['Subject'] = subject
    message['From'] = email_id
    message['To'] = reciever_email_id
    return message

def send_message():
    records = fetchEmailRecords()
    message = _generate_message(records)
    server = smtplib.SMTP("smtp.gmail.com:587")
    server.ehlo()
    server.starttls()
    server.login(email_id, password)
    server.sendmail(email_id, reciever_email_id, message.as_string())
    server.quit()

def isColumnPresent(col):
    conn = connectDB()
    cursor = conn.cursor()
    cursor.execute("select col_length('"+table_name+"','"+col+"')")
    for i in cursor:
        if(i[0]==None):
            return False
    return True

def addColumn(col):
    conn = connectDB()
    cursor = conn.cursor()
    cursor.execute('alter table '+table_name + ' add "'+col+'" varchar(100)')
    cursor.commit()


def updateValue(col, value):
    conn = connectDB()
    cursor = conn.cursor()
    cursor.execute("update top(1) "+table_name+ " set "+'"'+col+'"'+" = '"+value+"' where " + '"'+col+'"'+" is null")
    cursor.commit()
    return cursor.rowcount

def insertRow(col, value):
    conn = connectDB()
    cursor = conn.cursor()
    cursor.execute("insert into "+table_name+' ("'+col+'") values '+"('"+value+"')")
    cursor.commit()

def addValue(col, value):
    if updateValue(col, value) == 0:
        insertRow(col, value)

def parseMails():
    connection = imapLogin()
    _,mails = connection.search(None, '(SUBJECT "' + subject+'" UNSEEN)')

    for num in mails[0].split():
        _, data = connection.fetch(num, '(RFC822)')
        _, bytes_data = data[0]
        msg = email.message_from_bytes(bytes_data)
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                message = part.get_payload(decode=True).decode()
                for line in message.split('\r\n')[:-1]:
                    recordKey, recordValue = line.split(':')
                   
                    if not isColumnPresent(recordKey):
                        addColumn(recordKey)
                    addValue(recordKey, recordValue)
                
def connectDB():
    conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=;'
                      'Database=;' 
                      'UID=;'
                      'PWD=;'                     
                      'Trusted_Connection=No;')
    return conn

if __name__=="__main__":
    parseMails()
    send_message()
   