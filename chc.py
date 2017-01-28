import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from xlrd import open_workbook
from datetime import *
from tzlocal import *
import datetime as dt
import pytz


def TimeNow():
    now = dt.datetime.utcnow()
    time = datetime(now.year, now.month, now.day, now.hour, now.minute, now.second, tzinfo=pytz.utc)
    Zonetime = time.astimezone(get_localzone())
    return str(Zonetime)

file1=open("C:\Users\ChandrashekarChary\Desktop\logggg.txt","a")

log = "\n"+"LOG: " +TimeNow() + " Log file:" + "\t" + "Session Started" + "\n"
file1.write(log)
wb = open_workbook('C:\Users\ChandrashekarChary\Desktop\data.xlsx')
log = "LOG: " +TimeNow() + " Log file:" + "\t" + "Opening Data File(emails)" + "\n"
file1.write(log)
try:

    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        myemails = []
        rows = []
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                try:
                    value = str(int(value))
                except ValueError:
                    pass
                finally:
                    values.append(value)
            item = str(*values)
            myemails.append(item)
except:
    log = "LOG: " + TimeNow() + " Log file:" + "\t" + "Error reading DataFile" + "\n"
    file1.write(log)
print myemails[1]

fromaddr = "chandhu4004@gmail.com"
toaddr = "chandrashekarchary44@gmail.com"
msg = MIMEMultipart()
msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "SUBJECT OF THE MAIL"


body = "YOUR MEcdcdcdcdcd"
msg.attach(MIMEText(body, 'plain'))

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()

server.login(fromaddr, "chandhu9573331682")
log = "LOG: " +TimeNow() + " Log file:" + "\t" + "Successfully Logged in By email: " +str(fromaddr)+ "\n"
file1.write(log)
text = msg.as_string()
i=0
try:
    while(len(myemails)>0):
        server.sendmail(fromaddr, str(myemails[i]), text)
        log = "LOG: " + TimeNow() + " Log file:" + "\t" + "Message Sent To:  "+str(myemails[i]) + "\n"
        file1.write(log)
        i=i+1
except:
    pass
log = "LOG: " + TimeNow() + " Log file:" + "\t" + "Message Sent To:  " + str(len(myemails)) +" Email address(s)"+ "\n"
file1.write(log)
server.quit()