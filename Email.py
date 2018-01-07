import smtplib
import xlrd
import datetime
import ctypes
today = datetime.date.today()
sender = input()
passward = input()
workbook = xlrd.open_workbook(r"C:\Users\HP\PycharmProjects\untitled\data.xlsx")
worksheet = workbook.sheet_by_name("Sheet1")
date = worksheet.col_values(0)
date = date[1:]
receiver = worksheet.col_values(1)
receiver = receiver[1:]
name = worksheet.col_values(2)
name = name[1:]
message = 'Subject: WISH\n\nHAPPY BIRTHDAY'
mail = smtplib.SMTP()
mail.connect('smtp.gmail.com', 587)
mail.ehlo()
mail.starttls()  # use to make secure connection
mail.login(sender,passward)
count=3
while(count>0) :
    for i in range(0,3) :
     if str(today)==date[i] :

            receive = receiver[i]
            n=name[i]
            mail.sendmail(sender,receive,message)
            ctypes.windll.user32.MessageBoxW(0, n, "HAPPY BIRTHDAY", 1)
            count-=1

mail.close()