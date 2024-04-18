import pandas as pd
import tkinter as tk
from tkinter import filedialog
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import os
root = tk.Tk()
def open_file():
    global file_path
    global df
    button = tk.Button(root, text="Open File", command=open_file)
    button.grid(row=2)

    file_path = filedialog.askopenfilename()
    df=pd.read_excel(file_path)
open_file()
num_rows=len(df)
print("Emails to send are-",num_rows)
def search_file(file_name,search_path):
    file_paths=[]
    for abc, dirs ,files in os.walk(search_path):
        for file in files:
                if file==file_name:
                  file_paths.append(os.path.join(abc,file))
    return file_paths
def send_mail():
    for i in range(num_rows):
        pdf=df.iloc[i,0]
        all_rec=df.iloc[i,1]
        all_cc=df.iloc[i,2]
        body=df.iloc[i,3]
        file_name=pdf
        global pdf_file_path
        search_path=os.getcwd()
        file_paths=search_file(file_name,search_path)
        if file_paths:
            global pdf_file_path
            for pdf_file_path in file_paths:
                print(pdf_file_path)
        else:
            print("the file was not found")
        global pdf_msg
        pdf_msg = MIMEMultipart()
        with open(pdf_file_path, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="pdf")
            attachment.add_header('Content-Disposition', 'attachment', filename=pdf)
            pdf_msg.attach(attachment)
        pdf_msg.attach(MIMEText(body))

        from email.message import EmailMessage

        EMAIL_ADDRESS = 'youremailid'
        EMAIL_PASSWORD = "your app password"
        msg = EmailMessage()
        msg['Subject'] = 'Monthly Report'
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = all_rec
        msg['CC'] = all_cc
        msg.set_content(pdf_msg)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)

        print("SENT Email",(i+1),"of",num_rows)
send_mail()
print("ALL SENT")
root.mainloop()
