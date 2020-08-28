############################ PROGRAMMER: S.MUKILAN #############################
import smtplib
import tkinter as tk
import time
import re
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from PIL import Image,ImageTk

WINDOW =tk.Tk()
WINDOW.geometry('700x500')
WINDOW.title('e-MAIL SENDER')
bg = Image.open('bg_sending.png')
bg = ImageTk.PhotoImage(bg)
bg_s = tk.Label(WINDOW,image=bg).place(x=0,y=0)
sening_mail = " "

def single_mail():
    global reciver_email,sening_mail,notify,error,message
    try:
        error.destroy()
    except:
        pass
    try:
        message.destroy()
    except:
        pass
    notify.destroy()
    sening_mail = "single"
    notify = tk.Label(WINDOW,bg='silver',text='* ENTER RECIPENT EMAIL ID')
    notify.place(x=380,y=205)
    reciver_email = tk.Entry(WINDOW, borderwidth=5)
    reciver_email.place(x=145,y=205,height=25,width=220)

def multiple_mail():
    global reciver_email,sening_mail,notify,error,message
    try:
        error.destroy()
    except:
        pass
    try:
        message.destroy()
    except:
        pass
    notify.destroy()
    sening_mail = "multiple"
    notify = tk.Label(WINDOW,bg="silver",text='* ENTER CSV FILE WITH [E-mail] COLOUM')
    notify.place(x=380,y=205)
    reciver_email = tk.Entry(WINDOW, borderwidth=5)
    reciver_email.place(x=145,y=205,height=25,width=220)

def send_mail():
    global user_email,password_user,reciver_email,sening_mailattach,notify,error,message   
    try:
        error.destroy()
    except:
        pass
    try:
        message.destroy()
    except:
        pass
    try:
        user_id = user_email.get()
        user_pass = password_user.get()
        to_add = reciver_email.get()
        subject = header.get()
        bdy = body.get()
        attachment = attach.get()
        
        try:
            server = smtplib.SMTP('smtp.gmail.com',587)
            server.starttls()
            network = True
            try:
                server.login(user_id,user_pass)
                msg = MIMEMultipart()
                msg['From'] = user_id
                msg['Subject'] = subject
                msg.attach(MIMEText(bdy,'plain'))
                user_pass = True
            except:
                error = tk.Label(WINDOW,bg='silver',fg='red',text='CHECK YOUR USER ID AND PASSWORD',font=(None,15))
                error.place(x=160,y=405)
                user_pass = False
                
            if user_pass == True:
                try:
                    attach_ment  = open(attachment,'rb')
                    part = MIMEBase('application','octet-stream')
                    part.set_payload((attach_ment).read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition',"attachment; filename= "+str(attachment))
                    msg.attach(part)
                    attach_err = False
                except:
                    if attachment == '':
                        attach_err = False
                    else:
                        error = tk.Label(WINDOW,bg='silver',fg='red',text='ATTACHMENT FILE DOES NOT EXIST',font=(None,15))
                        error.place(x=160,y=405)
                        attach_err = True
                        
            if sening_mail == "single" and attach_err != True:
                if user_pass == True:
                    try:
                        error.destroy()
                    except:
                        pass
                    msg['To'] = to_add
                    text = msg.as_string()
                    server.sendmail(user_id,to_add,text)
                    message = tk.Label(WINDOW,bg='silver',fg='green',text='.................MAIL SENDED.............',font=(None,15))
                    message.place(x=160,y=405)
                    server.quit()
            
            if sening_mail == "multiple":
                if user_pass == True:
                    try:
                        error.destroy()
                    except:
                        pass
                    try:
                        csv_file = pd.read_csv(to_add)
                        csv_req = True
                    except:
                        error = tk.Label(WINDOW,bg='silver',fg='red',text='     CSV FILE DOES NOT EXIST',font=(None,15))
                        error.place(x=160,y=405)
                        csv_req = False
                    if csv_req == True:
                        try:
                            for i in csv_file['E-mail']:
                                msg['To'] = i
                                text = msg.as_string()
                                server.sendmail(user_id,i,text)
                            message = tk.Label(WINDOW,bg='silver',fg='green',text='.................MAIL SENDED.............',font=(None,15))
                            message.place(x=160,y=405)    
                            server.quit()
                        except:
                            error = tk.Label(WINDOW,bg='silver',fg='red',text='CSV FILE DOES NOT CONTAIL [E-mail] COLOUM',font=(None,15))
                            error.place(x=160,y=405)
        except:
           
            if network != True:
                error = tk.Label(WINDOW,bg='silver',fg='red',text='CHECK YOUR INTERNET CONNECTION',font=(None,15))
                error.place(x=160,y=405)
    except:
        error = tk.Label(WINDOW,bg='silver',fg='blue',text='SELECT SINGLE (OR)  MULTIPLE MAIL',font=(None,15))
        error.place(x=160,y=405)
    

single = tk.Button(WINDOW, borderwidth=5, bg = 'black', fg='red',text='SINGLE MAIL', command=lambda:single_mail()).place(x=150,y=105, height=30,width=150)
multiple = tk.Button(WINDOW, borderwidth=5, bg = 'black', fg='red',text='MULTIPLE MAIL', command=lambda:multiple_mail()).place(x=400,y=105, height=30,width=150)
mail_send = tk.Button(WINDOW, borderwidth=5, bg = 'black', fg='red',text='SEND MAIL', command=lambda:send_mail()).place(x=160,y=465, height=30,width=400)

notify = tk.Label(WINDOW,bg='silver',text='* SELECT SINGLE (OR) MULTIPLE MAIL',font=(None,12))
notify.place(x=145,y=205)
attach_notify = tk.Label(WINDOW,bg='silver',text='* ENTER FILE NAME WITH EXTENSION',font=(None,9)).place(x=465,y=340)
user_email = tk.Entry(WINDOW, borderwidth=5)
user_email.place(x=145,y=155,height=25,width=220)
password_user = tk.Entry(WINDOW, borderwidth=5,show="*",font=15)
password_user.place(x=500,y=155,height=25,width=190)
header = tk.Entry(WINDOW, borderwidth=5)
header.place(x=145,y=245,height=25,width=320)
body = tk.Entry(WINDOW, borderwidth=5)
body.place(x=145,y=290,height=25,width=320)
attach = tk.Entry(WINDOW, borderwidth=5)
attach.place(x=145,y=340,height=25,width=320)

WINDOW.mainloop()
