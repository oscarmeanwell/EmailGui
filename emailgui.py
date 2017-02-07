from tkinter import *
import smtplib, sys, imaplib, email
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import os
root = Tk()
Password = ''
Email = ''
Server = ''
class FirstWindow():
    def __init__(self, parent):
        self.parent = parent
        self.parent.title('Log In')
        #self.parent.resizable(0,0)
        self.parent.grid()
        self.parent.geometry("330x230")
        self.parent.configure(background="steel blue")
        self.AddFirstWindow()
        self.ToArray = []
##        self.LogIn()
        self.FromArray = []
        self.SubjectArray = []
    def Clear(self, event):
        self.PasswdEnt.delete(0, END)
        
    def Clear1(self, event):
        self.UsrnEnt.delete(0, END)

    def LogIn(self):
        global Password, Email, Server 
        Password = self.PasswdEnt.get()
        Email = self.UsrnEnt.get()
        Server = smtplib.SMTP("smtp.gmail.com", 587)
        Server.ehlo()
        Server.starttls()
        try:
            Server.login(Email, Password)
            print("Login Succesful")
            self.parent.destroy()
            
        except:
            print("Login Failed")
        
    def AddFirstWindow(self):
        self.firstfrm = Frame(self.parent, width="500", height="500", bg="steel blue")
        self.firstfrm.grid(row=0, column=0)
        self.lbl1 = Label(self.firstfrm, text="Welcome, Please enter your details!", bg="steel blue", font=("Calibri", 12))
        self.lbl1.grid(row=0, column=0, padx=43, pady=3)
        self.UsrnEnt = Entry(self.firstfrm, bd =4, width=25, font=("Calibri", 10))
        self.PasswdEnt = Entry(self.firstfrm, bd =4,show = "*", width=25, font=("Calibri", 10))  
        self.UsrnEnt.insert(END, "Email Address")
        self.PasswdEnt.insert(END, "Password")
        self.UsrnEnt.bind("<FocusIn>", self.Clear1)
        self.PasswdEnt.bind("<FocusIn>", self.Clear)
        self.SubBtn = Button(self.firstfrm, text="Submit",height=2,width=10,fg="black", activebackground="yellow",bg="light blue", command=self.LogIn)
        self.warning = Label(self.firstfrm, text="Oscar Meanwell Software",  bg="steel blue")
        self.warning.grid(row=5, column=0, padx=43, pady=7)
        self.UsrnEnt.grid(row=1, column=0, padx=43, pady=10)
        self.PasswdEnt.grid(row=2, column=0, padx=43, pady=3)
        self.SubBtn.grid(row=3, column=0, padx=43, pady=7)
        

class SecondWindow():
    def __init__(self, parent):
        self.parent = parent
        self.parent.title('Your Inbox')
        #self.parent.resizable(0,0)
        self.parent.geometry("1300x800")
        self.parent.configure(background="steel blue")
        self.parent.grid()
        self.AttatchedFlag = False
        self.FileArray = []
        self.AddSecondWindow()
        self.gettemails()
        self.AddToEnt()
##        self.addemails()

    def AddSecondWindow(self):
        self.secondfrm = Frame(self.parent, width="4000", height="4000", bg="steel blue")
        self.secondfrm.grid(row=1, column=1)
        self.EntFrm = Frame(self.parent, width="400", height="400", bg="steel blue")
        self.EntFrm.grid(row=0, column=1)
        self.BtnFrm = Frame(self.parent, width="400", height = "400", bg="steel blue")
        self.BtnFrm.grid(row=0, column=0)
        self.lstfrm = Frame(self.parent, width="400", height="400", bg="steel blue")
        self.lstfrm.grid(row=1, column=0)
        self.LstBox = Listbox(self.lstfrm, height=38, width=30, selectmode = SINGLE, bd=4)
        self.LstBox.grid(row=1, column=0, padx = 20, pady = 20)
        self.FromEnt = Entry(self.EntFrm, width="140", bd=4)
        self.FromEnt.grid(row=0, column=2, pady=20)
        self.SubEnt = Entry(self.EntFrm, width="140", bd=4)
        self.SubEnt.grid(row=1, column=2)
        self.TextEnt = Text(self.secondfrm,width="120", height="38", bg="white", wrap=NONE, bd=4)
        self.TextEnt.grid(row=0, column=0, padx = 20, pady = 20, rowspan=3, sticky=(N))
        self.LstBox.bind('<<ListboxSelect>>', self.addemails)
        self.FromLbl = Label(self.EntFrm, text="From: ", bg='steel blue')
        self.FromLbl.grid(row=0, column=1, padx = 20)
        self.SubLbl = Label(self.EntFrm, text="Subject: ", bg='steel blue')
        self.SubLbl.grid(row=1, column=1, padx = 20)
        self.lstscrollbar = Scrollbar(self.secondfrm)
        self.lstscrollbar.grid(row=0, column=1, rowspan=3,pady=20, sticky='ns')
        self.TextEnt.config(yscrollcommand=self.lstscrollbar.set)
        self.lstscrollbar.config(command=self.TextEnt.yview)
        self.NewButton = Button(self.BtnFrm, text='Compose',height=2,width=8,fg='blue', command=self.SendMail, bd=4)# activebackground="yellow", command=Action)    
        self.NewButton.grid(row=0, column=0, padx=20,pady=20)
    def gettemails(self):
        server = imaplib.IMAP4_SSL('imap.gmail.com', 993)
        server.login(Email,Password)
        stat,cnt = server.select('Inbox')
        stat, dta = server.search(None, 'All')
        self.SubjectArray = []
        self.FullFromArray = []
        self.ToArray = []
        self.FromArray = []
        self.BodyDictionary = {}
        self.Counter = 0
        for i in dta[0].split():
            stat, dta = server.fetch((i), '(RFC822)')
            Message = email.message_from_bytes((dta[0][1]))
            if Message.is_multipart():
                for part in Message.walk():
                   self.filename =  part.get_filename() # get name of attachments
                   if bool(self.filename):# if there is a filename
                       try:
                           File = open(self.filename, 'wb')
                           File.write(part.get_payload(decode=True))
                           File.close()
                       except:
                           pass
                   if part.get_content_type() == "text/plain":
                        try:
                            self.Counter += 1
                            self.body = part.get_payload(decode=True)
                            self.body = self.body.decode()
                            self.EmailConcatinated = 'Email' + str(self.Counter)
                            self.ToAdd = ''
                            self.CharFlag = 0
                            for y in Message['From']:
                                if y in['<','>','[',']','=']:
                                    self.CharFlag = 1
                                if self.CharFlag != 1 and y != '"' and y != "'":
                                    self.ToAdd += y
                                else:
                                    pass
                            self.FirstCharFlag = 0
                            self.ToRem = ''
                            for x in Message['From']:
                                if x == ">":
                                    self.FirstCharFlag = 1

                                if self.FirstCharFlag != 1: 
                                    self.ToRem += x
                            self.ToRem += '>'
                            print(self.ToRem)
                            self.ToArray.append(Message['To'])
                            self.FromArray.append(self.ToAdd)
                            self.FullFromArray.append(self.ToRem)
                            self.SubjectArray.append(Message['Subject'])
                            self.BodyDictionary[self.EmailConcatinated] = self.body
                            
                

                        except:
                            self.Counter -= 1
                   else:
                        pass
        server.close()

    def AddToEnt(self):
        self.Max = len(self.FromArray) 
        self.LstCount = 1
        for x in range(self.Max, 0, -1):
            self.LstBox.insert(self.LstCount, ' ' + self.FromArray[x - 1])
            self.LstCount += 1

    def addemails(self, event):
        self.Widget = event.widget
        self.selection = self.Widget.curselection()
        self.value = self.Widget.get(self.selection[0])
        self.new = ''
        for Char in str(self.selection):
            if Char not in ['(',')',',', "'"]:
                self.new += Char
            else:
                pass
        self.New = int(self.new) 
        
        self.New_3 = len(self.FullFromArray) - self.New
        self.New_2 = 'Email' + str(self.New_3)
        self.TextEnt.delete(1.0, END) 
        for Line in self.BodyDictionary[self.New_2]:
           self.TextEnt.insert(END, Line)
        self.FromEnt.delete(0, END)
        self.SubEnt.delete(0, END)
        self.FromEnt.insert(0, self.FullFromArray[int(self.New_3) - 1])
        self.SubEnt.insert(0, self.SubjectArray[int(self.New_3) - 1])

    def GetTxtEnt(self):
        global Email, Password
        self.TO = self.ToEnt1.get()
        self.SUBJECT = self.SubEnt1.get()
        self.EmailTxt = self.TextEnt1.get(1.0, END)
        self.server = smtplib.SMTP('smtp.gmail.com', 587)
        self.server.ehlo()
        self.server.starttls()
        self.server.ehlo()
        self.server.login(Email, Password)
        if self.AttatchedFlag == False:
            
            self.BODY = '\r\n'.join(['To: %s' % self.TO,
                                'From: %s' % Email,
                                'Subject: %s' % self.SUBJECT,
                                '', self.EmailTxt])
            print(self.BODY)
            try:
                self.server.sendmail(Email, [self.TO], self.BODY)
                print ('Email sent')
                self.parent2.destroy()
            except:
                print ('Error sending mail')

            self.server.quit()

        else:
            self.AttatchedFlag = True
            Count = 0
            msg = MIMEMultipart()
            msg['From'] = Email
            msg['To'] = self.TO
            msg['Subject'] = self.SUBJECT
            self.text = self.EmailTxt
            msg.attach(MIMEText(self.text))
            for Char in self.FileArray:
                attach = self.FileArray[Count]
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(open(attach, 'rb').read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition',
                       'attachment; filename="%s"' % os.path.basename(attach))
                msg.attach(part)
                Count += 1
            try:
                self.server.sendmail(Email, self.TO, msg.as_string())
                self.server.close()
                self.parent2.destroy()
            except:
                print('error')
            
            

    def ClearTo(self,event):
        self.ToEnt1.delete(0, END)

    def ClearSub(self,event):
        self.SubEnt1.delete(0, END)

    def SendMail(self):
        self.parent2 = Tk()
        self.parent2.resizable(0,0)
        self.parent2.geometry("860x725")
        self.parent2.title('Send Mail')
        self.parent2.configure(background="steel blue")
        self.parent2.grid()
        self.topfrm = Frame(self.parent2, width="4000", height="4000", bg="steel blue")
        self.topfrm.grid(row=0, column = 0)
        self.bottomfrm = Frame(self.parent2, width="4000", height="4000", bg="steel blue")
        self.bottomfrm.grid(row=2, column = 0)
        self.SendMailFrm = Frame(self.parent2, width="4000", height="4000", bg="steel blue")
        self.SendMailFrm.grid(row=1, column=0)
        self.ToEnt1 = Entry(self.topfrm, bd =4, width=98, font=('Calibri', 10))
        self.SubEnt1 = Entry(self.topfrm, bd =4, width=98, font=('Calibri', 10))
        self.labelto = Label(self.topfrm, text="To: ", bg="steel blue")
        self.labelsub = Label(self.topfrm, text="Subject: ", bg="steel blue")
        self.TextEnt1 = Text(self.SendMailFrm,width="95", height="30", bg='white', wrap=NONE, bd=4)

        self.SendBtn = Button(self.bottomfrm, text='Send',height=2,width=10,fg='black', activebackground="yellow",bg='light blue', command=self.GetTxtEnt, bd=4)
        self.AttachBtn = Button(self.bottomfrm, text = "Attach Files",height=2,width=10,fg='black', activebackground="yellow",bg='light blue', bd=4, command = self.Attatcher)
        self.lstscrollbar2 = Scrollbar(self.SendMailFrm)
        self.lstscrollbar2.grid(row=2, column=1, rowspan=3,pady=20, sticky='ns')
        self.TextEnt1.config(yscrollcommand=self.lstscrollbar2.set)
        self.lstscrollbar2.config(command=self.TextEnt.yview)
        self.labelto.grid(row=0, column=0)
        self.labelsub.grid(row=1, column=0)
        self.ToEnt1.grid(row=0, column=1, padx=20, pady=20)
        self.SubEnt1.grid(row=1, column=1, padx=20, pady=20)
        self.TextEnt1.grid(row=2, column=0, padx=20, pady=20, rowspan=3, sticky=(N))
        self.SendBtn.grid(row=0, column=0, padx=20, pady=20)
        self.AttachBtn.grid(row = 0, column = 1, padx = 20, pady = 20)
        self.parent2.mainloop()

    def Attatcher(self):
        self.AttatchedFlag = True
        self.root6 = Tk()
        self.root6.filename =  filedialog.askopenfilename(initialdir = "E:/Images",title = "choose your file",filetypes = (("jpeg files","*.jpg"),("all files","*.*")))
        print(self.root6.filename)
        self.root6.destroy()
        self.FileArray.append(self.root6.filename)
        
firstwindow = FirstWindow(root)
root.mainloop() 
root1 = Tk()
secondwindow = SecondWindow(root1)
root1.mainloop()




