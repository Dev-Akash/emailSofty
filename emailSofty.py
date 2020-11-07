import smtplib
import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import *
import re

sel = 0
chooseMa = 2
data = pd.DataFrame()

root = Tk()
root.title("EmailSofty")
root.resizable(False, False)

win1 = PanedWindow(root)
win1.pack(side = "left",fill="both",expand=1)

labelframe = LabelFrame(win1, text="Sender Settings")
labelframe.pack(fill="x", expand="yes")

left = Label(labelframe, text="Email:", width=50,anchor='w')
left.pack(fill="x", expand=1)

senderEmail = Entry(labelframe)
senderEmail.pack(fill="x", expand=1)

LPa = Label(labelframe, text="Password:", width=50,anchor='w')
LPa.pack(fill="x", expand=1)

senderPass = Entry(labelframe, show="*")
senderPass.pack(fill="x", expand=1)

Ema = LabelFrame(win1, text="Email Settings")
Ema.pack(fill="x", expand=1)

Lrec = Label(Ema, text="Recipients", width=50, anchor='w')
Lrec.pack(fill="x", expand=1)

RecipientsList = Entry(Ema)
RecipientsList.pack(fill="x", expand=1)

labSubject = Label(Ema, text= "Subject", anchor='w')
labSubject.pack(fill='x', expand=1)

subject = Entry(Ema)
subject.pack(fill="x", expand=1)

EType = Label(Ema, text="Email Type", width=50, anchor='w')
EType.pack(fill="x", expand=1)

def choose():
	global sel 
	sel = var.get()
	print(sel)

var = IntVar()

R1 = Radiobutton(Ema, text = "Text", variable=var, value=1, command=choose)
R1.pack(anchor=W)

R2 = Radiobutton(Ema, text = "HTML", variable=var, value=2, command=choose)
R2.pack(anchor=W)

mLabe = Label(Ema, text="Message", anchor='w')
mLabe.pack(fill="x", expand=1)

Emailmessage = Text(Ema, width=50, height=5)
Emailmessage.pack(fill="both")

def sendEmails():
	print(chooseMa)
	if(chooseMa == 2):
		reciver = RecipientsList.get()
		sub = subject.get()
		message=Emailmessage.get("1.0", END)
		#print(unam,upass, reciver, sub, message)
		finalmessage=" "
		#print(sel)
		if (sel == 1):
			finalmessage = """From: {}\nTo: {}\nSubject: {}\n{}""".format(unam,reciver,sub,message)
			#print(finalmessage)
		elif (sel == 2):
			print("woking in sec 2")
			finalmessage = """From: {}\nTo: {}\nMIME-Version: 1.0\nContent-type: text/html\nSubject: {}\n{}""".format(unam,reciver,sub,message)
		print(finalmessage)
		try:
			mail = smtplib.SMTP("smtp.gmail.com", 587)
			mail.starttls()
			mail.login(unam,upass)
			mail.sendmail(unam, reciver, finalmessage)
			mail.close()
			Pbar['value'] = 100
			print("Successfully sent Email")
			statusLabe.config(text="Status of Email: Successfully sent")
			root.update_idletasks()
		except Exception as e:
			print("Error: Unable to Send Email")
			statusLabe.config(text="Status of Email: Unable to Send")
	elif(chooseMa==1):
		r = RecipientsList.get()
		emails = data.loc[:,r.strip('{ }')]
		emails = np.asarray(emails)
		subarr = []
		sub = subject.get()
		val = re.findall(r'\{(.*?)\}', sub)
		print(val[0])
		s = data.loc[:,val[0]]
		s = np.asarray(s)
		for i in s:
			temp = subject.get()
			a = "{"+val[0]+"}"
			#print(a, i)
			#print(temp)
			subarr.append(temp.replace(a, i))
		print(subarr)



sendButton = Button(win1, text = "Send", command=sendEmails)
sendButton.pack(fill='x', expand=1)

win2= PanedWindow(root, width=250)
win2.pack(side="right", expand=1)

mmui = LabelFrame(win2, text="Mail Merge Settings")
mmui.pack(fill="x", expand=1)

def chooseMail():
	global chooseMa
	chooseMa = var2.get()

var2 = IntVar()
var2.set(2)

r3 = Radiobutton(mmui, text = "Enables Mail-Merge Option", variable=var2, value=1, command=chooseMail)
r3.pack(anchor="w")

r4 = Radiobutton(mmui, text = "Disable Mail-Merge Option", variable=var2, value=2, command=chooseMail)
r4.pack(anchor="w")

fileLabel = Label(mmui, text="Choose File (*.csv)", anchor="w",width=50)
fileLabel.pack(fill="x", expand=1)

fileEntry = Entry(mmui)
fileEntry.pack(fill="x",expand=1)

def pickFile():
	filename= filedialog.askopenfilenames(filetypes=[("Excel Files","*.csv")])
	fileEntry.delete(0, END)
	fileEntry.insert(0,filename)
	global data
	data = pd.read_csv(filename[0])
	fields = data.columns
	for i in fields:
		showFields.insert(END, i+"\n")

browseButton = Button(mmui, text="Browse", width=10, command=pickFile)
browseButton.pack(anchor="w")

fieldslable = Label(mmui, text = "Fields in Data:")
fieldslable.pack(anchor="w")

showFields = Text(mmui, width=50, height=10)
showFields.pack(fill="both", expand=1)

statusLableFrame = LabelFrame(win2, text="Status", height=110)
statusLableFrame.pack(fill="both", expand=1)

statusLabe = Label(statusLableFrame, text = "Status of Email: ")
statusLabe.pack(anchor="w")

progLabel = Label(statusLableFrame, text = "Progress")
progLabel.pack(anchor="w")

Pbar = Progressbar(statusLableFrame, orient= HORIZONTAL, length=100, mode="determinate")
Pbar.pack(fill="x", expand=1)

root.mainloop()