import tkinter as tk
from datetime import datetime
import threading
from typing import DefaultDict
import gspread 
import datetime
from oauth2client.service_account import ServiceAccountCredentials 
from time import sleep
from tkinter import *
from tkinter import ttk
import yfinance as yf
from tkinter import messagebox
import time
x = "กำลังทำงาน"
p = "กำลังบันทึก"
y = "หยุดทำงาน"
s = "บันทึกสำเร็จ"
n="ยังไม่เปิดใช้งาน"
wait_time = 1
filename = 'ONE-UGG-RA'
#เชื่อมต่อกับชีท
headers = {
    'x-ibm-client-id': "90345183-71fc-4150-a771-bb84cef18d56",
    'accept': "application/json"
    }
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r'creds.json', scope)
client = gspread.authorize(creds)

worksheet1 = client.open(filename).sheet1

worksheet = client.open(filename).get_worksheet(3)
worksheet2 = client.open('ONE-UGG-RA').worksheet('stock name')

def addtitle() :
  worksheet_Test = client.open('ONE-UGG-RA').worksheet('DATA')
  worksheet_Name = client.open('ONE-UGG-RA').worksheet('stock name')
  havecol = len(worksheet_Test.row_values(1))
  namekey = 'Name', 'Open',	'High',	'Low',	'Close',	'Adj Close',	'Volume'
  if havecol == 0 or worksheet_Test.cell(1,1).value == '' or worksheet_Test.cell(1,1).value.lower() != 'date' :
    worksheet_Test.update_cell(1,1,'Date')
    havecol = len(worksheet_Test.row_values(1))
    

  name = worksheet_Name.row_values(1)
  needcol = (len(name)*len(namekey))+1
  if havecol == needcol :
    print('มี Title อยู่แล้ว')

  else :
    if worksheet_Test.col_count <= needcol :
      worksheet_Test.add_cols(needcol-worksheet_Test.col_count)
      worksheet_Test = client.open('ONE-UGG-RA').worksheet('DATA')

    nameintitle = worksheet_Test.row_values(1)
    for i in range (len(name)) :
      for j in range (len(namekey)) :
        nameintitle.append(namekey[j])

    title_list = worksheet_Test.range(1,1,1,needcol)
    for i in range (len(title_list)) :
      title_list[i].value = nameintitle[i]

    worksheet_Test.update_cells(title_list)
    print("Add Title Succeed")
    charktitle()

def adddatas() :
  worksheet_Test = client.open('ONE-UGG-RA').worksheet('DATA')
  worksheet_Name = client.open('ONE-UGG-RA').worksheet('stock name')
  if worksheet_Test.cell(2,1).value != str(datetime.datetime.now())[0:10] :
    name = worksheet_Name.row_values(1)
    needcol = (len(name)*7)+1
    adddata = [str(datetime.datetime.now())[0:10]]
    if worksheet_Test.col_count <= needcol :
      worksheet_Test.add_cols(needcol-worksheet_Test.col_count)
      worksheet_Test = client.open('ONE-UGG-RA').worksheet('DATA')

    for i in range (len(name)) :
      comm = yf.download(tickers = name[i],period='1d')
      z = 2
      while comm.empty :
        comm = yf.download(tickers = name[i],period=str(str(z)+'d'))
        z = z+1
      namekey = list(comm.keys())
      adddata.append(str(name[i]))
      for i in range (len(namekey)) : 
        adddata.append(float(comm.get(namekey[i]).values[-1]))

    worksheet_Test.insert_row(adddata,2)
    print("Add Data Succeed")
  else :
    print("ข้อมูลของวันนี้บันทึกไปแล้วนะ")

def charktitle() :
  worksheet_Test = client.open('ONE-UGG-RA').worksheet('DATA')
  title = worksheet_Test.row_values(1)
  newdata = worksheet_Test.row_values(2)
  namekey = 'Name', 'Open',	'High',	'Low',	'Close',	'Adj Close',	'Volume'
  if len(title) < 1 :
    while (len(title)) != (len(newdata)):
      if (len(title)) < (len(newdata)) :
        if title[-1].lower() == namekey[0].lower() :
          for i in range (1,len(namekey)) :
            title.append(namekey[i])
        elif title[-1].lower() == namekey[1].lower() :
          for i in range (2,len(namekey)) :
            title.append(namekey[i])
        elif title[-1].lower() == namekey[2].lower() :
          for i in range (3,len(namekey)) :
              title.append(namekey[i])
        elif title[-1].lower() == namekey[3].lower() :
          for i in range (4,len(namekey)) :
            title.append(namekey[i])
        elif title[-1].lower() == namekey[4].lower() :
          for i in range (5,len(namekey)) :
            title.append(namekey[i])
        elif title[-1].lower() == namekey[5].lower() :
          for i in range (6,len(namekey)) :
            title.append(namekey[i])
        elif title[-1].lower() == namekey[6].lower() :
          for i in range (len(namekey)) :
            title.append(namekey[i])
        else :
          title.pop(-1)
          title.append('')
      if (len(newdata)) < (len(title)) :
        title.pop(-1)
        title.append('')
      title_list = worksheet_Test.range(1,1,1,(len(title)))
      for i in range (len(title_list)) :
        title_list[i].value = title[i]
      worksheet_Test.update_cells(title_list)
      title = worksheet_Test.row_values(1)
  print("Chark Title Succeed")

def startProgram() :
    doit = True
    btRun["state"] = "disabled"
    chacktime = _select_date_time()  
    while btRun["state"] == "disabled":
        lb2["text"] = x
        lbsh["text"] = chacktime
        worksheet_Test = client.open('ONE-UGG-RA').worksheet('DATA')
        timenow = (str(datetime.datetime.now().strftime("%X")))
        if timenow == chacktime and doit == True :
            addtitle()
            adddatas()
            charktitle()
            if worksheet_Test.cell(2,1).value != str(datetime.datetime.now())[0:10] :
                addtitle()
                adddatas()
                charktitle()
            else:
                print("บันทึกแล้ว")
                doit = False
                lb2["text"] = s
                sleep(10)

        elif timenow == chacktime and doit == False :
            sleep(wait_time)

        elif timenow != chacktime and doit == False :
            doit = True
            sleep(wait_time)

        elif timenow != chacktime and doit == True :
            print('Timenow :' + timenow)
            print('Chacktime :' + chacktime)
            sleep(wait_time)

        else :
            print('Timenow :' + timenow)
            print('Chacktime :' + chacktime)
            sleep(wait_time)
def resave():
        lb2["text"] = p
        addtitle()
        adddatas()
        charktitle()
        lb2["text"] = n
        msg="บันทึกเรียบร้อย"
        messagebox.showinfo("บันทึกเรียบร้อย",msg)

def stopProgram():
    lb2["text"] = y
    btRun["state"] = "normal"

def show():
    if hour_combobox.get() == '' or hour_combobox.get().replace(' ','') == '' and minutes_combobox.get() == '' or minutes_combobox.get().replace(' ','') == ''and seconds_combobox.get() == '' or seconds_combobox.get().replace(' ','') == ''  :
        msg="กรุณาตั้งเวลา"
        messagebox.showinfo("กรุณาตั้งเวลา",msg)
    else:
        _select_date_time(),runThread()
 
def openNewWindow():
    newWindow = Toplevel(master=window)
    newWindow.title("เพิ่มจำนวชื่อหุ้น")
    newWindow.resizable(0,0) 
    def insert():

        if entry.get() == '' or entry.get().replace(' ','') == '' :
            msg="กรุณากรอกใหม่"
            messagebox.showinfo("กรุณากรอกใหม่",msg)
        else:
            
            comm = yf.download(tickers = entry.get(),period='10d')
            if comm.empty :
                msg="บันทึก"+entry.get()+"ไม่สำเร็จกรุณากรอกใหม่"
                messagebox.showinfo("กรุณากรอกใหม่",msg)
            else :
                name = worksheet2.row_values(1)
                add = entry.get()
                if worksheet2.col_count <= (len(name)+1) :
                    worksheet2.add_cols((len(name)+1)-worksheet2.col_count)
                worksheet2.update_cell(1,(len(name)+1),add)
                msg="บันทึก"+entry.get()+"สำเร็จ"
                messagebox.showinfo("บันทึกสำเร็จ",msg,)
                newWindow.destroy()
                
             
    

    LS=Label(newWindow,text ="เพิ่มจำนวนชื่อหุ้น")
    LS.grid(row=0, column=0, padx=(20,5), pady=(20,5))
    LS.config(font=("Tahoma", 28))
    ES=Entry(newWindow,textvariable=entry)
    ES.grid(row=1, column=0, padx=(5,5), pady=(20,5))
    ES.config(font=("Tahoma", 25))
    BS=Button(newWindow,text="บันทึก", command = insert)
    BS.grid(row=1, column=1, padx=(20,5), pady=(20,5))
    BS.config(font=("Tahoma", 25))

def runThread():
    
    x = threading.Thread(target=startProgram)
    x.start()

def rerunThread():
    
    x = threading.Thread(target=resave)
    x.start()

def _select_date_time():
    hour = hour_combobox.get()
    minutes = minutes_combobox.get()
    seconds = seconds_combobox.get()
    time25 = ""
    date_time = {}

    if int(hour) < 10:
        hour = '0' + hour
    
    if int(minutes) < 10:
        minutes = '0' + minutes
    
    if int(seconds) < 10:
        seconds = '0' + seconds

    date_time['hour'] = hour
    date_time['minutes'] = minutes
    date_time['seconds'] = seconds

    date = date_time
    if date:
        hour = date['hour']
        minutes = date['minutes']
        seconds = date['seconds']

        time2 = hour + ':' + minutes + ':' + seconds
        time25 = time2
    return(time25)

def digital_clock(): 
   time_live = time.strftime("%H:%M:%S")
   lbtnow.config(text=time_live) 
   lbtnow.after(200, digital_clock)
	
#GUI โปรแกรม
window = tk.Tk()
entry = StringVar()  
window.title("โปรแกรมบันทึกข้อมูลหุ้นอัตโนมัติ")
window.resizable(0,0)  #don't allow resize in x or y direction
window.iconbitmap(default='money.ico')
lb1 = tk.Label(master=window, text="โปรแกรมบันทึกข้อมูลหุ้นอัตโนมัติ")
lb1.grid(row=0, padx=(15,15), pady=(5,5))  #padx=(left,right), pady=(top,button)
lb1.config(font=("Tahoma", 28))

lb1 = tk.Label(master=window, text="สถานะระบบ: ")
lb1.grid(row=1, column=0, padx=(20,250), pady=(5,20))  #padx=(left,right), pady=(top,button)
lb1.config(font=("Tahoma", 28))

lb2 = tk.Label(master=window, text="ยังไม่เปิดใช้งาน")
lb2.grid(row=1, column=0, padx=(250,20), pady=(5,20))
lb2.config(font=("Tahoma", 28))


lbtn = tk.Label(master=window, text="กรุณาตั้งเวลาบันทึก")
lbtn.grid(row=2, column=0, padx=(20,275), pady=(0,0)) 
lbtn.config(font=("Tahoma",28))

lbh=Label(master=window, text='Hours')
lbh.grid(row=2, column=0, padx=(150,20), pady=(40,5))
hour_combobox = ttk.Combobox(window.master, width=3, values=[*range(0,24)])
hour_combobox.grid(row=2, column=0, padx=(150,20), pady=(0,0))


lbchm=Label(master=window, text=':')
lbchm.grid(row=2, column=0, padx=(337,20), pady=(0,0))
lbchm.config(font=("Tahoma",10))

lbm=Label(master=window, text='Minutes')
lbm.grid(row=2, column=0, padx=(275,20), pady=(40,5))
minutes_combobox = ttk.Combobox(window.master, width=3, values=[*range(0,60)])
minutes_combobox.grid(row=2, column=0, padx=(275,20), pady=(0,0))


lbcms=Label(master=window, text=':')
lbcms.grid(row=2, column=0, padx=(212,20), pady=(0,0))
lbcms.config(font=("Tahoma",10))

lbs=Label(window.master, text='Seconds')
lbs.grid(row=2, column=0, padx=(400,20), pady=(40,5))
seconds_combobox = ttk.Combobox(window.master, width=3,values=[*range(0,60)])
seconds_combobox.grid(row=2, column=0, padx=(400,20), pady=(0,0))

lbsh = tk.Label(master=window,text="ขณะนี้เวลา")
lbsh.grid(row=3, column=0, padx=(20,275), pady=(0,0)) 
lbsh.config(font=("Tahoma",28))

lbtnow = tk.Label(master=window)
lbtnow.grid(row=3, column=0, padx=(275,20), pady=(0,0)) 
lbtnow.config(font=("Tahoma",28))

lbsh = tk.Label(master=window,text="เวลาบันทึกข้อมูล")
lbsh.grid(row=4, column=0, padx=(20,275), pady=(0,0)) 
lbsh.config(font=("Tahoma",28))

lbsh = tk.Label(master=window,text="")
lbsh.grid(row=4, column=0, padx=(275,20), pady=(0,0)) 
lbsh.config(font=("Tahoma",28))



btRun = tk.Button(master=window, text="start", command=show)
btRun.grid(row=5, column=0, padx=(20,250), pady=(5,20))
btRun.config(font=("Tahoma", 25), width=10)

btStop = tk.Button(master=window, text="stop", command=stopProgram)
btStop.grid(row=5, column=0, padx=(250,20), pady=(5,20))
btStop.config(font=("Tahoma", 25), width=10)

btn = Button(master=window,text ="กรณีข้อมูลไม่เข้า",command = rerunThread)
btn.grid(row=6, column=0, padx=(20,5), pady=(5,20))
btn.config(font=("Tahoma",15))

btn = Button(master=window,text ="เพิ่มชื่อหุ้นที่ต้องการบันทึก",command = openNewWindow)
btn.grid(row=7, column=0, padx=(20,5), pady=(5,20))
btn.config(font=("Tahoma",15))

digital_clock()
window.mainloop()





