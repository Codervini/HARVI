import csv
from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as mbx
import os
import sys
import subprocess
import pickle
import getpass
import win32com.client 

APP_LINK_FILE = "Taskdata/applinktaskdata.csv"
WEB_LINK_FILE = "Taskdata/weblinktaskdata.csv"


try:
    with open(APP_LINK_FILE, "r") as taskfile:
        pass
    with open(WEB_LINK_FILE, "r") as taskfile:
        pass
except FileNotFoundError:
  os.makedirs("Taskdata")
  with open(APP_LINK_FILE, "w", newline="") as taskfile:
        writer = csv.writer(taskfile)
        writer.writerow(["Name of task", "Task Objective", "Task Description", "Mode of Adding"])
  with open(WEB_LINK_FILE, "w", newline="") as taskfile:
      writer = csv.writer(taskfile)
      writer.writerow(["Name of task", "Task Objective", "Task Description", "Mode of Adding"])

root = Tk()
root.title("Task Adder")
root.resizable(False, False)

TStaskmode = StringVar(root)
TStaskname = StringVar(root)
TStaskloc = StringVar(root)
TStaskdesc = StringVar(root)



def openfileloc():
  filetypes = (
        ('Executable files', '*.exe'),
        ('All files', '*.*')      
      )
  fileloc= fd.askopenfilename(title="Open file", initialdir='C:/', filetypes=filetypes)
  entobj.delete(0, "end")
  TStaskloc.set(fileloc)
  if entobj['state'] == "readonly":
    entobj.config(state="normal")
    entobj.insert(0, TStaskloc.get())
    entobj.config(state="readonly")
  else:
      entobj.insert(0, TStaskloc.get())
 
def entrydisabler():
  if v.get() == "applink":
    entobj.delete(0, "end")
    entobj.config(state="readonly")
    entobjbtn.config(state="normal")
  else:
    entobjbtn.config(state="disabled")
    entobj.config(state="normal")
    entobj.delete(0, "end")

    
def restart_program():
    """Restarts the current program.
    Note: this function does not return. Any cleanup action (like
    saving data) must be done before calling this function."""
    python = sys.executable
    os.execl(python, python, * sys.argv)


def savetaskdata(type, file):
  """Writes the the recieved record to the appropraite files"""
  if "" not in [v.get(), entry.get(), entobj.get()]:
    taskfile = open(file, "a", newline="")
    twriter = csv.writer(taskfile)
    if len(txtbox.get("1.0","end")) <=1:
      txtbox.delete("1.0","end")
      txtbox.insert("1.0", "No description")   
    if type == "SE":
      TStaskname.set(entry.get())
      TStaskloc.set(entobj.get())
      TStaskdesc.set(txtbox.get('1.0', 'end-1c'))
      twriter.writerow([TStaskname.get().lower(), TStaskloc.get(), TStaskdesc.get(), "User Added"])
      taskfile.close()
      root.destroy()
    elif type == "SA":
      TStaskname.set(entry.get())
      TStaskloc.set(entobj.get())
      TStaskdesc.set(txtbox.get('1.0', 'end-1c'))
      twriter.writerow([TStaskname.get().lower(), TStaskloc.get(), TStaskdesc.get(), "User Added"])
      taskfile.flush()
      restart_program()
  else:
    mbx.showerror(title="Error", message="One or more fields not filed!")

def file_decider(mode, operationtype):
  if mode == "applink":
    savetaskdata(operationtype, APP_LINK_FILE)
  elif mode == "weblink":
    savetaskdata(operationtype, WEB_LINK_FILE)

def dir_cmd_getter(keyword : str):

  command = f'dir \"{keyword}*.exe"/s'
  results = subprocess.Popen(command, shell=True, cwd='C:\\', stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
  stdout, stderr = results.communicate()
  print(stdout, stderr)

def lnk_location_getter():
  try:
      os.mkdir(".appdata")
  except FileExistsError:
      pass

  username = getpass.getuser()
  FILE = ".appdata/cmdlnkdata.dat"
  DIRLIST = ['C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs',
          f'C:\\Users\\{username}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs'
          ]
  loclist = []

  with open(APP_LINK_FILE, "r", newline="") as file:
    reader = csv.reader(file)
    for i in reader:
      if i[3] == "User Added":
        with open(".appdata/tempcsv.csv", "a", newline="") as f:
          writer = csv.writer(f)
          writer.writerow(i)

  for i in DIRLIST:
      with open(FILE,  "wb") as f:
          c = 'dir \"*.lnk"/s'
          results = subprocess.Popen(c, shell=True, cwd=i, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
          stdout, stderr = results.communicate()
          pickle.dump(stdout,f)
      with open(FILE,  "rb") as f:
          bin = pickle.load(f)
          bin_list =  bin.split("\n")
          for i in bin_list:
              if i != " ":
                  # print(i)
                  if "Directory of" in i:
                      i.lstrip()
                      abs_path = i [14:]
                  if '.lnk' in i:
                      str_list = i.split("         ")
                      whitespace_index = str_list[1].lstrip().index(" ")
                      lnk_file = str_list[1].lstrip() [whitespace_index+1:]
                      shell = win32com.client.Dispatch("WScript.Shell")
                      shortcut = shell.CreateShortCut(f'{abs_path}\{lnk_file}')
                      if len(shortcut.Targetpath) != 0:
                        loclist.append([lnk_file[:len(lnk_file)-4].lower(), shortcut.Targetpath, "No description", "Program Added"])

      with open(APP_LINK_FILE, "w", newline="") as taskfile:
        twriter = csv.writer(taskfile)
        twriter.writerows(loclist)
      templist = []
      with open(".appdata/tempcsv.csv", "r", newline="") as f:
        read = csv.reader(f)
        for i in read:
          templist.append(i)
      with open(APP_LINK_FILE, "a", newline="") as taskfile:
        twriter = csv.writer(taskfile)
        twriter.writerows(templist)


        


# Mode Choices
radframe = LabelFrame(root, text= "Mode: ", height=100, width=100)
radframe.grid(padx=10, pady=10)
v = StringVar(radframe, "weblink")
values = {"Applink" : "applink",
          "Weblink" : "weblink",
        }
for text, value in values.items():
    rad = Radiobutton(radframe, text=text, variable=v, value=value, command=entrydisabler)
    rad.grid(padx= 70)



# Name of the Task
entframe = LabelFrame(root, text="Name of Task: ", padx=10, pady=10)
entframe.grid(padx=10, pady=10)
entry = Entry(entframe, width=30)
entry.grid(padx=5)

# Task Objective
entobjframe = LabelFrame(root, text="Tasks Objective: ", padx=10, pady=10)
entobjframe.grid(padx=10, pady=10)
entobj = Entry(entobjframe, width=32)
entobj.grid()
entobjbtn = Button(entobjframe, text="Open file location", command=openfileloc, state="disabled")
entobjbtn.grid(row=1, pady=(10,0))



# Describe The Task
txtboxframe = LabelFrame(root, text="Describe the task (Optional): ")
txtboxframe.grid(padx=10, pady=10)
txtbox = Text(txtboxframe, width=25, height=5)
txtbox.grid(padx=8, pady=10)


saveandexitbtn = Button(root, text="Save and Exit", width=10, command=lambda: file_decider(v.get(),"SE")).grid(row= 0 , column=1, padx=5)
addanotherbtn = Button(root, text="Add Another", width=10, command=lambda: file_decider(v.get(), "SA")).grid(row=0, column=1, pady=(70,0))
exitbtn = Button(root, text="Exit", width=10, command=root.destroy).grid(row=3, column=1)







def taskwriter():
    root.mainloop()

def taskreader(taskmode):
  if taskmode == "applink":
    file  = APP_LINK_FILE
  elif taskmode == "weblink":
    file = WEB_LINK_FILE
  with open(file, "r", newline="") as taskfile:
    treader = csv.reader(taskfile) 
    next(treader)
    taskdata=[]
    for name, objective, description in treader:
      d={name:[ objective, description]}
      taskdata.append(d)
    return taskdata
                


lnk_location_getter()
# dir_cmd_getter("zoom")

#cmd_dir_processor()

#taskwriter()
# for i in taskreader("weblink"):
#   print(i)







