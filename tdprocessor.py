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


APP_LINK_FILE = ".harvidata/applinktaskdata.csv"
WEB_LINK_FILE = ".harvidata/weblinktaskdata.csv"


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


def get_app_locations():
  '''
  returns the list of all the installed programs added in the start menu
  '''
  link_list = []  # list of all the exe file locations
  username = getpass.getuser()
  shell = win32com.client.Dispatch("WScript.Shell")

  DIRLIST = ['C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs',
             f'C:\\Users\\{username}\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs']

  for dirs in DIRLIST:
      for (dirpath, dirnames, filenames) in os.walk(dirs):
          for name in filenames:
              if name.endswith(".lnk"):
                  path = os.path.join(dirpath, name)
                  shortcut = shell.CreateShortCut(path)
                  link_list.append(shortcut.Targetpath)
  return link_list


def write_app_locations():
  '''
  void function writes all the app locations to a csv
  '''
  link_list = get_app_locations()

  # create .harvidata directory
  try:
      # os.chdir(f"c:\\users\\{username}")
      os.mkdir(".harvidata")
      os.system('attrib +h ".harvidata"')
  except FileExistsError:
      pass
  

  with open(APP_LINK_FILE, "w", newline="") as taskfile:
    twriter = csv.writer(taskfile)
    twriter.writerow(["Name of task", "Task link", "Mode of Adding"])
    # twriter.writerows(loclist)
    templist = []
  



def get_app_locations_depreciated():
  username = getpass.getuser()
  try:
      # os.chdir(f"c:\\users\\{username}")
      os.mkdir(".appdata")
      os.system('attrib +h ".appdata"')
  except FileExistsError:
      pass

  
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
      with open(".appdata/cmdlnkdata.dat",  "wb") as f:
          c = 'dir \"*.lnk"/s'
          results = subprocess.Popen(c, shell=True, cwd=i, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
          stdout, stderr = results.communicate()
          pickle.dump(stdout,f)
      with open(".appdata/cmdlnkdata.dat",  "rb") as f:
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
        twriter.writerow(["Name of task", "Task Objective", "Task Description", "Mode of Adding"])
        twriter.writerows(loclist)
      templist = []
      try: 
        with open(".appdata/tempcsv.csv", "r", newline="") as f:
          read = csv.reader(f)
          for i in read:
            templist.append(i)
      except FileNotFoundError:
        pass
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
                


get_app_locations()
# dir_cmd_getter("zoom")

#cmd_dir_processor()

#taskwriter()
# for i in taskreader("weblink"):
#   print(i)







