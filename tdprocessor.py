import csv
from tkinter import *
from tkinter import filedialog as fd
from tkinter import messagebox as mbx
import os
import sys
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
  os.makedirs(".harvidata")
  os.system('attrib +h ".harvidata"')
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
  path_list = []  # list of all the exe file locations
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
                  target_path = shortcut.Targetpath
                  if target_path.endswith(".exe") or target_path.endswith(".EXE") or target_path.endswith(".msc"):
                    filename = (name.split(".")[0]).lower()
                    path_list.append((filename, shortcut.Targetpath))
  return path_list


def write_app_locations():
  '''
  void function writes all the app locations to a csv
  '''
  path_list = get_app_locations()
  write_list = []

  # create .harvidata directory
  try:
      # os.chdir(f"c:\\users\\{username}")
      os.mkdir(".harvidata")
      os.system('attrib +h ".harvidata"')
  except FileExistsError:
      pass

  # creating a list of rows to write to the csv
  for name, path in path_list:
      head_tail = os.path.split(path)
      # task_name = head_tail[1].split(".")[0].lower()
      write_list.append([name, path, "No description", "Program Added"])
  
  # reading the csv files for user added tasks and appending them to a list
  userprogam_list = []

  with open(APP_LINK_FILE, "r", newline="") as taskfile:
    reader = csv.reader(taskfile)
    for i in reader:
      if i[3] == "User Added":
        userprogam_list.append(i)
  
  # Creating app_link csv
  with open(APP_LINK_FILE, "w", newline="") as taskfile:
    twriter = csv.writer(taskfile)
    twriter.writerow(["Name of task", "Task path", "Task Description", "Mode of Adding"])
    write_list.extend(userprogam_list)
    twriter.writerows(write_list)

  
  
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
                


write_app_locations()


#taskwriter()
# for i in taskreader("weblink"):
#   print(i)







