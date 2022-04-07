# Tkinter UI
from tkinter import *
from PIL import ImageTk
import os
from tkinter import filedialog,messagebox
from damm_you_door_v1 import excel_door_open

root = Tk()
root.geometry("240x240")
root.title('FuddyDuddy-Excel')
root.resizable(0, 0)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

path = resource_path("XLS.png")
img = PhotoImage(file = path, master= root)

# img = PhotoImage(file='XLS.png', master= root)
img_label = Label(root, image = img)
img_label.place(x = 0, y = 0)

s_filepath = ""
d_filepath = ""

# Select Source Path
def s_open_file():
    global s_filepath

    file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.xlsx')])
    if file:
        s_filepath = os.path.abspath(file.name)
        # messagebox.showinfo('Information','File Selected')
        source_button.configure(bg = 'light green')
    else:
        messagebox.showwarning('Warning','File Not Selected')
        source_button.configure(bg = 'red')

# Select Destination Path
def d_open_file():
    global d_filepath

    d_filepath = filedialog.askdirectory()
    if(d_filepath):
        # messagebox.showinfo('Information','Destination Selected')
        destination_button.configure(bg='light green')
    else:
        messagebox.showwarning('Warning','Destination Not Selected')
        destination_button.configure(bg='red')

# 
def excel_automation():
    if(s_filepath and d_filepath):
        # messagebox.showinfo('Information',s_filepath)
        # messagebox.showinfo('Information',d_filepath)
        excel_door_open(s_filepath, d_filepath)
    else:
        messagebox.showwarning('Warning','Please Select Source and Destination.')

    root.destroy()


# Create a Button - Take Source File
source_button = Button(root, text="Source", command = s_open_file)
source_button.place(x = 94, y = 50)

# Create a Button - Select Destination Path
destination_button = Button(root, text="Destination", command = d_open_file)
destination_button.place(x = 83, y = 80)

# print(s_filepath)
# print(d_filepath)

action_button = Button(root, text = "Action", command = excel_automation)
action_button.place(x = 96, y = 120)


root.mainloop()