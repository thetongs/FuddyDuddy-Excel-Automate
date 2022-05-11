# Tkinter UI
from tkinter import *
from PIL import ImageTk
import os
from tkinter import filedialog,messagebox
from damm_you_door_v3 import excel_door_open
from excel_to_pdf_v3 import excel_conversion
from pathlib import Path

# Tkinter GUI
root = Tk()
root.geometry("240x265")
root.title('FuddyDuddy-Excel')
root.resizable(0, 0)
root.wm_attributes('-transparentcolor', '#ab23ff')

# Pyinstaller Path Requirements
def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Background Image
path = resource_path("XLS.png")
img = PhotoImage(file = path, master= root)
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

# Excel Automate
def excel_automation():
    if(s_filepath and d_filepath):
        excel_door_open(s_filepath, d_filepath)
        action_button.configure(bg = 'light green')
    else:
        messagebox.showwarning('Warning','Please Select Source and Destination.')
        action_button.configure(bg = 'red')

    root.destroy()

# Excel to PDF
def excel_to_pdf():
    # base_file_name = os.path.basename(s_filepath).split('.')[0] 
    # updated_file_name = d_filepath + "/" +"{}_Updated.xlsx".format(base_file_name)

    # excel_path_source = Path(updated_file_name)

    # if(excel_path_source.is_file()):
    if(s_filepath and d_filepath):
        excel_conversion(s_filepath, d_filepath)
        root.destroy()
    else:
        messagebox.showwarning('Warning','Please Select Source and Destination.')

def user_manual():
    messagebox.showinfo('User Manual', """Step 1 - Select Source
For first time select Source as excel file in which you want to perform operations below,
| Add Worksheets | Add Hyperlink | Add Formula | Add Home |

Step 2 - Select Destination
Select destination folder. In this folder _Updated.xlsx will be
saved after performing operations listed in Step 1.

Step 3 - Click Action
Perform operations mentioned in Step 1 and save result on destination 
with name mentioned in Step 2.

[Window Will Be Close After Step 3]
Update the generated excel file.

After updating open exe file again.

Step 4 - Select Source
Now select updated excel file for which you want to generate PDF of each 
sheetname.

Step 5 - Select Destination
Select Desination folder in which, one folder will be created as _PDF Files 

Step 6 - Click To PDF
Generate PDF of each sheet present in select excel file in destination folder.

[Window Will Be Close After Step 6.]


Note 
1. In Destination Folder Result.xlsx and 
    Result_hyperlinked.xlsx file will generate and delete 
    automatically. 
2. User Manual 
    Will give you user manual of application.
3. PDF Conversion will take more time you can see the result in
    PDF_Files Folder.
""")


# Create a Button - Take Source File
source_button = Button(root, text="Source", command = s_open_file)
source_button.place(x = 94, y = 50)

# Create a Button - Select Destination Path
destination_button = Button(root, text="Destination", command = d_open_file)
destination_button.place(x = 83, y = 80)

# Create a Button - Action Button
action_button = Button(root, text = "Action", command = excel_automation)
action_button.place(x = 68, y = 120)

# Create a Button - To PDF
excel_to_pdf_button = Button(root, text = "To PDF", command = excel_to_pdf)
excel_to_pdf_button.place(x = 125, y = 120)

# Label - Branding
signature = Label(root, text = "vConstruct-DPR", bg = 'green', fg = 'white')
signature.place(x = 75, y = 205)

# Create a Button - User Manual 
user_manual = Button(root, text = "User Manual", command = user_manual)
user_manual.place(x = 83, y = 240)

messagebox.showinfo('User Manual', """Step 1 - Select Source
For first time select Source as excel file in which you want to perform operations below,
| Add Worksheets | Add Hyperlink | Add Formula | Add Home |

Step 2 - Select Destination
Select destination folder. In this folder _Updated.xlsx will be
saved after performing operations listed in Step 1.

Step 3 - Click Action
Perform operations mentioned in Step 1 and save result on destination 
with name mentioned in Step 2.

[Window Will Be Close After Step 3]
Update the generated excel file.

After updating open exe file again.

Step 4 - Select Source
Now select updated excel file for which you want to generate PDF of each 
sheetname.

Step 5 - Select Destination
Select Desination folder in which, one folder will be created as _PDF Files 

Step 6 - Click To PDF
Generate PDF of each sheet present in select excel file in destination folder.

[Window Will Be Close After Step 6.]

Note 
1. In Destination Folder Result.xlsx and 
    Result_hyperlinked.xlsx file will generate and delete 
    automatically. 
2. User Manual 
    Will give you user manual of application.
3. PDF Conversion will take more time you can see the result in
    PDF_Files Folder.
""")


root.mainloop()