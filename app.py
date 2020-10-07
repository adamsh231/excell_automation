# %%
from tkinter import filedialog
from tkinter import *
import calculation
from calculation import *

def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    print(filename)
    print(filename.replace("/", "//"))
    calculation.path = filename.replace("/", "//")

def calc():
    # print(calculation.path)
    generate()

root = Tk()
folder_path = StringVar()
lbl1 = Label(master=root,textvariable=folder_path)
lbl1.grid(row=0, column=1)
button2 = Button(text="Browse", command=browse_button)
button2.grid(row=0, column=3)
button3 = Button(text="Generate", command=calc)
button3.grid(row=1, column=3)
mainloop()