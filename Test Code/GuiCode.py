import PySimpleGUI as sg


import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.filedialog import asksaveasfile
from tkinter.messagebox import showinfo
"""
sg.theme("DarkTeal2")
layout = [[sg.T("")], [sg.Text("Choose a file: "), sg.Input(), sg.FileBrowse(key="-IN-")],[sg.Button("Select")]]

###Building Window
window = sg.Window('My File Browser', layout, size=(600,150))
    
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Select":
        print(values["-IN-"])




sg.theme("DarkTeal2")
layout = [[sg.T("")], [sg.Text("Choose a folder: "), sg.Input(key="-IN2-" ,change_submits=True), sg.FolderBrowse(key="-IN-")],[sg.Button("Submit")]]

###Building Window
window = sg.Window('My File Browser', layout, size=(600,150))
    
while True:
    event, values = window.read()
    print(values["-IN2-"])
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Submit":
        print(values["-IN-"])

"""

#filename = fd.askopenfilename()




# create the root window

"""
root = Tk()
root.geometry('200x150')
  
# function to call when user press
# the save button, a filedialog will
# open and ask to save file
def save():
    files = [('All Files', '*.*'), 
             ('Python Files', '*.py'),
             ('Text Document', '*.txt')]
    file = asksaveasfile(filetypes = files, defaultextension = files)
  
btn = ttk.Button(root, text = 'Save', command = lambda : save())
btn.pack(side = TOP, pady = 20)
  
mainloop()
"""

fob = asksaveasfile(defaultextension=".txt",mode='w')
#fob = filedialog.asksaveasfilename(defaultextension=".txt")
#fob=open(fob,'w') # creates file object
fob.write("Welcome to plus2net")
fob.close()