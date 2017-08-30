from Tkinter import Tk, Label, Button, StringVar, Entry, OptionMenu, Checkbutton, IntVar, Menu
import ttk
from tkFileDialog import *
from Tkconstants import VERTICAL, LEFT
class GUI:
    LABEL_TEXT = [
        "Test string 1",
        "Second test string",
        "Stringeroni",
        "Pepperoni",
        "Kekeronei",
        ]
    
    def __init__(self, master):
        master.title("BADAP")
        self.label_index = 0
        self.label_text = StringVar()
        self.label_text.set(self.LABEL_TEXT[self.label_index])
        self.label = Label(master, textvariable=self.label_text)
        self.label.bind("<Button-1>", self.cycle_label_text)
        self.label.pack()
        
        self.greet_button =  Button(master, text="Hello", command = self.greet)
        self.greet_button.pack()
        
        self.close_button = Button(master, text = "Close", command = master.quit)
        self.close_button.pack()
        

    def greet(self):
        print "Hi!!!"
        
    def cycle_label_text(self, event):
        self.label_index += 1
        self.label_index %= len(self.LABEL_TEXT)
        self.label_text.set(self.LABEL_TEXT[self.label_index])
        
root = Tk()
root.geometry('1024x768')
root.title('BADAP')
#gui = GUI(root)
rows=0
while rows<50:
    root.rowconfigure(rows, weight=1)
    root.columnconfigure(rows, weight=2)
    rows += 1
    
n = ttk.Notebook(root)

n.grid(row=1, column=0, columnspan=60, rowspan=60, sticky='NESW')

f1 = ttk.Frame(n)
f2 = ttk.Frame(n)
f3 = ttk.Frame(n)
f4 = ttk.Frame(n)
"""
from Tkinter import *
root = Tk(className ="My first GUI")
svalue = StringVar() # defines the widget state as string
w = Entry(root,textvariable=svalue) # adds a textarea widget
w.pack()
def act():
    print "you entered"
    print '%s' % svalue.get()
foo = Button(root,text="Press Me", command=act)
foo.pack()
root.mainloop()"""

n.add(f1, text='IV')
n.add(f2, text='CV')
n.add(f3, text='SPA IV')
n.add(f4, text='Bot')

svalue = StringVar()
def act():
    print "Entered"+svalue.get()
    if use is 1:
        print "USING"
    else:
        print "NOT USING"
foo = Button(f1, text="Start", command = act)
foo.pack()

lst1 = ['Option1','Option2','Option3']
var1 = StringVar()
var1.set(lst1[0])
drop = OptionMenu(f1,var1, *lst1)
drop.pack()

lbl = Label(f1, text="Start Volt")
lbl.pack()

use = IntVar()
use.set(0)
c = Checkbutton(f1, text="Email?", variable = use)
c.pack()

r = Label(f1, text="test1")
r.pack(padx=5, pady=0, side=LEFT)
r = Entry(f1, textvariable=svalue)
r.pack(padx=5, pady=0, side=LEFT)
def do_exit():
    root.destroy()
def file_save():
    fout = asksaveasfile(mode='w', defaultextension=".txt")
    text2save = "test world"
    fout.write(text2save)
    fout.close()
    
menu = Menu(f1)
root.config(menu=menu)
filemenu = Menu(f1, tearoff=0)
menu.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="New")
filemenu.add_command(label="Save", command = file_save)
filemenu.add_separator()
filemenu.add_command(label="Exit", command = do_exit)

root.mainloop()