from Tkinter import Tk, Label, Button, StringVar, Entry, OptionMenu
import ttk
from tkFileDialog import asksaveasfile
from Tkconstants import LEFT, RIGHT
import tkFileDialog
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


class Application:
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
        
    def cycle_label_text(self):
        self.label_index += 1
        self.label_index %= len(self.LABEL_TEXT)
        self.label_text.set(self.LABEL_TEXT[self.label_index])
        
root = Tk()
root.geometry('600x750')
root.title('Adap')
#gui = GUI(root)
"""
rows=0
while rows<50:
    root.rowconfigure(rows, weight=1)
    root.columnconfigure(rows, weight=2)
    rows += 1
   """ 
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


"""
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
"""
def savefile():
    filename = tkFileDialog.asksaveasfilename(initialdir = "/",title = "Save data",filetypes = (("Microsoft Excel file","*.xlsx"),("all files","*.*")))
    print filename
    
"""
r = Label(f1, text="Start Volt")
r.pack(padx=5, pady=0, side=LEFT)
r = Entry(f1, textvariable=svalue)
r.pack(padx=5, pady=0, side=LEFT)
r = Label(f1, text="V")
r.pack(padx=5, pady=0, side=LEFT)
"""
start_volt = StringVar()
end_volt = StringVar()
step_volt = StringVar()
hold_time = StringVar()
compliance = StringVar()

s = Label(f1, text="Start Volt")
s.pack(side=LEFT)
s.grid(row=0, column=1)

s = Entry(f1, textvariable = start_volt)
s.pack(side=LEFT)
s.grid(row=0, column=2)

s = Label(f1, text="V")
s.pack(side=LEFT)
s.grid(row=0, column=3)

s = Label(f1, text="End Volt")
s.pack(side=LEFT)
s.grid(row=1, column=1)

s = Entry(f1, textvariable = end_volt)
s.pack(side=LEFT)
s.grid(row=1, column=2)

s = Label(f1, text="V")
s.pack(side=LEFT)
s.grid(row=1, column=3)

s = Label(f1, text="Step Volt")
s.pack(side=LEFT)
s.grid(row=2, column=1)

s = Entry(f1, textvariable = step_volt)
s.pack(side=LEFT)
s.grid(row=2, column=2)

s = Label(f1, text="V")
s.pack(side=LEFT)
s.grid(row=2, column=3)

s = Label(f1, text="Hold Time")
s.pack(side=LEFT)
s.grid(row=3, column=1)

s = Entry(f1, textvariable = hold_time)
s.pack(side=LEFT)
s.grid(row=3, column=2)

s = Label(f1, text="s")
s.pack(side=LEFT)
s.grid(row=3, column=3)

s = Label(f1, text="Compliance")
s.pack(side=LEFT)
s.grid(row=4, column=1)

s = Entry(f1, textvariable = compliance)
s.pack(side=LEFT)
s.grid(row=4, column=2)

compliance_choices = {'mA', 'uA', 'nA'}
compliance_scale = StringVar()
compliance_scale.set('uA')
s = OptionMenu(f1, compliance_scale, *compliance_choices)
s.pack(side=LEFT)
s.grid(row=4, column=3)

s = Label(f1, text="Percent Complete:")
s.pack(side=LEFT)
s.grid(row=11, column=1)

pb = ttk.Progressbar(f1, orient="horizontal", length=200, mode="determinate")
pb.pack(side=LEFT)
pb.grid(row=11, column= 2, columnspan=5)
pb["maximum"] = 100
pb["value"] = 90


f = Figure(figsize=(5,5), dpi=100)
a = f.add_subplot(111)
a.plot([1,2,3,4,5,6,7,8],[5,6,1,3,8,9,3,5])

canvas = FigureCanvasTkAgg(f, master=f1)
canvas.get_tk_widget().grid(row=10, columnspan=10)
canvas.draw()

def getvalues():
    try:
        x = float(start_volt.get())
        print "Start_volt= "+str(x)
    except ValueError:
        print "Please fill in all fields!"

s = Button(f1, text="Start", command=getvalues)
s.pack(side=RIGHT)
s.grid(row=2, column=5)

def file_save():
    fout = asksaveasfile(mode='w', defaultextension=".txt")
    text2save = "test world"
    fout.write(text2save)
    fout.close()
   
""" 
menu = Menu(f1)
root.config(menu=menu)
filemenu = Menu(f1, tearoff=0)
menu.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="New")
filemenu.add_command(label="Save", command = file_save)
filemenu.add_separator()
filemenu.add_command(label="Exit", command = do_exit)
"""


root.mainloop()

