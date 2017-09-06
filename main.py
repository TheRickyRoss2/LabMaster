from Keithley import Keithley2400, Keithley2657a
from Agilent import AgilentE4980a, Agilent4156
from emailbot import sendMail
from Tkinter import Tk, Label, Button, StringVar, Entry, OptionMenu
import ttk
from Tkconstants import LEFT, RIGHT
import matplotlib
import threading
from random import randint
matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import pyplot as plt
import time
import visa
import tkFileDialog
import xlsxwriter
import Queue
import random

test=True

rm = visa.ResourceManager()
print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")

def GetIV(sourceparam, sourcemeter, dataout):
    (start_volt, end_volt, step_volt, delay_time, compliance) = sourceparam
    
    currents = []
    voltages = []
    keithley = 0
    if sourcemeter is 0:
        keithley = Keithley2400()
    else:
        keithley = Keithley2657a()
    
    if test is True:
        pass
    else:
        keithley.init()
        keithley.configure_measurement(1, 0, compliance)
    last_volt = 0
    badCount = 0
    
    if start_volt>end_volt:
        step_volt = -1*step_volt
    
    print "looping now"
    for volt in xrange(start_volt, end_volt, step_volt):
        curr = 0
        if test is True:
            pass
        else:
            keithley.set_output(volt)
            
        time.sleep(delay_time)
        
        if test is True:
            curr = volt+randint(0, 10)
        else:
            curr = keithley.get_current()
        #curr = volt
        
        if(abs(curr-compliance)<10e-9):
            badCount = badCount + 1
        
        if(badCount>=5):
            print "Compliance reached"
            break
        
        currents.append(curr)
        voltages.append(volt)
        """
        #TODO Live graphics
        Y = currents
        graph.set_ydata(Y)
        graph.draw()
        """
        last_volt = volt
        dataout.put(((voltages, currents), 100*abs((volt+step_volt)/float(end_volt))))
        
        #pb["value"] = volt/(end_volt-step_volt)
        
        #graph point here
        
    while abs(last_volt)>5:
        if test is True:
            pass
        else:
            keithley.set_output(last_volt)
        time.sleep(delay_time/2.0)
        if last_volt < 0:
            last_volt +=5
        else:
            last_volt -=5
        
    if test is True:
        pass
    else:
        keithley.set_output(0)
        keithley.enable_output(False)
    return (voltages, currents)

def GetCV(sourceparam, lcrparam, sourcemeter, dataout):
    capacitance = []
    voltages = []
    
    keithley = 0
    if sourcemeter is 0:
        keithley = Keithley2400()
    else:
        keithley = Keithley2657a()
    keithley.init()
    keithley.configure_measurement()
    last_volt = 0
    (frequencies, meas_time, avg_factor, signal_type, level, function, impedance) = lcrparam
    
    (start_volt, end_volt, step_volt, hold_time, delay_time, compliance) = sourceparam
    
    agilent = AgilentE4980a()
    agilent.init()
    agilent.configure_measurement(function, impedance)
    agilent.configure_aperture(meas_time, avg_factor)
    badCount = 0
    if start_volt>end_volt:
        step_volt = -1*step_volt
    
    for volt in xrange(start_volt, end_volt, step_volt):
        
        keithley.set_output(volt)
        time.sleep(delay_time)
        for f in frequencies:
            agilent.configure_measurement_signal(f, signal_type, level)
            data = agilent.read_data()
            capacitance.append(data)
        
        
        curr = keithley.read_single_point(hold_time)
        if(curr[1] > compliance):
            badCount = badCount + 1
            
        if(badCount>=5):
            print "Compliance reached"
            break
        last_volt = volt
        
        #graph point here
    for volt in xrange(last_volt, start_volt, step_volt*-2):
        keithley.set_output(volt)
        time.sleep(delay_time/2.0)
    
    keithley.enable_output(False)
    
    for i in xrange(0,len(frequencies), 1):
        print "Frequency: " +str(frequencies[i]) 
        print capacitance[i::len(frequencies)]
        
    return capacitance

def spa_iv(sourceparam, meas_param):
    (start_volt, end_volt, step_volt, delay_time, compliance) = sourceparam

    current_smu1 = []
    current_smu2 = []
    current_source = []
    voltage_source = Keithley2657a()
    voltage_source.init()
    voltage_source.configure_measurement()
    voltage_source.configure_source(0.1)
    voltage_source.enable_output(True)
    
    daq = Agilent4156()
    daq.init()
    daq.configure_integration_time()
    for i in xrange(0, 4, 1):
        daq.configure_channel(i)

    for x in xrange(0, 10, 1):

        voltage_source.set_output(x)
        time.sleep(0.5)
        daq.configure_measurement()
        daq.configure_sampling_measurement()
        daq.configure_sampling_stop()
        if x is 0:
            daq.inst.write(":PAGE:DISP:GRAP:Y2:NAME \'I2\';")
            print daq.inst.query(":PAGE:DISP:LIST?")
            daq.inst.write(":PAGE:DISP:LIST \'@TIME\', \'I1\', \'I2\'")
        
        daq.measurement_actions()
        daq.wait_for_acquisition()
        
        current_smu1.append(daq.read_trace_data("I1"))
        current_smu2.append(daq.read_trace_data("I2"))
        current_source.append(voltage_source.get_current())
        print current_smu1
        print current_smu2
        print current_source
        

    for x in xrange(0, 10, -2*1):
        voltage_source.set_output(x)
        time.sleep(delay_time)
        
    voltage_source.set_output(0)
    voltage_source.enable_output(False)
    print current_smu1
    print current_smu2
    print current_source    



class GuiPart:
    
    def __init__(self, master, inputdata, outputdata, endcommand):
        print "in guipart"
        
        self.inputdata = inputdata
        self.outputdata = outputdata
        self.start_volt = StringVar()
        self.end_volt = StringVar()
        self.step_volt = StringVar()
        self.hold_time = StringVar()
        self.compliance = StringVar() 
        self.recipients = StringVar()   
        self.compliance_scale = StringVar()
        self.source_choice = StringVar()
        
        self.cv_start_volt = StringVar()
        self.cv_end_volt = StringVar()
        self.cv_step_volt = StringVar()
        self.cv_hold_time = StringVar()
        self.cv_compliance = StringVar() 
        self.cv_recipients = StringVar()   
        self.cv_compliance_scale = StringVar()
        self.cv_
        self.cv_source_choice = StringVar()
        
        self.started = False
        
        self.start_volt.set("0.0")
        self.end_volt.set("100.0")
        self.step_volt.set("5.0")
        self.hold_time.set("1.0")
        self.compliance.set("1.0")
        
        self.f = plt.figure(figsize=(3,3), dpi=100)
        self.a = self.f.add_subplot(111)

        n = ttk.Notebook(root)
        n.grid(row=1, column=0, columnspan=60, rowspan=60, sticky='NESW')
        f1 = ttk.Frame(n)
        f2 = ttk.Frame(n)
        f3 = ttk.Frame(n)
        n.add(f1, text='IV')
        n.add(f2, text='CV')
        n.add(f3, text='SPA IV')
        
        s = Label(f1, text="Start Volt")
        s.pack(side=LEFT)
        s.grid(row=0, column=1)
        
        s = Entry(f1, textvariable = self.start_volt)
        s.pack(side=LEFT)
        s.grid(row=0, column=2)
        
        s = Label(f1, text="V")
        s.pack(side=LEFT)
        s.grid(row=0, column=3)
        
        s = Label(f1, text="End Volt")
        s.pack(side=LEFT)
        s.grid(row=1, column=1)
        
        s = Entry(f1, textvariable = self.end_volt)
        s.pack(side=LEFT)
        s.grid(row=1, column=2)
        
        s = Label(f1, text="V")
        s.pack(side=LEFT)
        s.grid(row=1, column=3)
        
        s = Label(f1, text="Step Volt")
        s.pack(side=LEFT)
        s.grid(row=2, column=1)
        
        s = Entry(f1, textvariable = self.step_volt)
        s.pack(side=LEFT)
        s.grid(row=2, column=2)
        
        s = Label(f1, text="V")
        s.pack(side=LEFT)
        s.grid(row=2, column=3)
        
        s = Label(f1, text="Hold Time")
        s.pack(side=LEFT)
        s.grid(row=3, column=1)
        
        s = Entry(f1, textvariable = self.hold_time)
        s.pack(side=LEFT)
        s.grid(row=3, column=2)
        
        s = Label(f1, text="s")
        s.pack(side=LEFT)
        s.grid(row=3, column=3)
        
        s = Label(f1, text="Compliance")
        s.pack(side=LEFT)
        s.grid(row=4, column=1)
        
        s = Entry(f1, textvariable = self.compliance)
        s.pack(side=LEFT)
        s.grid(row=4, column=2)
        
        compliance_choices = {'mA', 'uA', 'nA'}
        self.compliance_scale.set('uA')
        s = OptionMenu(f1, self.compliance_scale, *compliance_choices)
        s.pack(side=LEFT)
        s.grid(row=4, column=3)
        
        s = Label(f1, text="Email data to:")
        s.pack(side=LEFT)
        s.grid(row=5, column=1)
        
        s = Entry(f1, textvariable = self.recipients)
        s.pack(side=LEFT)
        s.grid(row=5, column=2)
        
        s = Label(f1, text="Separate with ','")
        s.pack(side=LEFT)
        s.grid(row=5, column=3)
    
        source_choices = {'Keithley 2400', 'Keithley 2657a'}
        self.source_choice.set('Keithley 2657a')
        s = OptionMenu(f1, self.source_choice, *source_choices)
        s.pack(side=LEFT)
        s.grid(row=0, column=7)
        
        s = Label(f1, text="Progress:")
        s.pack(side=LEFT)
        s.grid(row=11, column=1)
        
        self.pb = ttk.Progressbar(f1, orient="horizontal", length=200, mode="determinate")
        self.pb.pack(side=LEFT)
        self.pb.grid(row=11, column= 2, columnspan=5)
        self.pb["maximum"] = 100
        self.pb["value"] = 0
        
        self.canvas = FigureCanvasTkAgg(self.f, master=f1)
        self.canvas.get_tk_widget().grid(row=6, columnspan=10)
        plt.xlabel("Voltage")
        plt.ylabel("Current")
        plt.title("IV")
        #plt.show()
        self.canvas.draw()
        
        s = Button(f1, text="Start IV", command=self.prepare_values)
        s.pack(side=RIGHT)
        s.grid(row=3, column=7)
        
        s = Button(f1, text="Stop", command=endcommand)
        s.pack(side=RIGHT)
        s.grid(row=4, column=7)
        print "finished drawing"
        
        
    def update(self):
        while self.outputdata.qsize():
            try:
                (data, percent) = self.outputdata.get(0)
                print "Percent done:" +str(percent)
                self.pb["value"] = percent
                self.pb.update()
                (voltages, currents) = data
                self.a.plot(voltages, currents)
                self.canvas.draw()

            except Queue.Empty:
                pass
    
    def quit(self):
        import sys
        sys.exit(1)
        #root.destroy()
    
    def prepare_values(self):
        print "preparing values"
        input_params = (self.compliance.get(), self.compliance_scale.get(), self.start_volt.get(), self.end_volt.get(), self.step_volt.get(), self.hold_time.get(), self.source_choice.get())
        #getvalues(self.input_params)
        self.inputdata.put(input_params)   
         
        
def getvalues(input_params, dataout):
    print input_params
    filename = tkFileDialog.asksaveasfilename(initialdir = "/",title = "Save data",filetypes = (("Microsoft Excel file","*.xlsx"),("all files","*.*")))
    (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice) = input_params
    
    try:
        comp = float(float(compliance)*({'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale, 1e-6)))
        source_params = (int(float(start_volt)), int(float(end_volt)), int(float(step_volt)),
                             float(hold_time), comp)
        print source_params
    except ValueError:
        print "Please fill in all fields!"
    data = ()
    if source_params is None:
        pass
    else:
        data = GetIV(source_params, 1, dataout)
            
    data_out = xlsxwriter.Workbook(filename+".xlsx")
    path = filename+".xlsx"
    worksheet = data_out.add_worksheet()
    (v, i) = data
    values = []
    for x in xrange(0,len(v), 1):
        values.append((v[x], i[x]))
    row=0
    col=0
    chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
    
    for volt, cur in values:
        worksheet.write(row, col, volt)
        worksheet.write(row, col+1, cur)
        row+=1
    
    chart.add_series({'categories': '=Sheet1!$A$1:$A$'+str(row), 'values': '=Sheet1!$B$1:$B$'+str(row)})
    chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
    chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
    chart.set_legend({'none':True})
    worksheet.insert_chart('D2', chart)
    data_out.close()
    
    try:
        mails = recipients.get().split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())
                
        print sentTo
        sendMail(path, sentTo)
    except:
        pass
        
    print data

class ThreadedProgram:
    
    def __init__(self, master):
        self.master = master
        self.inputdata = Queue.Queue()
        self.outputdata = Queue.Queue()
        print "making gui"
        
        self.running = 1
        self.gui = GuiPart(master, self.inputdata, self.outputdata, self.endapp)

        self.thread1 = threading.Thread(target=self.workerThread1)
        self.thread1.start()
        self.periodicCall()
        self.measuring = False
        
    def periodicCall(self):
        self.gui.update()
        if not self.running:
            import sys
            sys.exit(1)
        self.master.after(200, self.periodicCall)
        
    def workerThread1(self):
        while self.running:
            #print "looping"
            if self.inputdata.empty() is False and self.measuring is False:
                self.measuring= True
                print "doing stuff"
                print self.inputdata
                getvalues(self.inputdata.get(), self.outputdata)
                self.measuring=False
        
    def endapp(self):
        self.running = 0
        
if __name__=="__main__":
    
    root = Tk()
    root.geometry('610x570')
    root.title('Adap')
    client = ThreadedProgram(root)
    root.mainloop() 
    
    """
    
    agilent = Agilent4156()
    agilent.configure_vmu(discharge=True, _vmu=2, _mode = 1, name="GOOD")
    agilent.read_trace_data("VM1")
    """