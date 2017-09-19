# -*- coding: utf-8 -*-
from Keithley import Keithley2400, Keithley2657a
from Agilent import AgilentE4980a, Agilent4156
from emailbot import sendMail
from Tkinter import Tk, Label, Button, StringVar, Entry, OptionMenu
import ttk
from Tkconstants import LEFT, RIGHT, RAISED
import matplotlib
import threading
from random import randint
from _ast import Param
from platform import platform
matplotlib.use("TkAgg")
import platform

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import pyplot as plt
import time
import visa
import tkFileDialog
import xlsxwriter
import Queue
import random

test = True
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
    print "Meter" + str(sourcemeter)
    
    if test:
        pass
    else:
        if sourcemeter is 0:
            keithley = Keithley2400()
        else:
            keithley = Keithley2657a()
        keithley.configure_measurement(1, 0, compliance)
    last_volt = 0
    badCount = 0
    
    scaled = False
    if step_volt < 1.0:
        start_volt *=1000
        end_volt *=1000
        step_volt*=1000
        scaled = True
    
    if start_volt>end_volt:
        step_volt = -1*step_volt
    
    print "looping now"
    
    for volt in xrange(start_volt, end_volt, int(step_volt)):
        start_time = time.time()
        
        curr = 0
        if test:
            pass
        else:
            if scaled:
                keithley.set_output(volt/1000.0)
            else:
                keithley.set_output(volt)
        
        time.sleep(delay_time)
        
        if test:
            curr = (volt+randint(0, 10))*1e-9
        else:
            curr = keithley.get_current()
        #curr = volt
        
        if abs(curr)>abs(compliance-50e-9):
            badCount = badCount + 1        
        else:
            badCount = 0    
        
        if badCount>=5 :
            print "Compliance reached"
            break
        
        currents.append(curr)
        if scaled:
            voltages.append(volt/1000.0)
        else:
            voltages.append(volt)

        if scaled:
            last_volt = volt/1000.0
        else:
            last_volt = volt
        
        time_remain = (time.time()-start_time)*(abs((end_volt-volt)/step_volt))
        
        dataout.put( ((voltages, currents), 100*abs((volt+step_volt)/float(end_volt)), time_remain) )
        
        
    while abs(last_volt)>5:
        if test:
            pass
        else:
            keithley.set_output(last_volt)
        
        time.sleep(0.5)
        if last_volt < 0:
            last_volt +=5
        else:
            last_volt -=5

    time.sleep(0.5)
    if test:
        pass
    else:
        keithley.set_output(0)
        keithley.enable_output(False)
    return (voltages, currents)

def GetCV(params, sourcemeter, dataout):
    
    capacitance = []
    voltages = []
    p2=  []
    c = []
    keithley = 0
    
    if test:
        pass
    else:
        if sourcemeter is 0:
            keithley = Keithley2400()
        else:
            keithley = Keithley2657a()
        keithley.configure_measurement()
    
    last_volt = 0
    
    (start_volt, end_volt, step_volt, delay_time, compliance,
     frequencies, level, function, impedance, int_time) = params
    
    if test:
        pass
    else:
        agilent = AgilentE4980a()
        agilent.init()
        agilent.configure_measurement(function)
        agilent.configure_aperture(int_time)
    badCount = 0
    
    scaled = False
    
    if step_volt < 1.0:
        start_volt *=1000
        end_volt *=1000
        step_volt*=1000
        scaled = True
    
    if start_volt>end_volt:
        step_volt = -1*step_volt
    
    start_time = time.time()
    for volt in xrange(start_volt, end_volt, step_volt):
    
        start_time = time.time()
        if test:
            pass
        else:
            if scaled:
                keithley.set_output(volt/1000.0)
            else:
                keithley.set_output(volt)
            
        curr = 0
        for f in frequencies:
            time.sleep(delay_time)

            if test:
                capacitance.append((volt+int(f)*randint(0, 10)))
                curr = volt*1e-10
                c.append(curr)
                p2.append(volt*10)
            else:
                agilent.configure_measurement_signal(float(f), 0, level)
                (data, aux) = agilent.read_data()
                capacitance.append(data)
                p2.append(aux)
                curr= keithley.get_current()
                c.append(curr)
        
            
        if abs(curr)>abs(compliance-50e-9):
            badCount = badCount + 1        
        else:
            badCount = 0    
        
        if badCount>=5 :
            print "Compliance reached"
            break
        
        time_remain = (time.time()-start_time)*(abs((end_volt-volt)/step_volt))
        
        if scaled:
            voltages.append(volt/1000.0)
        else:
            voltages.append(volt)
        formatted_cap = []
        parameter2 = []
        currents = []
        for i in xrange(0,len(frequencies), 1):
            formatted_cap.append(capacitance[i::len(frequencies)])
            parameter2.append(p2[i::len(frequencies)])
            currents.append(c[i::len(frequencies)])
        dataout.put(((voltages, formatted_cap), 100*abs((volt+step_volt)/float(end_volt)), time_remain))
        
        time_remain = time.time()+(time.time()-start_time)*(abs((volt-end_volt)/end_volt))
        
        last_volt = volt
        #graph point here
        
    if test:
        pass
    else:
        while last_volt > 0:
            if last_volt<=5:
                keithley.set_output(0)
                last_volt = 0
            else:
                keithley.set_output(last_volt-5)
                last_volt -= 5
                
            time.sleep(0.5)
    
    if test:
        pass
    else:
        keithley.enable_output(False)
    
    
    return (voltages, currents, formatted_cap, parameter2)

def spa_iv(params, dataout):
    (start_volt, end_volt, step_volt, hold_time, compliance, int_time) = params

    print params
    voltage_smua = []
    current_smua = []
    
    current_smu1 = []
    current_smu2 = []
    current_smu3 = []
    current_smu4 = []
    voltage_vmu1 = []

    voltage_source = Keithley2657a()
    voltage_source.configure_measurement(1, 0, compliance)
    voltage_source.enable_output(True)
    
    daq = Agilent4156()
    daq.configure_integration_time(_int_time = int_time)

    scaled = False
    if step_volt < 1.0:
        start_volt *=1000
        end_volt *=1000
        step_volt*=1000
        scaled = True
    
    if start_volt>end_volt:
        step_volt = -1*step_volt
        
    for i in xrange(0, 4, 1):
        daq.configure_channel(i)  
    daq.configure_vmu()    
          
    last_volt = 0
    for volt in xrange(start_volt, end_volt, step_volt):

        if test:
            pass
        else:
            if scaled:
                voltage_source.set_output(volt/1000.0)
            else:
                voltage_source.set_output(volt)
        time.sleep(hold_time)
        
        daq.configure_measurement()
        daq.configure_sampling_measurement()
        daq.configure_sampling_stop()
        
        #daq.inst.write(":PAGE:DISP:GRAP:Y2:NAME \'I2\';")
        daq.inst.write(":PAGE:DISP:LIST \'@TIME\', \'I1\', \'I2\', \'I3\', \'I4\', \'VMU1\'")
        daq.measurement_actions()
        daq.wait_for_acquisition()
        
        current_smu1.append(daq.read_trace_data("I1"))
        current_smu2.append(daq.read_trace_data("I2"))
        
        #daq.inst.write(":PAGE:DISP:LIST \'@TIME\', \'I2\', \'I3\'")
      

        current_smu3.append(daq.read_trace_data("I3"))
        current_smu4.append(daq.read_trace_data("I4"))
        voltage_vmu1.append(daq.read_trace_data("VMU1"))
        current_smua.append(voltage_source.get_current())
        
        if scaled:
            voltage_smua.append(volt/1000.0)
            last_volt = volt/1000.0
        else:
            voltage_smua.append(volt)
            last_volt = volt
        
        print "SMU1-4"
        print current_smu1
        print current_smu2
        print current_smu3
        print current_smu4
        print "SMUA"
        print current_smua
        print "VMU1"
        print voltage_vmu1
        dataout.put((voltage_vmu1, current_smua, current_smu1, current_smu2, current_smu3, current_smu4))
        
        
    while abs(last_volt)>=4:
        time.sleep(0.5)

        if test:
            pass
        else:
            voltage_source.set_output(last_volt)
        
        if last_volt < 0:
            last_volt +=5
        else:
            last_volt -=5
        

    time.sleep(0.5)
    voltage_source.set_output(0)
    voltage_source.enable_output(False)
    
    
    return(voltage_smua, current_smua, current_smu1, current_smu2, current_smu3, current_smu4, voltage_vmu1)

class GuiPart:
    
    def __init__(self, master, inputdata, outputdata, stopq):
        print "in guipart"
        self.master = master
        self.inputdata = inputdata
        self.outputdata = outputdata
        self.stop = stopq
        
        self.start_volt = StringVar()
        self.end_volt = StringVar()
        self.step_volt = StringVar()
        self.hold_time = StringVar()
        self.compliance = StringVar() 
        self.recipients = StringVar()   
        self.compliance_scale = StringVar()
        self.source_choice = StringVar()
        self.filename = StringVar()
        
        self.cv_start_volt = StringVar()
        self.cv_end_volt = StringVar()
        self.cv_step_volt = StringVar()
        self.cv_hold_time = StringVar()
        self.cv_compliance = StringVar() 
        self.cv_recipients = StringVar()   
        self.cv_compliance_scale = StringVar()
        self.cv_source_choice = StringVar()
        self.cv_impedance_scale = StringVar()
        self.cv_amplitude = StringVar()
        self.cv_frequencies = StringVar()
        self.cv_integration = StringVar()
        self.started = False
        
        self.start_volt.set("0.0")
        self.end_volt.set("100.0")
        self.step_volt.set("5.0")
        self.hold_time.set("1.0")
        self.compliance.set("1.0")
        
        
        self.f = plt.figure(figsize=(6, 4), dpi=60)
        self.a = self.f.add_subplot(111)
        
        self.cv_f = plt.figure(figsize=(6, 4), dpi=60)
        self.cv_a = self.cv_f.add_subplot(111)
        
        n = ttk.Notebook(root, width=800)
        n.grid(row=0, column=0, columnspan=100, rowspan=100, sticky='NESW')
        f1 = ttk.Frame(n)
        f2 = ttk.Frame(n)
        f3 = ttk.Frame(n)
        n.add(f1, text='IV')
        n.add(f2, text='CV')
        n.add(f3, text='SPA')
        
        self.f1 = f1
        self.f2 = f2
        self.f3 = f3
        if "Windows" in platform.platform():
            self.filename.set("iv_data.xlsx")
            s = Label(f1, text="File name:")
            s.grid(row=0, column=1)
            s = Entry(f1, textvariable = self.filename)
            s.grid(row=0, column=2)
        
        s = Label(f1, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(f1, textvariable = self.start_volt)
        s.grid(row=1, column=2)
        s = Label(f1, text="V")
        s.grid(row=1, column=3)
        
        s = Label(f1, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(f1, textvariable = self.end_volt)
        s.grid(row=2, column=2)
        s = Label(f1, text="V")
        s.grid(row=2, column=3)
        
        s = Label(f1, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(f1, textvariable = self.step_volt)
        s.grid(row=3, column=2)
        s = Label(f1, text="V")
        s.grid(row=3, column=3)
        
        s = Label(f1, text="Hold Time")
        s.grid(row=4, column=1)
        s = Entry(f1, textvariable = self.hold_time)
        s.grid(row=4, column=2)
        s = Label(f1, text="s")
        s.grid(row=4, column=3)
        
        s = Label(f1, text="Compliance")
        s.grid(row=5, column=1)
        s = Entry(f1, textvariable = self.compliance)
        s.grid(row=5, column=2)
        compliance_choices = {'mA', 'uA', 'nA'}
        self.compliance_scale.set('uA')
        s = OptionMenu(f1, self.compliance_scale, *compliance_choices)
        s.grid(row=5, column=3)
        
        self.recipients.set("adapbot@gmail.com")
        s = Label(f1, text="Email data to:")
        s.grid(row=6, column=1)
        s = Entry(f1, textvariable = self.recipients)
        s.grid(row=6, column=2)
    
        source_choices = {'Keithley 2400', 'Keithley 2657a'}
        self.source_choice.set('Keithley 2657a')
        s = OptionMenu(f1, self.source_choice, *source_choices)
        s.grid(row=0, column=7)
        
        s = Label(f1, text="Progress:")
        s.grid(row=11, column=1)
       
        s = Label(f1, text="Est finish at:")
        s.grid(row=12, column=1)
        
        timetext = str(time.asctime(time.localtime(time.time())))
        self.timer = Label(f1, text=timetext)
        self.timer.grid(row=12, column=2)
        
        self.pb = ttk.Progressbar(f1, orient="horizontal", length=200, mode="determinate")
        self.pb.grid(row=11, column= 2, columnspan=5)
        self.pb["maximum"] = 100
        self.pb["value"] = 0
        
        self.canvas = FigureCanvasTkAgg(self.f, master=f1)
        self.canvas.get_tk_widget().grid(row=7, columnspan=10)
        self.a.set_title("IV")
        self.a.set_xlabel("Voltage")
        self.a.set_ylabel("Current")

        #plt.xlabel("Voltage")
        #plt.ylabel("Current")
        #plt.title("IV")
        self.canvas.draw()
        
        s = Button(f1, text="Start IV", command=self.prepare_values)
        s.grid(row=3, column=7)
        
        s = Button(f1, text="Stop", command=self.quit)
        s.grid(row=4, column=7)
        
        """
        /***********************************************************
         * CV GUI
         **********************************************************/
        """
        self.cv_filename = StringVar()
        self.cv_filename.set("cv_data.xlsx")
        
        if "Windows" in platform.platform():
            s = Label(f2, text="File name")
            s.grid(row=0, column=1)
            s = Entry(f2, textvariable = self.cv_filename)
            s.grid(row=0, column=2)
        
        self.cv_start_volt.set("0.0")
        s = Label(f2, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(f2, textvariable = self.cv_start_volt)
        s.grid(row=1, column=2)       
        s = Label(f2, text="V")
        s.grid(row=1, column=3)
        
        self.cv_end_volt.set("40.0")
        s = Label(f2, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(f2, textvariable = self.cv_end_volt)
        s.grid(row=2, column=2)
        s = Label(f2, text="V")
        s.grid(row=2, column=3)
        
        self.cv_step_volt.set("1.0")
        s = Label(f2, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(f2, textvariable = self.cv_step_volt)
        s.grid(row=3, column=2)
        s = Label(f2, text="V")
        s.grid(row=3, column=3)
        
        self.cv_hold_time.set("1.0")
        s = Label(f2, text="Hold Time")
        s.grid(row=4, column=1)
        s = Entry(f2, textvariable = self.cv_hold_time)
        s.grid(row=4, column=2)
        s = Label(f2, text="s")
        s.grid(row=4, column=3)
        
        self.cv_compliance.set("1.0")
        s = Label(f2, text="Compliance")
        s.grid(row=5, column=1)
        s = Entry(f2, textvariable = self.cv_compliance)
        s.grid(row=5, column=2)
        self.cv_compliance_scale.set('uA')
        s = OptionMenu(f2, self.cv_compliance_scale, *compliance_choices)
        s.grid(row=5, column=3)
        
        self.cv_recipients.set("adapbot@gmail.com")
        s = Label(f2, text="Email data to:")
        s.grid(row=6, column=1)
        s = Entry(f2, textvariable = self.cv_recipients)
        s.grid(row=6, column=2)
        
        s = Label(f2, text="Agilent LCRMeter Parameters", relief=RAISED)
        s.grid(row=7, column=1, columnspan=2)
        
        self.cv_impedance = StringVar()
        s = Label(f2, text="Function")
        s.grid(row=8, column=1)
        function_choices = {"CPD", "CPQ",  "CPG",   "CPRP",  "CSD",  "CSQ", "CSRS",   "LPD",
                 "LPQ", "LPG", "LPRP", "LPRD", "LSD", "LSQ", "LSRS", "LSRD",
                 "RX", "ZTD", "ZTR", "GB",   "YTD", "YTR", "VDID"}
        self.cv_function_choice = StringVar()
        self.cv_function_choice.set('CPD')
        s = OptionMenu(f2, self.cv_function_choice, *function_choices)
        s.grid(row=8, column=2)
        
        self.cv_impedance.set("2000")
        s = Label(f2, text="Impedance")
        s.grid(row=9, column=1)
        s = Entry(f2, textvariable=self.cv_impedance )
        s.grid(row=9, column=2)
        s = Label(f2, text="Ω")
        s.grid(row=9, column=3)
        
        self.cv_frequencies.set("100, 200, 1000, 2000")
        s = Label(f2, text="Frequencies")
        s.grid(row=10, column=1)
        s = Entry(f2, textvariable=self.cv_frequencies )
        s.grid(row=10, column=2)
        s = Label(f2, text="Hz")
        s.grid(row=10, column=3)
        
        self.cv_amplitude.set("5.0")
        s = Label(f2, text="Signal Amplitude")
        s.grid(row=11, column=1)
        s = Entry(f2, textvariable=self.cv_amplitude )
        s.grid(row=11, column=2)
        s = Label(f2, text="V")
        s.grid(row=11, column=3)
        
        cv_int_choices = {"Short", "Medium", "Long"}
        s = Label(f2, text="Integration time")
        s.grid(row=12, column=1)
        self.cv_integration.set("Short")
        s = OptionMenu(f2, self.cv_integration, *cv_int_choices)
        s.grid(row=12, column=2)
    
        self.cv_source_choice.set('Keithley 2657a')
        s = OptionMenu(f2, self.cv_source_choice, *source_choices)
        s.grid(row=0, column=7)
        
        s = Label(f2, text="Progress:")
        s.grid(row=14, column=1)
        
        self.cv_pb = ttk.Progressbar(f2, orient="horizontal", length=200, mode="determinate")
        self.cv_pb.grid(row=14, column= 2, columnspan=5)
        self.cv_pb["maximum"] = 100
        self.cv_pb["value"] = 0
        
        s = Label(f2, text="Est finish at:")
        s.grid(row=15, column=1)
        cv_timetext = str(time.asctime(time.localtime(time.time())))
        self.timer = Label(self.f2, text=cv_timetext)
        self.timer.grid(row=15, column=2)
        
        self.cv_canvas = FigureCanvasTkAgg(self.cv_f, master=f2)
        self.cv_canvas.get_tk_widget().grid(row=13, column=0, columnspan=10)
        self.cv_a.set_title("CV")
        self.cv_a.set_xlabel("Voltage")
        self.cv_a.set_ylabel("Capacitance")
        self.cv_canvas.draw()
        
        s = Button(f2, text="Start CV", command=self.cv_prepare_values)
        s.grid(row=3, column=7)
        
        s = Button(f2, text="Stop", command=self.quit)
        s.grid(row=4, column=7)
        
        print "finished drawing"
        
    def update(self):
        while self.outputdata.qsize():
            try:
                (data, percent, timeremain) = self.outputdata.get(0)
                
                if self.type is 0:
                    print "Percent done:" +str(percent)
                    self.pb["value"] = percent
                    self.pb.update()
                    (voltages, currents) = data
                    negative = False
                    for v in voltages:
                        if v < 0:
                            negative = True
                    if negative:
                        line,= self.a.plot(map(lambda x: x*-1.0, voltages), map(lambda x: x*-1.0, currents))
                    else:
                        line, = self.a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_color('r')
                    self.a.set_title("IV")
                    self.a.set_xlabel("Voltage [V]")
                    self.a.set_ylabel("Current [A]")
                    self.canvas.draw()

                    timetext = str(time.asctime(time.localtime(time.time()+timeremain)))
                    self.timer = Label(self.f1, text=timetext)
                    self.timer.grid(row=12, column=2)
                    
                    
                elif self.type is 1:
                    (voltages, caps) = data
                    print "Percent done:" +str(percent)
                    self.cv_pb["value"] = percent
                    self.cv_pb.update()
                    #print "Caps:+++++++"
                    #print caps
                    #print "============="
                    colors = {0:'b', 1:'g', 2:'r', 3:'c', 4:'m', 5:'k'}
                    i = 0
                    for c in caps:
                        """
                        print "VOLTS++++++"
                        print voltages
                        print "ENDVOLTS===="
                        #(a, b) = c[0]
                        print "CAPSENSE+++++"
                        print c
                        
                        print "ENDCAP======="
                        """
                        
                        if self.first:
                            line, = self.cv_a.plot(voltages, c, label=(self.cv_frequencies.get().split(",")[i]+"Hz"))
                            self.cv_a.legend()
                        else:
                            line, = self.cv_a.plot(voltages, c)
                        line.set_antialiased(True)
                        line.set_color(colors.get(i))   
                        i += 1
                        self.cv_a.set_title("CV")
                        self.cv_a.set_xlabel("Voltage [V]")
                        self.cv_a.set_ylabel("Capacitance [F]")
                        self.cv_canvas.draw()
                        
                    timetext = str(time.asctime(time.localtime(time.time()+timeremain)))
                    self.timer = Label(self.f2, text=timetext)
                    self.timer.grid(row=15, column=2)
                    
                self.first = False
            except Queue.Empty:
                pass
                
    def quit(self):
        print "placing order"
        self.stop.put("random")
    
    def prepare_values(self):
        print "preparing iv values"
        input_params = ((self.compliance.get(), self.compliance_scale.get(), self.start_volt.get(), self.end_volt.get(), self.step_volt.get(), self.hold_time.get(), self.source_choice.get(), self.recipients.get(), self.filename.get()), 0)
        self.inputdata.put(input_params)   
        self.f.clf()        
        self.a = self.f.add_subplot(111)
        self.type = 0
        
    def cv_prepare_values(self):
        print "preparing cv values"
        self.first = True
        input_params = ((self.cv_compliance.get(), self.cv_compliance_scale.get(), self.cv_start_volt.get(), self.cv_end_volt.get(), self.cv_step_volt.get(), self.cv_hold_time.get(), self.cv_source_choice.get(),
                         map(lambda x: x.strip(), self.cv_frequencies.get().split(",")), self.cv_function_choice.get(), self.cv_amplitude.get(), self.cv_impedance.get(), self.cv_integration.get(), self.cv_recipients.get()
                         , self.cv_filename.get()), 1)
        print input_params
        self.inputdata.put(input_params)  
        self.cv_f.clf()
        self.cv_a = self.cv_f.add_subplot(111)
        self.type = 1
    
def getvalues(input_params, dataout):
    if "Windows" in platform.platform():
            (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, filename) = input_params
    else:
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, thowaway) = input_params
        filename = tkFileDialog.asksaveasfilename(initialdir = "/",title = "Save data",filetypes = (("Microsoft Excel file","*.xlsx"),("all files","*.*")))+".xlsx"
    print "File done"
    
    try:
        comp = float(float(compliance)*({'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale, 1e-6)))
        source_params = (int(float(start_volt)), int(float(end_volt)), (float(step_volt)),
                             float(hold_time), comp)
    except ValueError:
        print "Please fill in all fields!"
    data = ()
    if source_params is None:
        pass
    else:
        data = GetIV(source_params, {"Keithley 2675a":1, "Keithley 2400":0}.get(source_choice, 0), dataout)
            
    data_out = xlsxwriter.Workbook(filename)
    path = filename
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
        mails = recipients.split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())
    
        print sentTo
        sendMail(path, sentTo)
    except:
        pass
    
def cv_getvalues(input_params, dataout):
    print input_params
    if "Windows" in platform.platform():
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, frequencies, function, amplitude, impedance, integration, recipients, filename) = input_params
    else:
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, frequencies, function, amplitude, impedance, integration, recipients, thowaway) = input_params
        filename = tkFileDialog.asksaveasfilename(initialdir = "/",title = "Save data",filetypes = (("Microsoft Excel file","*.xlsx"),("all files","*.*")))
    
    try:
        comp = float(float(compliance)*({'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale, 1e-6)))
        params = (int(float(start_volt)), int(float(end_volt)), int(float(step_volt)),
                             float(hold_time), comp, frequencies, float(amplitude), function, int(impedance), {"Short":0, "Medium":1, "Long":2}.get(integration))
        print params
    except ValueError:
        print "Please fill in all fields!"
    data = ()
    if params is None:
        pass
    else:
        data = GetCV(params, {"Keithley 2657a":1, "Keithley 2400":0}.get(source_choice), dataout)
    
    data_out = xlsxwriter.Workbook(filename)
    path = filename+str(time.asctime(time.localtime(time.time())))+".xlsx"
    worksheet = data_out.add_worksheet()
    
    (v, i, c, r) = data
    
    row=9
    col=0
    
    chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
    worksheet.write(8, 0, "V")
    for volt in v:
        worksheet.write(row, col, volt)
        row+=1
    
    col +=1
    last_col = col
    for f in frequencies:
        worksheet.write(7, col, "Freq="+f+"Hz")
        col +=3
        
    col = last_col
    row=9
    for frequency in i:
        worksheet.write(8, col, "I")
        row = 9
        for current in frequency:
            worksheet.write(row, col, current)
            row+=1
        col+=3
    
    col = last_col+1
    last_col = col
    for frequency in c:
        worksheet.write(8, col, "C")
        row=9
        for cap in frequency:
            worksheet.write(row, col, cap)
            row+=1
        col+=3
    
    col = last_col+1
    last_col = col
    
    fs=0
    for frequency in r:
        fs +=1
        worksheet.write(8, col, "R")
        row=9
        for res in frequency:
            worksheet.write(row, col, res)
            row+=1
        col+=3
    row +=5
    if fs >= 1:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$B$10:$B$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('D'+str(row), chart)

        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$C$10:$C$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [F]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('D'+str(row+20), chart)
        
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$D$10:$D$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('D'+str(row+40), chart)
        
    if fs >= 2:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$E$10:$E$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('L'+str(row), chart)
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$F$10:$F$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [C]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('L'+str(row+20), chart)
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$G$10:$G$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('L'+str(row+40), chart)
        
    if fs >= 3:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$H$10:$H$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('T'+str(row), chart)
        
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$I$10:$I$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [C]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('T'+str(row+20), chart)
    
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$J$10:$J$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('T'+str(row+40), chart)
    
    if fs >= 4:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$K$10:$K$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('AB'+str(row), chart)
    
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$L$10:$L$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [C]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('AB'+str(row+20), chart)
    
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$M$10:$M$'+str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('AB'+str(row+40), chart)
    
    data_out.close()
    
    try:
        mails = recipients.split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())
                
        print sentTo
        sendMail(path, sentTo)
    except:
        print "Failed to get recipients"
        pass
    
    
    
class ThreadedProgram:
    
    def __init__(self, master):
        self.master = master
        self.inputdata = Queue.Queue()
        self.outputdata = Queue.Queue()
        self.stopqueue = Queue.Queue()
        print "making gui"
        
        self.running = 1
        self.gui = GuiPart(master, self.inputdata, self.outputdata, self.stopqueue)
        
        self.thread1 = threading.Thread(target=self.workerThread1)
        self.thread1.start()
        self.periodicCall()
        self.measuring = False
    
    def periodicCall(self):
        #print "Period"
        self.gui.update()
        if not self.stopqueue.empty():
            print "quitting"
            import sys
            self.master.destroy()
            self.running = 0
            sys.exit(0)
            
        self.master.after(200, self.periodicCall)
    
    def workerThread1(self):
        while self.running:
            #print "looping"
            if self.inputdata.empty() is False and self.measuring is False:
                self.measuring= True
                print "doing stuff"
                #print self.inputdata
                (params, type) = self.inputdata.get()
                if type is 0:
                    getvalues(params, self.outputdata)
                elif type is 1:
                    cv_getvalues(params, self.outputdata)
                else:
                    pass
                    #spa_getvalues(params, self.outputdata)
                self.measuring=False
    
    def endapp(self):
        self.running = 0
    
class test:
    
    def current_monitoring(self, source_params, sourcemeter, dataout):
        
        (voltage_point, step_volt, hold_time, compliance, hours, minutes, seconds) = source_params
        
        currents = []
        timestamps = []
        voltages = []
        
        keithley = 0
        
        total_time = seconds+60*minutes+3600*hours
        start_time = time.time()
        
        if test:
            pass
        else:
            if sourcemeter is 0:
                keithley = Keithley2400()
            else:
                keithley = Keithley2657a()
            keithley.configure_measurement(1, 0, compliance)
        
        last_volt = 0
        badCount = 0
        
        scaled = False
        
        if step_volt < 1.0:
            start_volt *=1000
            voltage_point *=1000
            step_volt*=1000
            scaled = True
        
        if 0>voltage_point:
            step_volt = -1*step_volt
            
        for volt in xrange(0, voltage_point, step_volt):
            
            curr = 0
            if test:
                pass
            else:
                if scaled:
                    keithley.set_output(volt/1000.0)
                else:
                    keithley.set_output(volt)
                
            time.sleep(hold_time)
            
            if test:
                curr = (volt+randint(0, 10))*1e-9
            else:
                curr = keithley.get_current()
            #curr = volt
            
            if abs(curr)>abs(compliance-50e-9):
                badCount = badCount + 1        
            else:
                badCount = 0    
            
            if badCount>=5 :
                print "Compliance reached"
                break
            
            if scaled:
                last_volt = volt/1000.0
            else:
                last_volt = volt
            
            dataout.put(((timestamps, currents), 0, total_time))
            print """ramping up"""
            
        print "current time"
        print time.time()
        print "Start time"
        print start_time
        print "total time"
        print total_time
        
        start_time = time.time()
        while(time.time()<start_time+total_time):
            time.sleep(20)
            
            dataout.put(((timestamps, currents), 0, total_time))
            currents.append(randint(0, 10)*1e-9)
            timestamps.append(time.time()-start_time)
            print "timestamprs"
            print timestamps
            print "currents"
            print currents  
        print "Finished"
    
if __name__=="__main__":
    """
    x = test()
    data = Queue.Queue()
    print time.time()
    x.current_monitoring((10, 1, 1, 1, 0, 2, 0), 0, data)
    
    params = (0, -20, 2, 0.5, 0.1, 0)

    dataout = Queue.Queue()
    spa_iv(params, dataout)
    """
    root = Tk()
    root.geometry('800x800')
    root.title('Adap')
    client = ThreadedProgram(root)
    root.mainloop() 
    """
    daq = Agilent4156()
    daq.configure_vmu(discharge=True, _vmu=1, _mode = 0, name="VMU1")
    daq.configure_measurement()
    daq.configure_sampling_measurement()
    daq.configure_sampling_stop()
    print daq.inst.query(":PAGE:DISP:LIST?")
    daq.inst.write(":PAGE:DISP:LIST \'@TIME\', \'I1\', \'VMU1\';")
    
    daq.measurement_actions()
    daq.wait_for_acquisition()
    
    print daq.read_trace_data("I1")
    print "VMU+++++"
    print daq.read_trace_data("VMU1")
    print "VMU====="
    """
