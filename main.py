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

test=False

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
            curr = (volt+randint(0, 10))
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
        voltages.append(volt)
        """
        #TODO Live graphics
        Y = currents
        graph.set_ydata(Y)
        graph.draw()
        """
        last_volt = volt
        dataout.put(((voltages, currents), 100*abs((volt+step_volt)/float(end_volt))))
        
        
    while abs(last_volt)>5:
        if test is True:
            pass
        else:
            keithley.set_output(last_volt)
        
        time.sleep(0.5)
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

def GetCV(params, sourcemeter, dataout):
    
    capacitance = []
    voltages = []
    parameter2 = []
    p2=  []
    c = []
    currents = []
    
    keithley = 0
    
    if test is True:
        pass
    else:
        if sourcemeter is 0:
            keithley = Keithley2400()
        else:
            keithley = Keithley2657a()
        keithley.init()
        keithley.configure_measurement()
        
    last_volt = 0
    
    (start_volt, end_volt, step_volt, delay_time, compliance,
     frequencies, level, function, impedance, int_time) = params
    print "CV PARAMS---------"
    print params
    print "CV PARAMS END______"
    if test is True:
        pass
    else:
        agilent = AgilentE4980a()
        agilent.init()
        agilent.configure_measurement(function)
        agilent.configure_aperture(int_time)
    badCount = 0
    
    if start_volt>end_volt:
        step_volt = -1*step_volt
    
    for volt in xrange(start_volt, end_volt, step_volt):
        
        if test is True:
            pass
        else:
            keithley.set_output(volt)
            
        curr = 0
        for f in frequencies:
            time.sleep(delay_time)

            if test is True:
                capacitance.append((volt*float(f)))
            else:
                agilent.configure_measurement_signal(float(f), 0, level)
                (data, aux) = agilent.read_data()
                capacitance.append(data)
                p2.append(aux)
                
                if test is True:
                    curr = volt
                else:
                    curr= keithley.get_current()
                c.append(curr)
        
            
        if abs(curr)>abs(compliance-50e-9):
            badCount = badCount + 1        
        else:
            badCount = 0    
        
        if badCount>=5 :
            print "Compliance reached"
            break
        
        voltages.append(volt)
        formatted_cap = []
        for i in xrange(0,len(frequencies), 1):
            print "Frequency: " +str(frequencies[i]) 
            print capacitance[i::len(frequencies)]
            formatted_cap.append(capacitance[i::len(frequencies)])
            parameter2.append(p2[i::len(frequencies)])
            currents.append(c[i::len(frequencies)])
        dataout.put(((voltages, formatted_cap), 100*abs((volt+step_volt)/float(end_volt))))
        
        last_volt = volt
        #graph point here
        
    while last_volt > 0:
        
        if last_volt<=5:
            keithley.set_output(0)
            last_volt = 0
        else:
            keithley.set_output(last_volt-5)
            last_volt -= 5
            
        time.sleep(0.5)
    
    if test is True:
        pass
    else:
        keithley.enable_output(False)
    
    
    return (voltages, formatted_cap, parameter2)

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
        
        self.f = plt.figure(figsize=(4, 4), dpi=75)
        self.a = self.f.add_subplot(111)
        
        self.cv_f = plt.figure(figsize=(4, 4), dpi=75)
        self.cv_a = self.cv_f.add_subplot(111)

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
        
        self.recipients.set("adapbot@gmail.com")
        s = Label(f1, text="Email data to:")
        s.pack(side=LEFT)
        s.grid(row=5, column=1)
        
        s = Entry(f1, textvariable = self.recipients)
        s.pack(side=LEFT)
        s.grid(row=5, column=2)
    
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
        self.a.set_title("IV")
        
        self.a.set_xlabel("Voltage")
        self.a.set_ylabel("Current")

        #plt.xlabel("Voltage")
        #plt.ylabel("Current")
        #plt.title("IV")
        self.canvas.draw()
        
        s = Button(f1, text="Start IV", command=self.prepare_values)
        s.pack(side=RIGHT)
        s.grid(row=3, column=7)
        
        s = Button(f1, text="Stop", command=endcommand)
        s.pack(side=RIGHT)
        s.grid(row=4, column=7)
        
        """
        /***********************************************************
         * CV GUI
         **********************************************************/
        """
        
        self.cv_start_volt.set("0.0")
        s = Label(f2, text="Start Volt")
        s.pack(side=LEFT)
        s.grid(row=0, column=1)
        s = Entry(f2, textvariable = self.cv_start_volt)
        s.pack(side=LEFT)
        s.grid(row=0, column=2)       
        s = Label(f2, text="V")
        s.pack(side=LEFT)
        s.grid(row=0, column=3)
        
        self.cv_end_volt.set("40.0")
        s = Label(f2, text="End Volt")
        s.pack(side=LEFT)
        s.grid(row=1, column=1)
        s = Entry(f2, textvariable = self.cv_end_volt)
        s.pack(side=LEFT)
        s.grid(row=1, column=2)
        s = Label(f2, text="V")
        s.pack(side=LEFT)
        s.grid(row=1, column=3)
        
        self.cv_step_volt.set("1.0")
        s = Label(f2, text="Step Volt")
        s.pack(side=LEFT)
        s.grid(row=2, column=1)
        s = Entry(f2, textvariable = self.cv_step_volt)
        s.pack(side=LEFT)
        s.grid(row=2, column=2)
        s = Label(f2, text="V")
        s.pack(side=LEFT)
        s.grid(row=2, column=3)
        
        self.cv_hold_time.set("1.0")
        s = Label(f2, text="Hold Time")
        s.pack(side=LEFT)
        s.grid(row=3, column=1)
        s = Entry(f2, textvariable = self.cv_hold_time)
        s.pack(side=LEFT)
        s.grid(row=3, column=2)
        s = Label(f2, text="s")
        s.pack(side=LEFT)
        s.grid(row=3, column=3)
        
        self.cv_compliance.set("1.0")
        s = Label(f2, text="Compliance")
        s.pack(side=LEFT)
        s.grid(row=4, column=1)
        s = Entry(f2, textvariable = self.cv_compliance)
        s.pack(side=LEFT)
        s.grid(row=4, column=2)
        self.cv_compliance_scale.set('uA')
        s = OptionMenu(f2, self.cv_compliance_scale, *compliance_choices)
        s.pack(side=LEFT)
        s.grid(row=4, column=3)
        
        self.cv_recipients.set("adapbot@gmail.com")
        s = Label(f2, text="Email data to:")
        s.pack(side=LEFT)
        s.grid(row=5, column=1)
        s = Entry(f2, textvariable = self.cv_recipients)
        s.pack(side=LEFT)
        s.grid(row=5, column=2)
        
        s = Label(f2, text="Agilent LCRMeter Parameters", relief=RAISED)
        s.pack(side=LEFT)
        s.grid(row=6, column=1, columnspan=2)
        
        self.cv_impedance = StringVar()
        s = Label(f2, text="Function")
        s.pack(side=LEFT)
        s.grid(row=7, column=1)
        function_choices = {"CPD", "CPQ",  "CPG",   "CPRP",  "CSD",  "CSQ", "CSRS",   "LPD",
                 "LPQ", "LPG", "LPRP", "LPRD", "LSD", "LSQ", "LSRS", "LSRD",
                 "RX", "ZTD", "ZTR", "GB",   "YTD", "YTR", "VDID"}
        self.cv_function_choice = StringVar()
        self.cv_function_choice.set('CPD')
        s = OptionMenu(f2, self.cv_function_choice, *function_choices)
        s.pack(side=LEFT)
        s.grid(row=7, column=2)
        
        self.cv_impedance.set("2000")
        s = Label(f2, text="Impedance")
        s.pack(side=LEFT)
        s.grid(row=8, column=1)
        s = Entry(f2, textvariable=self.cv_impedance )
        s.pack(side=LEFT)
        s.grid(row=8, column=2)
        s = Label(f2, text="Ω")
        s.pack(side=LEFT)
        s.grid(row=8, column=3)
        
        self.cv_frequencies.set("100, 200, 1000, 2000")
        s = Label(f2, text="Frequencies")
        s.pack(side=LEFT)
        s.grid(row=9, column=1)
        s = Entry(f2, textvariable=self.cv_frequencies )
        s.pack(side=LEFT)
        s.grid(row=9, column=2)
        s = Label(f2, text="Hz")
        s.pack(side=LEFT)
        s.grid(row=9, column=3)
        
        self.cv_amplitude.set("5.0")
        s = Label(f2, text="Signal Amplitude")
        s.pack(side=LEFT)
        s.grid(row=10, column=1)
        s = Entry(f2, textvariable=self.cv_amplitude )
        s.pack(side=LEFT)
        s.grid(row=10, column=2)
        s = Label(f2, text="V")
        s.pack(side=LEFT)
        s.grid(row=10, column=3)
        
        cv_int_choices = {"Short", "Medium", "Long"}
        s = Label(f2, text="Integration time")
        s.pack(side=LEFT)
        s.grid(row=11, column=1)
        self.cv_integration.set("Short")
        s = OptionMenu(f2, self.cv_integration, *cv_int_choices)
        s.pack(side=LEFT)
        s.grid(row=11, column=2)
    
        self.cv_source_choice.set('Keithley 2657a')
        s = OptionMenu(f2, self.cv_source_choice, *source_choices)
        s.pack(side=LEFT)
        s.grid(row=0, column=7)
        
        s = Label(f2, text="Progress:")
        s.pack(side=LEFT)
        s.grid(row=13, column=1)
        
        self.cv_pb = ttk.Progressbar(f2, orient="horizontal", length=200, mode="determinate")
        self.cv_pb.pack(side=LEFT)
        self.cv_pb.grid(row=13, column= 2, columnspan=5)
        self.cv_pb["maximum"] = 100
        self.cv_pb["value"] = 0
        
        self.cv_canvas = FigureCanvasTkAgg(self.cv_f, master=f2)
        self.cv_canvas.get_tk_widget().grid(row=12, column=0, columnspan=10)
        self.cv_a.set_title("CV")
        self.cv_a.set_xlabel("Voltage")
        self.cv_a.set_ylabel("Capacitance")
        self.cv_canvas.draw()
        
        s = Button(f2, text="Start CV", command=self.cv_prepare_values)
        s.pack(side=RIGHT)
        s.grid(row=3, column=7)
        
        s = Button(f2, text="Stop", command=endcommand)
        s.pack(side=RIGHT)
        s.grid(row=4, column=7)
        
        print "finished drawing"
        
        
    def update(self):
        while self.outputdata.qsize():
            try:
                (data, percent) = self.outputdata.get(0)
                
                if self.type is 0:
                    print "Percent done:" +str(percent)
                    self.pb["value"] = percent
                    self.pb.update()
                    (voltages, currents) = data
                    line, = self.a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_linestyle('solid')
                    line.set_color('r')
                    self.canvas.draw
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
                        line, = self.cv_a.plot(voltages, c)
                        line.set_antialiased(True)
                        line.set_color(colors.get(i))
                        i += 1
                        self.cv_canvas.draw()
                
            except Queue.Empty:
                pass
                
    def quit(self):
        import sys
        sys.exit(1)
        #root.destroy()
    
    def prepare_values(self):
        print "preparing iv values"
        input_params = ((self.compliance.get(), self.compliance_scale.get(), self.start_volt.get(), self.end_volt.get(), self.step_volt.get(), self.hold_time.get(), self.source_choice.get()), 0)
        self.inputdata.put(input_params)   
        self.f.clf()        
        self.a = self.f.add_subplot(111)
        self.type = 0
        
        
    def cv_prepare_values(self):
        print "preparing cv values"
        input_params = ((self.cv_compliance.get(), self.cv_compliance_scale.get(), self.cv_start_volt.get(), self.cv_end_volt.get(), self.cv_step_volt.get(), self.cv_hold_time.get(), self.cv_source_choice.get(),
                         map(lambda x: x.strip(), self.cv_frequencies.get().split(",")), self.cv_function_choice.get(), self.cv_amplitude.get(), self.cv_impedance.get(), self.cv_integration.get()
                         ), 1)
        print input_params
        self.inputdata.put(input_params)  
        self.cv_f.clf()
        self.cv_a = self.cv_f.add_subplot(111)
        self.type = 1
        
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
        data = GetIV(source_params, {"Keithley 2675a":1, "Keithley 2400":0}.get(source_choice, 0), dataout)
            
    data_out = xlsxwriter.Workbook(filename+".xlsx")
    path = filename+".xlsx"
    worksheet = data_out.add_worksheet()
    (v, i, c, r) = data
    values = []
    
    print "V+++++"
    print v
    print "V=====\n"
    print "I+++++"
    print i
    print "I=====\n"
    print "C+++++"
    print c
    print "C=====\n"
    print "R+++++"
    print r
    print "R=====\n"
    
    for x in xrange(0,len(v), 1):
        for f in xrange(0, len(i), 1):
            values.append((v[x], i[f][x]))
    
    row=9
    col=0
    chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
    
    for volt in values:
        worksheet.write(row, col, volt)
        row+=1
    col +=1
    last_col = col
    row=9
    for frequency in i:
        row = 9
        for current in frequency:
            worksheet.write(row, col, current)
            row+=1
        col+=3
        
    col = last_col+1
    last_col = col
    for frequency in c:
        row=9
        for cap in frequency:
            worksheet.write(row, col, cap)
            row+=1
        col+=3
        
    col = last_col+1
    last_col = col
    for frequency in r:
        row=9
        for res in frequency:
            worksheet.write(row, col, res)
            row+=1
        col+=3
        
    
    chart.add_series({'categories': '=Sheet1!$A$10:$A$'+str(row), 'values': '=Sheet1!$B$10:$B$'+str(row)})
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

def cv_getvalues(input_params, dataout):
    print input_params
    filename = tkFileDialog.asksaveasfilename(initialdir = "/",title = "Save data",filetypes = (("Microsoft Excel file","*.xlsx"),("all files","*.*")))
    (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, frequencies, function, amplitude, impedance, integration) = input_params
    
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
    
    data_out = xlsxwriter.Workbook(filename+".xlsx")
    path = filename+".xlsx"
    worksheet = data_out.add_worksheet()
    (v, i) = data
    values = []
    print "\nDATA++++++"
    print data
    print "DATA======\n"
    for x in xrange(0,len(v), 1):
        print "V++++++++"
        print v[x]
        print "V--------\n"
        print "C++++++++"
        print i[x]
        print "C========\n"
        values.append((v[x], i[x]))
    row=0
    col=0
    chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
    
    for volt, cur in values:
        worksheet.write(row, col, volt)
        worksheet.write(row, col+1, cur)
        row+=1
    
    chart.add_series({'categories': '=Sheet1!$A$1:$A$'+str(row), 'values': '=Sheet1!$B$1:$B$'+str(row), 'marker': {'type': 'star'}})
    chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
    chart.set_y_axis({'name':'Capacitance [F]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
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
                #print self.inputdata
                (params, type) = self.inputdata.get()
                if type is 0:
                    getvalues(params, self.outputdata)
                elif type is 1:
                    cv_getvalues(params, self.outputdata)
                else:
                    spa_getvalues(params, self.outputdata)
                self.measuring=False
        
    def endapp(self):
        self.running = 0
        
if __name__=="__main__":
    
    root = Tk()
    root.geometry('610x800')
    root.title('Adap')
    client = ThreadedProgram(root)
    root.mainloop() 
    
    """
    
    agilent = Agilent4156()
    agilent.configure_vmu(discharge=True, _vmu=2, _mode = 1, name="GOOD")
    agilent.read_trace_data("VM1")
    """