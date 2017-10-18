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

debug = False 
rm = visa.ResourceManager()
print(rm.list_resources())
# inst = rm.open_resource(rm.list_resources()[0])
# print(inst.query("REMOTE 716"))
# print(inst.query("CLEAR 7"))
# x = raw_input(">")

def GetIV(sourceparam, sourcemeter, dataout):
    (start_volt, end_volt, step_volt, delay_time, compliance) = sourceparam
    
    currents = []
    voltages = []
    keithley = 0
    
    if debug:
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
        start_volt *= 1000
        end_volt *= 1000
        step_volt *= 1000
        scaled = True
    
    if start_volt > end_volt:
        step_volt = -1 * step_volt
    
    print "looping now"
    
    for volt in xrange(start_volt, end_volt, int(step_volt)):
        start_time = time.time()
        
        curr = 0
        if debug:
            pass
        else:
            if scaled:
                keithley.set_output(volt / 1000.0)
            else:
                keithley.set_output(volt)
        
        time.sleep(delay_time)
        
        if debug:
            curr = (volt + randint(0, 10)) * 1e-9
        else:
            curr = keithley.get_current()
            
        print "Current Reading: "+str(curr)
        
        # curr = volt
        time.sleep(1)
        if abs(curr) > abs(compliance - 50e-9):
            badCount = badCount + 1        
        else:
            badCount = 0    
        
        if badCount >= 5 :
            print "Compliance reached"
            dataout.put(((voltages, currents), 100, 0))
            break
        
        currents.append(curr)
        if scaled:
            voltages.append(volt / 1000.0)
        else:
            voltages.append(volt)

        if scaled:
            last_volt = volt / 1000.0
        else:
            last_volt = volt
        
        time_remain = (time.time() - start_time) * (abs((end_volt - volt) / step_volt))
        
        dataout.put(((voltages, currents), 100 * abs((volt + step_volt) / float(end_volt)), time_remain))
        
        
    while abs(last_volt) > 25:
        if debug:
            pass
        else:
            keithley.set_output(last_volt)
        
        time.sleep(delay_time/2.0)
        
        if last_volt < 0:
            last_volt += abs(step_volt*2.0)
        else:
            last_volt -= abs(step_volt*2.0)
    
    time.sleep(delay_time/2.0)
    if debug:
        pass
    else:
        keithley.set_output(0)
        keithley.enable_output(False)
    return (voltages, currents)

def GetCV(params, sourcemeter, dataout):
    
    capacitance = []
    voltages = []
    p2 = []
    c = []
    keithley = 0
    
    if debug:
        pass
    else:
        if sourcemeter is 0:
            keithley = Keithley2400()
        else:
            keithley = Keithley2657a()
    
    last_volt = 0
    
    (start_volt, end_volt, step_volt, delay_time, compliance,
     frequencies, level, function, impedance, int_time) = params
    if debug:
        pass
    else:
        keithley.configure_measurement(1, 0, compliance)
    
    if debug:
        pass
    else:
        agilent = AgilentE4980a()
        agilent.init()
        agilent.configure_measurement(function)
        agilent.configure_aperture(int_time)
    badCount = 0
    
    scaled = False
    
    if step_volt < 1.0:
        start_volt *= 1000
        end_volt *= 1000
        step_volt *= 1000
        scaled = True
    
    if start_volt > end_volt:
        step_volt = -1 * step_volt
    
    start_time = time.time()
    for volt in xrange(start_volt, end_volt, step_volt):
    
        start_time = time.time()
        if debug:
            pass
        else:
            if scaled:
                keithley.set_output(volt / 1000.0)
            else:
                keithley.set_output(volt)
            
        curr = 0
        for f in frequencies:
            time.sleep(delay_time)

            if debug:
                capacitance.append((volt + int(f) * randint(0, 10)))
                curr = volt * 1e-10
                c.append(curr)
                p2.append(volt * 10)
            else:
                agilent.configure_measurement_signal(float(f), 0, level)
                (data, aux) = agilent.read_data()
                capacitance.append(data)
                p2.append(aux)
                curr = keithley.get_current()
                c.append(curr)
            
        if abs(curr) > abs(compliance - 50e-9):
            badCount = badCount + 1        
        else:
            badCount = 0    
        
        if badCount >= 5 :
            print "Compliance reached"
            break
        
        time_remain = (time.time() - start_time) * (abs((end_volt - volt) / step_volt))
        
        if scaled:
            voltages.append(volt / 1000.0)
        else:
            voltages.append(volt)
        formatted_cap = []
        parameter2 = []
        currents = []
        for i in xrange(0, len(frequencies), 1):
            formatted_cap.append(capacitance[i::len(frequencies)])
            parameter2.append(p2[i::len(frequencies)])
            currents.append(c[i::len(frequencies)])
        dataout.put(((voltages, formatted_cap), 100 * abs((volt + step_volt) / float(end_volt)), time_remain))
        
        time_remain = time.time() + (time.time() - start_time) * (abs((volt - end_volt) / end_volt))
        
        last_volt = volt
        # graph point here

    if debug:
        pass
    else:
        while abs(last_volt) > abs(step_volt):
            if last_volt <= step_volt:
                keithley.set_output(0)
                last_volt = 0
            else:
                keithley.set_output(last_volt - step_volt)
                last_volt -= step_volt

            time.sleep(1)

    if debug:
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
    daq.configure_integration_time(_int_time=int_time)
    
    scaled = False
    if step_volt < 1.0:
        start_volt *= 1000
        end_volt *= 1000
        step_volt *= 1000
        scaled = True
    
    if start_volt > end_volt:
        step_volt = -1 * step_volt
    
    for i in xrange(0, 4, 1):
        daq.configure_channel(i)  
    daq.configure_vmu()    
    
    last_volt = 0
    for volt in xrange(start_volt, end_volt, step_volt):
        
        if debug:
            pass
        else:
            if scaled:
                voltage_source.set_output(volt / 1000.0)
            else:
                voltage_source.set_output(volt)
        time.sleep(hold_time)
        
        daq.configure_measurement()
        daq.configure_sampling_measurement()
        daq.configure_sampling_stop()
        
        # daq.inst.write(":PAGE:DISP:GRAP:Y2:NAME \'I2\';")
        daq.inst.write(":PAGE:DISP:LIST \'@TIME\', \'I1\', \'I2\', \'I3\', \'I4\', \'VMU1\'")
        daq.measurement_actions()
        daq.wait_for_acquisition()
        
        current_smu1.append(daq.read_trace_data("I1"))
        current_smu2.append(daq.read_trace_data("I2"))
        
        # daq.inst.write(":PAGE:DISP:LIST \'@TIME\', \'I2\', \'I3\'")

        current_smu3.append(daq.read_trace_data("I3"))
        current_smu4.append(daq.read_trace_data("I4"))
        voltage_vmu1.append(daq.read_trace_data("VMU1"))
        current_smua.append(voltage_source.get_current())
        
        if scaled:
            voltage_smua.append(volt / 1000.0)
            last_volt = volt / 1000.0
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
    while abs(last_volt) >= 4:
        time.sleep(0.5)

        if debug:
            pass
        else:
            voltage_source.set_output(last_volt)
        
        if last_volt < 0:
            last_volt += 5
        else:
            last_volt -= 5
        
    time.sleep(0.5)
    voltage_source.set_output(0)
    voltage_source.enable_output(False)
    return(voltage_smua, current_smua, current_smu1, current_smu2, current_smu3, current_smu4, voltage_vmu1)

# TODO: current monitor bugfixes and fifo implementation
def curmon(source_params, sourcemeter, dataout):
        
    (voltage_point, step_volt, hold_time, compliance, minutes) = source_params
    print "(voltage_point, step_volt, hold_time, compliance, minutes)"
    print source_params
    currents = []
    timestamps = []
    voltages = []
    
    total_time = minutes*60
    
    keithley=0
    if debug:
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
      
    if step_volt < 1:
        voltage_point *= 1000
        step_volt *= 1000
        scaled = True
    else:
        step_volt = int(step_volt)
        
    if 0 > voltage_point:
        step_volt = -1 * step_volt
    
    start_time = time.time()
    
    for volt in xrange(0, voltage_point, step_volt):
            
        curr = 0
        if debug:
            pass
        else:
            if scaled:
                keithley.set_output(volt / 1000.0)
            else:
                keithley.set_output(volt)
                
        time.sleep(hold_time)
        
        if debug:
            curr = (volt + randint(0, 10)) * 1e-9
        else:
            curr = keithley.get_current()
        # curr = volt
        
        if abs(curr) > abs(compliance - 50e-9):
            badCount = badCount + 1        
        else:
            badCount = 0    
            
        if badCount >= 5 :
            print "Compliance reached"
            break
        
        if scaled:
            last_volt = volt / 1000.0
        else:
            last_volt = volt
        
        dataout.put(((timestamps, currents), 0, total_time+start_time))
        print """ramping up"""
    
    print "current time"
    print time.time()
    print "Start time"
    print start_time
    print "total time"
    print total_time
    
    start_time = time.time()
    while(time.time() < start_time + total_time):
        time.sleep(5)
        
        dataout.put(((timestamps, currents), 100*((time.time()-start_time)/total_time), start_time+total_time))
        if debug:
            currents.append(randint(0, 10) * 1e-9)
        else:
            currents.append(keithley.get_current())
        timestamps.append(time.time() - start_time)
        print "timestamps"
        print timestamps
        print "currents"
        print currents  
    print "Finished"
    
    while abs(last_volt) > 5:
        if debug:
            pass
        else:
            keithley.set_output(last_volt)
        
        time.sleep(hold_time/2.0)
        if last_volt < 0:
            last_volt += 5
        else:
            last_volt -= 5
    
    time.sleep(hold_time/2.0)
    if debug:
        pass
    else:
        keithley.set_output(0)
        keithley.enable_output(False)
    
    return (timestamps, currents)
    
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
        
        self.multiv_start_volt = StringVar()
        self.multiv_end_volt = StringVar()
        self.multiv_step_volt = StringVar()
        self.multiv_hold_time = StringVar()
        self.multiv_compliance = StringVar() 
        self.multiv_recipients = StringVar()   
        self.multiv_compliance_scale = StringVar()
        self.multiv_source_choice = StringVar()
        self.multiv_filename = StringVar()
        self.multiv_times = StringVar()
        
        self.curmon_start_volt = StringVar()
        self.curmon_end_volt = StringVar()
        self.curmon_step_volt = StringVar()
        self.curmon_hold_time = StringVar()
        self.curmon_compliance = StringVar() 
        self.curmon_recipients = StringVar()   
        self.curmon_compliance_scale = StringVar()
        self.curmon_source_choice = StringVar()
        self.curmon_filename = StringVar()
        self.curmon_time = StringVar()
        
        """
        IV GUI
        """
        ""
        self.start_volt.set("0.0")
        self.end_volt.set("100.0")
        self.step_volt.set("5.0")
        self.hold_time.set("1.0")
        self.compliance.set("1.0")
        
        self.f = plt.figure(figsize=(6, 4), dpi=60)
        self.a = self.f.add_subplot(111)
        
        self.cv_f = plt.figure(figsize=(6, 4), dpi=60)
        self.cv_a = self.cv_f.add_subplot(111)
        
        self.multiv_f = plt.figure(figsize=(6, 4), dpi=60)
        self.multiv_a = self.multiv_f.add_subplot(111)
        
        self.curmon_f = plt.figure(figsize=(6, 4), dpi=60)
        self.curmon_a = self.curmon_f.add_subplot(111)
        
        n = ttk.Notebook(root, width=800)
        n.grid(row=0, column=0, columnspan=100, rowspan=100, sticky='NESW')
        self.f1 = ttk.Frame(n)
        self.f2 = ttk.Frame(n)
        self.f3 = ttk.Frame(n)
        self.f4 = ttk.Frame(n)
        self.f5 = ttk.Frame(n)
        n.add(self.f1, text='Basic IV')
        n.add(self.f2, text='CV')
        n.add(self.f3, text='Param Analyzer IV ')
        n.add(self.f4, text='Multiple IV')
        n.add(self.f5, text='Current Monitor')
        
        if "Windows" in platform.platform():
            self.filename.set("iv_data")
            s = Label(self.f1, text="File name:")
            s.grid(row=0, column=1)
            s = Entry(self.f1, textvariable=self.filename)
            s.grid(row=0, column=2)
        
        s = Label(self.f1, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(self.f1, textvariable=self.start_volt)
        s.grid(row=1, column=2)
        s = Label(self.f1, text="V")
        s.grid(row=1, column=3)
        
        s = Label(self.f1, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(self.f1, textvariable=self.end_volt)
        s.grid(row=2, column=2)
        s = Label(self.f1, text="V")
        s.grid(row=2, column=3)
        
        s = Label(self.f1, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(self.f1, textvariable=self.step_volt)
        s.grid(row=3, column=2)
        s = Label(self.f1, text="V")
        s.grid(row=3, column=3)
        
        s = Label(self.f1, text="Hold Time")
        s.grid(row=4, column=1)
        s = Entry(self.f1, textvariable=self.hold_time)
        s.grid(row=4, column=2)
        s = Label(self.f1, text="s")
        s.grid(row=4, column=3)
        
        s = Label(self.f1, text="Compliance")
        s.grid(row=5, column=1)
        s = Entry(self.f1, textvariable=self.compliance)
        s.grid(row=5, column=2)
        compliance_choices = {'mA', 'uA', 'nA'}
        self.compliance_scale.set('uA')
        s = OptionMenu(self.f1, self.compliance_scale, *compliance_choices)
        s.grid(row=5, column=3)
        
        self.recipients.set("adapbot@gmail.com")
        s = Label(self.f1, text="Email data to:")
        s.grid(row=6, column=1)
        s = Entry(self.f1, textvariable=self.recipients)
        s.grid(row=6, column=2)
    
        source_choices = {'Keithley 2400', 'Keithley 2657a'}
        self.source_choice.set('Keithley 2657a')
        s = OptionMenu(self.f1, self.source_choice, *source_choices)
        s.grid(row=0, column=7)
        
        s = Label(self.f1, text="Progress:")
        s.grid(row=11, column=1)
       
        s = Label(self.f1, text="Est finish at:")
        s.grid(row=12, column=1)
        
        timetext = str(time.asctime(time.localtime(time.time())))
        self.timer = Label(self.f1, text=timetext)
        self.timer.grid(row=12, column=2)
        
        self.pb = ttk.Progressbar(self.f1, orient="horizontal", length=200, mode="determinate")
        self.pb.grid(row=11, column=2, columnspan=5)
        self.pb["maximum"] = 100
        self.pb["value"] = 0
        
        self.canvas = FigureCanvasTkAgg(self.f, master=self.f1)
        self.canvas.get_tk_widget().grid(row=7, columnspan=10)
        self.a.set_title("IV")
        self.a.set_xlabel("Voltage")
        self.a.set_ylabel("Current")

        # plt.xlabel("Voltage")
        # plt.ylabel("Current")
        # plt.title("IV")
        self.canvas.draw()
        
        s = Button(self.f1, text="Start IV", command=self.prepare_values)
        s.grid(row=3, column=7)
        
        s = Button(self.f1, text="Stop", command=self.quit)
        s.grid(row=4, column=7)
        
        """
        /***********************************************************
         * CV GUI
         **********************************************************/
        """
        self.cv_filename = StringVar()
        self.cv_filename.set("cv_data")
        
        if "Windows" in platform.platform():
            s = Label(self.f2, text="File name")
            s.grid(row=0, column=1)
            s = Entry(self.f2, textvariable=self.cv_filename)
            s.grid(row=0, column=2)
        
        self.cv_start_volt.set("0.0")
        s = Label(self.f2, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(self.f2, textvariable=self.cv_start_volt)
        s.grid(row=1, column=2)       
        s = Label(self.f2, text="V")
        s.grid(row=1, column=3)
        
        self.cv_end_volt.set("40.0")
        s = Label(self.f2, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(self.f2, textvariable=self.cv_end_volt)
        s.grid(row=2, column=2)
        s = Label(self.f2, text="V")
        s.grid(row=2, column=3)
        
        self.cv_step_volt.set("1.0")
        s = Label(self.f2, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(self.f2, textvariable=self.cv_step_volt)
        s.grid(row=3, column=2)
        s = Label(self.f2, text="V")
        s.grid(row=3, column=3)
        
        self.cv_hold_time.set("1.0")
        s = Label(self.f2, text="Hold Time")
        s.grid(row=4, column=1)
        s = Entry(self.f2, textvariable=self.cv_hold_time)
        s.grid(row=4, column=2)
        s = Label(self.f2, text="s")
        s.grid(row=4, column=3)
        
        self.cv_compliance.set("1.0")
        s = Label(self.f2, text="Compliance")
        s.grid(row=5, column=1)
        s = Entry(self.f2, textvariable=self.cv_compliance)
        s.grid(row=5, column=2)
        self.cv_compliance_scale.set('uA')
        s = OptionMenu(self.f2, self.cv_compliance_scale, *compliance_choices)
        s.grid(row=5, column=3)
        
        self.cv_recipients.set("adapbot@gmail.com")
        s = Label(self.f2, text="Email data to:")
        s.grid(row=6, column=1)
        s = Entry(self.f2, textvariable=self.cv_recipients)
        s.grid(row=6, column=2)
        
        s = Label(self.f2, text="Agilent LCRMeter Parameters", relief=RAISED)
        s.grid(row=7, column=1, columnspan=2)
        
        self.cv_impedance = StringVar()
        s = Label(self.f2, text="Function")
        s.grid(row=8, column=1)
        function_choices = {"CPD", "CPQ", "CPG", "CPRP", "CSD", "CSQ", "CSRS", "LPD",
                 "LPQ", "LPG", "LPRP", "LPRD", "LSD", "LSQ", "LSRS", "LSRD",
                 "RX", "ZTD", "ZTR", "GB", "YTD", "YTR", "VDID"}
        self.cv_function_choice = StringVar()
        self.cv_function_choice.set('CPD')
        s = OptionMenu(self.f2, self.cv_function_choice, *function_choices)
        s.grid(row=8, column=2)
        
        self.cv_impedance.set("2000")
        s = Label(self.f2, text="Impedance")
        s.grid(row=9, column=1)
        s = Entry(self.f2, textvariable=self.cv_impedance)
        s.grid(row=9, column=2)
        s = Label(self.f2, text="Ω")
        s.grid(row=9, column=3) 
        
        self.cv_frequencies.set("100, 200, 1000, 2000")
        s = Label(self.f2, text="Frequencies")
        s.grid(row=10, column=1)
        s = Entry(self.f2, textvariable=self.cv_frequencies)
        s.grid(row=10, column=2)
        s = Label(self.f2, text="Hz")
        s.grid(row=10, column=3)
        
        self.cv_amplitude.set("5.0")
        s = Label(self.f2, text="Signal Amplitude")
        s.grid(row=11, column=1)
        s = Entry(self.f2, textvariable=self.cv_amplitude)
        s.grid(row=11, column=2)
        s = Label(self.f2, text="V")
        s.grid(row=11, column=3)
        
        cv_int_choices = {"Short", "Medium", "Long"}
        s = Label(self.f2, text="Integration time")
        s.grid(row=12, column=1)
        self.cv_integration.set("Short")
        s = OptionMenu(self.f2, self.cv_integration, *cv_int_choices)
        s.grid(row=12, column=2)
    
        self.cv_source_choice.set('Keithley 2657a')
        s = OptionMenu(self.f2, self.cv_source_choice, *source_choices)
        s.grid(row=0, column=7)
        
        s = Label(self.f2, text="Progress:")
        s.grid(row=14, column=1)
        
        self.cv_pb = ttk.Progressbar(self.f2, orient="horizontal", length=200, mode="determinate")
        self.cv_pb.grid(row=14, column=2, columnspan=5)
        self.cv_pb["maximum"] = 100
        self.cv_pb["value"] = 0
        
        s = Label(self.f2, text="Est finish at:")
        s.grid(row=15, column=1)
        cv_timetext = str(time.asctime(time.localtime(time.time())))
        self.timer = Label(self.f2, text=cv_timetext)
        self.timer.grid(row=15, column=2)
        
        self.cv_canvas = FigureCanvasTkAgg(self.cv_f, master=self.f2)
        self.cv_canvas.get_tk_widget().grid(row=13, column=0, columnspan=10)
        self.cv_a.set_title("CV")
        self.cv_a.set_xlabel("Voltage")
        self.cv_a.set_ylabel("Capacitance")
        self.cv_canvas.draw()
        
        s = Button(self.f2, text="Start CV", command=self.cv_prepare_values)
        s.grid(row=3, column=7)
        
        s = Button(self.f2, text="Stop", command=self.quit)
        s.grid(row=4, column=7)
        
        print "finished drawing"
        
        """
        Multiple IV GUI
        """
        
        if "Windows" in platform.platform():
            self.multiv_filename.set("iv_data")
            s = Label(self.f4, text="File name:")
            s.grid(row=0, column=1)
            s = Entry(self.f4, textvariable=self.multiv_filename)
            s.grid(row=0, column=2)
        
        s = Label(self.f4, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(self.f4, textvariable=self.multiv_start_volt)
        s.grid(row=1, column=2)
        s = Label(self.f4, text="V")
        s.grid(row=1, column=3)
        
        s = Label(self.f4, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(self.f4, textvariable=self.multiv_end_volt)
        s.grid(row=2, column=2)
        s = Label(self.f4, text="V")
        s.grid(row=2, column=3)
        
        s = Label(self.f4, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(self.f4, textvariable=self.multiv_step_volt)
        s.grid(row=3, column=2)
        s = Label(self.f4, text="V")
        s.grid(row=3, column=3)

        s = Label(self.f4, text="Repeat Times")
        s.grid(row=4, column=1)
        s = Entry(self.f4, textvariable=self.multiv_times)
        s.grid(row=4, column=2)
        
        s = Label(self.f4, text="Hold Time")
        s.grid(row=5, column=1)
        s = Entry(self.f4, textvariable=self.multiv_hold_time)
        s.grid(row=5, column=2)
        s = Label(self.f4, text="s")
        s.grid(row=5, column=3)
        
        s = Label(self.f4, text="Compliance")
        s.grid(row=6, column=1)
        s = Entry(self.f4, textvariable=self.multiv_compliance)
        s.grid(row=6, column=2)
        self.multiv_compliance_scale.set('uA')
        s = OptionMenu(self.f4, self.multiv_compliance_scale, *compliance_choices)
        s.grid(row=6, column=3)
        
        self.multiv_recipients.set("adapbot@gmail.com")
        s = Label(self.f4, text="Email data to:")
        s.grid(row=7, column=1)
        s = Entry(self.f4, textvariable=self.multiv_recipients)
        s.grid(row=7, column=2)
    
        source_choices = {'Keithley 2400', 'Keithley 2657a'}
        self.multiv_source_choice.set('Keithley 2657a')
        s = OptionMenu(self.f4, self.multiv_source_choice, *source_choices)
        s.grid(row=0, column=7)
        
        s = Label(self.f4, text="Progress:")
        s.grid(row=11, column=1)
       
        s = Label(self.f4, text="Est finish at:")
        s.grid(row=12, column=1)
        
        self.multiv_timer = Label(self.f4, text=timetext)
        self.multiv_timer.grid(row=12, column=2)
        
        self.multiv_pb = ttk.Progressbar(self.f4, orient="horizontal", length=200, mode="determinate")
        self.multiv_pb.grid(row=11, column=2, columnspan=5)
        self.multiv_pb["maximum"] = 100
        self.multiv_pb["value"] = 0
        
        self.multiv_canvas = FigureCanvasTkAgg(self.multiv_f, master=self.f4)
        self.multiv_canvas.get_tk_widget().grid(row=8, columnspan=10)
        self.multiv_a.set_title("IV")
        self.multiv_a.set_xlabel("Voltage")
        self.multiv_a.set_ylabel("Current")

        self.multiv_canvas.draw()
        
        s = Button(self.f4, text="Start IVs", command=self.multiv_prepare_values)
        s.grid(row=3, column=7)
        
        s = Button(self.f4, text="Stop", command=self.quit)
        s.grid(row=4, column=7)
        
        """
        Current Monitor IV
        """
        
        
        if "Windows" in platform.platform():
            self.curmon_filename.set("iv_data")
            s = Label(self.f5, text="File name:")
            s.grid(row=0, column=1)
            s = Entry(self.f5, textvariable=self.curmon_filename)
            s.grid(row=0, column=2)
        
        s = Label(self.f5, text="Start Volt")
        s.grid(row=1, column=1)
        s = Entry(self.f5, textvariable=self.curmon_start_volt)
        s.grid(row=1, column=2)
        s = Label(self.f5, text="V")
        s.grid(row=1, column=3)
        
        s = Label(self.f5, text="End Volt")
        s.grid(row=2, column=1)
        s = Entry(self.f5, textvariable=self.curmon_end_volt)
        s.grid(row=2, column=2)
        s = Label(self.f5, text="V")
        s.grid(row=2, column=3)
        
        s = Label(self.f5, text="Step Volt")
        s.grid(row=3, column=1)
        s = Entry(self.f5, textvariable=self.curmon_step_volt)
        s.grid(row=3, column=2)
        s = Label(self.f5, text="V")
        s.grid(row=3, column=3)

        s = Label(self.f5, text="Test Time")
        s.grid(row=4, column=1)
        s = Entry(self.f5, textvariable=self.curmon_time)
        s.grid(row=4, column=2)
        s = Label(self.f5, text="M")
        s.grid(row=4, column=3)
        
        s = Label(self.f5, text="Hold Time")
        s.grid(row=5, column=1)
        s = Entry(self.f5, textvariable=self.curmon_hold_time)
        s.grid(row=5, column=2)
        s = Label(self.f5, text="s")
        s.grid(row=5, column=3)
        
        s = Label(self.f5, text="Compliance")
        s.grid(row=6, column=1)
        s = Entry(self.f5, textvariable=self.curmon_compliance)
        s.grid(row=6, column=2)
        self.curmon_compliance_scale.set('uA')
        s = OptionMenu(self.f5, self.curmon_compliance_scale, *compliance_choices)
        s.grid(row=6, column=3)
        
        self.curmon_recipients.set("adapbot@gmail.com")
        s = Label(self.f5, text="Email data to:")
        s.grid(row=7, column=1)
        s = Entry(self.f5, textvariable=self.curmon_recipients)
        s.grid(row=7, column=2)
    
        source_choices = {'Keithley 2400', 'Keithley 2657a'}
        self.curmon_source_choice.set('Keithley 2657a')
        s = OptionMenu(self.f5, self.curmon_source_choice, *source_choices)
        s.grid(row=0, column=7)
        
        
        s = Label(self.f5, text="Progress:")
        s.grid(row=11, column=1)
       
        s = Label(self.f5, text="Est finish at:")
        s.grid(row=12, column=1)
        
        self.curmon_timer = Label(self.f5, text=timetext)
        self.curmon_timer.grid(row=12, column=2)
        
        self.curmon_pb = ttk.Progressbar(self.f5, orient="horizontal", length=200, mode="determinate")
        self.curmon_pb.grid(row=11, column=2, columnspan=5)
        self.curmon_pb["maximum"] = 100
        self.curmon_pb["value"] = 0
        
        self.curmon_canvas = FigureCanvasTkAgg(self.curmon_f, master=self.f5)
        self.curmon_canvas.get_tk_widget().grid(row=8, columnspan=10)
        self.curmon_a.set_title("IV")
        self.curmon_a.set_xlabel("Voltage")
        self.curmon_a.set_ylabel("Current")

        self.curmon_canvas.draw()
        
        s = Button(self.f5, text="Start CurMon", command=self.curmon_prepare_values)
        s.grid(row=3, column=7)
        
        s = Button(self.f5, text="Stop", command=self.quit)
        s.grid(row=4, column=7)
        
        
    def update(self):
        while self.outputdata.qsize():
            try:
                (data, percent, timeremain) = self.outputdata.get(0)
                
                if self.type is 0:
                    print "Percent done:" + str(percent)
                    self.pb["value"] = percent
                    self.pb.update()
                    (voltages, currents) = data
                    negative = False
                    for v in voltages:
                        if v < 0:
                            negative = True
                    if negative:
                        line, = self.a.plot(map(lambda x: x * -1.0, voltages), map(lambda x: x * -1.0, currents))
                    else:
                        line, = self.a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_color('r')
                    self.a.set_title("IV")
                    self.a.set_xlabel("Voltage [V]")
                    self.a.set_ylabel("Current [A]")
                    self.canvas.draw()

                    timetext = str(time.asctime(time.localtime(time.time() + timeremain)))
                    self.timer = Label(self.f1, text=timetext)
                    self.timer.grid(row=12, column=2)
                    
                    
                elif self.type is 1:
                    (voltages, caps) = data
                    print "Percent done:" + str(percent)
                    self.cv_pb["value"] = percent
                    self.cv_pb.update()
                    # print "Caps:+++++++"
                    # print caps
                    # print "============="
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
                            
                            line, = self.cv_a.plot(voltages, c, label=(self.cv_frequencies.get().split(",")[i] + "Hz"))
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
                        
                    timetext = str(time.asctime(time.localtime(time.time() + timeremain)))
                    self.timer = Label(self.f2, text=timetext)
                    self.timer.grid(row=15, column=2)
                    self.first=False
                    
                elif self.type is 2:
                    pass
                
                elif self.type is 3:
                    if self.first:
                        #self.multiv_f.clf()
                        pass
                    print "Percent done:" + str(percent)
                    self.multiv_pb["value"] = percent
                    self.multiv_pb.update()
                    (voltages, currents) = data
                    negative = False
                    for v in voltages:
                        if v < 0:
                            negative = True
                    if negative:
                        line, = self.multiv_a.plot(map(lambda x: x * -1.0, voltages), map(lambda x: x * -1.0, currents))
                    else:
                        line, = self.multiv_a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_color('r')
                    self.multiv_a.set_title("IV")
                    self.multiv_a.set_xlabel("Voltage [V]")
                    self.multiv_a.set_ylabel("Current [A]")
                    self.multiv_canvas.draw()

                    timetext = str(time.asctime(time.localtime(time.time() + timeremain)))
                    self.multiv_timer = Label(self.f4, text=timetext)
                    self.multiv_timer.grid(row=12, column=2)
                    
                elif self.type is 4:
                    
                    print "Percent done:" + str(percent)
                    self.curmon_pb["value"] = percent
                    self.curmon_pb.update()
                    (voltages, currents) = data
                    negative = False
                    for v in voltages:
                        if v < 0:
                            negative = True
                    if negative:
                        line, = self.curmon_a.plot(map(lambda x: x * -1.0, voltages), map(lambda x: x * -1.0, currents))
                    else:
                        line, = self.curmon_a.plot(voltages, currents)
                    line.set_antialiased(True)
                    line.set_color('r')
                    self.curmon_a.set_title("IV")
                    self.curmon_a.set_xlabel("Voltage [V]")
                    self.curmon_a.set_ylabel("Current [A]")
                    self.curmon_canvas.draw()

                    timetext = str(time.asctime(time.localtime(time.time() + timeremain)))
                    self.curmon_timer = Label(self.f5, text=timetext)
                    self.curmon_timer.grid(row=12, column=2)
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
        
    def multiv_prepare_values(self):
        
        print "preparing mult iv values"
        self.first = True
        input_params = ((self.multiv_compliance.get(), self.multiv_compliance_scale.get(), self.multiv_start_volt.get(), self.multiv_end_volt.get(), self.multiv_step_volt.get(), self.multiv_hold_time.get(), self.multiv_source_choice.get(), self.multiv_recipients.get(), self.multiv_filename.get(), self.multiv_times.get()), 3)
        self.inputdata.put(input_params)   
        self.multiv_f.clf()        
        self.multiv_a = self.multiv_f.add_subplot(111)
        self.type = 3
    
    def curmon_prepare_values(self):
        
        print "preparing current monitor values"
        self.first = True
        input_params = ((self.curmon_compliance.get(), self.curmon_compliance_scale.get(), self.curmon_start_volt.get(), self.curmon_end_volt.get(), self.curmon_step_volt.get(), self.curmon_hold_time.get(), self.curmon_source_choice.get(), self.curmon_recipients.get(), self.curmon_filename.get(), self.curmon_time.get()), 4)
        self.inputdata.put(input_params)   
        self.curmon_f.clf()        
        self.curmon_a = self.curmon_f.add_subplot(111)
        self.type = 4
        
def getvalues(input_params, dataout):
    if "Windows" in platform.platform():
            (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, filename) = input_params
    else:
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, thowaway) = input_params
        filename = tkFileDialog.asksaveasfilename(initialdir="~", title="Save data", filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")))
    print "File done"

    try:
        comp = float(float(compliance) * ({'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale, 1e-6)))
        source_params = (int(float(start_volt)), int(float(end_volt)), (float(step_volt)),
                             float(hold_time), comp)
    except ValueError:
        print "Please fill in all fields!"
    data = ()
    if source_params is None:
        pass
    else:
        print source_choice
        choice = 0
        if "2657a" in source_choice:
            print "asdf keithley 366"
            choice = 1
        data = GetIV(source_params, choice, dataout)
            
    fname = (((filename+"_"+str(time.asctime(time.localtime(time.time())))+".xlsx").replace(" ", "_")).replace(":", "_"))
    data_out = xlsxwriter.Workbook(fname)
    worksheet = data_out.add_worksheet()
    
    (v, i) = data
    values = []
    pos = v[0]>v[1]
    for x in xrange(0, len(v), 1):
        values.append((v[x], i[x]))
    row = 0
    col = 0
    
    chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
    
    for volt, cur in values:
        worksheet.write(row, col, volt)
        worksheet.write(row, col + 1, cur)
        row += 1
    
    chart.add_series({'categories': '=Sheet1!$A$1:$A$' + str(row), 'values': '=Sheet1!$B$1:$B$' + str(row)})
    chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}, 'reverse':pos})
    chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}, 'reverse':pos})
    chart.set_legend({'none':True})
    worksheet.insert_chart('D2', chart)
    data_out.close()
    
    try:
        mails = recipients.split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())
    
        print sentTo
        print fname
        sendMail(fname, sentTo)
    except:
        pass
    
def cv_getvalues(input_params, dataout):
    print input_params
    if "Windows" in platform.platform():
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, frequencies, function, amplitude, impedance, integration, recipients, filename) = input_params
    else:
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, frequencies, function, amplitude, impedance, integration, recipients, thowaway) = input_params
        filename = tkFileDialog.asksaveasfilename(initialdir="~", title="Save data", filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")))
    
    try:
        comp = float(float(compliance) * ({'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale, 1e-6)))
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
        fname = (((filename+str(time.asctime(time.localtime(time.time())))+".xlsx").replace(" ", "_")).replace(":", "_"))
    
    data_out = xlsxwriter.Workbook(fname)
    worksheet = data_out.add_worksheet()
    
    (v, i, c, r) = data
    
    row = 9
    col = 0
    
    chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
    worksheet.write(8, 0, "V")
    for volt in v:
        worksheet.write(row, col, volt)
        row += 1
    
    col += 1
    last_col = col
    for f in frequencies:
        worksheet.write(7, col, "Freq=" + f + "Hz")
        col += 3
        
    col = last_col
    row = 9
    for frequency in i:
        worksheet.write(8, col, "I")
        row = 9
        for current in frequency:
            worksheet.write(row, col, current)
            row += 1
        col += 3
    
    col = last_col + 1
    last_col = col
    for frequency in c:
        worksheet.write(8, col, "C")
        row = 9
        for cap in frequency:
            worksheet.write(row, col, cap)
            row += 1
        col += 3
    
    col = last_col + 1
    last_col = col
    
    fs = 0
    for frequency in r:
        fs += 1
        worksheet.write(8, col, "R")
        row = 9
        for res in frequency:
            worksheet.write(row, col, res)
            row += 1
        col += 3
    row += 5
    if fs >= 1:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$B$10:$B$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('D' + str(row), chart)

        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$C$10:$C$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [F]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('D' + str(row + 20), chart)
        
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$D$10:$D$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('D' + str(row + 40), chart)
        
    if fs >= 2:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$E$10:$E$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('L' + str(row), chart)
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$F$10:$F$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [C]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('L' + str(row + 20), chart)
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$G$10:$G$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('L' + str(row + 40), chart)
        
    if fs >= 3:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$H$10:$H$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('T' + str(row), chart)
        
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$I$10:$I$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [C]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('T' + str(row + 20), chart)
    
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$J$10:$J$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('T' + str(row + 40), chart)
    
    if fs >= 4:
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$K$10:$K$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('AB' + str(row), chart)
    
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$L$10:$L$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Capacitance [C]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('AB' + str(row + 20), chart)
    
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        chart.add_series({'categories': '=Sheet1!$A$10:$A$' + str(row), 'values': '=Sheet1!$M$10:$M$' + str(row), 'marker': {'type': 'star'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_y_axis({'name':'Resistance [R]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}})
        chart.set_legend({'none':True})
        worksheet.insert_chart('AB' + str(row + 40), chart)
    
    data_out.close()
    
    try:
        mails = recipients.split(",")
        sentTo = []
        for mailee in mails:
            sentTo.append(mailee.strip())
                
        print sentTo
        sendMail(fname, sentTo)
    except:
        print "Failed to get recipients"
        pass
    
# TODO: Implement value parsing from gui
def spa_getvalues(input_params, dataout):
    pass

def multiv_getvalues(input_params, dataout):
    if "Windows" in platform.platform():
            (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, filename, times_str) = input_params
    else:
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, thowaway, times_str) = input_params
        filename = tkFileDialog.asksaveasfilename(initialdir="~", title="Save data", filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")))
    print "File done"
    
    try:
        comp = float(float(compliance) * ({'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale, 1e-6)))
        source_params = (int(float(start_volt)), int(float(end_volt)), (float(step_volt)),
                             float(hold_time), comp)
        times = int(times_str)
        
    except ValueError:
        print "Please fill in all fields!"
    data = ()
    
    while times>0:

        if source_params is None:
            pass
        else:
            print source_choice
            choice = 0
            if "2657a" in source_choice:
                print "asdf keithley 366"
                choice = 1
            data = GetIV(source_params, choice, dataout)
        fname = (((filename+str(time.asctime(time.localtime(time.time())))+".xlsx").replace(" ", "_")).replace(":", "_"))
        print fname
        data_out = xlsxwriter.Workbook(fname)
        worksheet = data_out.add_worksheet()
        
        (v, i) = data
        values = []
        pos = v[0]>v[1]
        for x in xrange(0, len(v), 1):
            values.append((v[x], i[x]))
        row = 0
        col = 0
        chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
        for volt, cur in values:
            worksheet.write(row, col, volt)
            worksheet.write(row, col + 1, cur)
            row += 1
        chart.add_series({'categories': '=Sheet1!$A$1:$A$' + str(row), 'values': '=Sheet1!$B$1:$B$' + str(row), 'marker':{'type':'triangle'}})
        chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}, 'reverse':pos})
        chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}, 'reverse':pos})
        chart.set_legend({'none':True})
        worksheet.insert_chart('D2', chart)
        data_out.close()
        
        try:
            mails = recipients.split(",")
            sentTo = []
            for mailee in mails:
                sentTo.append(mailee.strip())
        
            print sentTo
            sendMail(fname, sentTo)
        except:
            pass
        data_out.close()
        times-=1
    
    
# TODO: Implement value parsing from gui
def curmon_getvalues(input_params, dataout):
    
    if "Windows" in platform.platform():
            (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, filename, total_time) = input_params
            filename = ((filename+"_"+str(time.asctime(time.localtime(time.time())))+".xlsx").replace(" ", "_")).replace(":","_")
    else:
        (compliance, compliance_scale, start_volt, end_volt, step_volt, hold_time, source_choice, recipients, thowaway, total_time) = input_params
        filename = tkFileDialog.asksaveasfilename(initialdir="~", title="Save data", filetypes=(("Microsoft Excel file", "*.xlsx"), ("all files", "*.*")))
    print "File done"
    
    try:
        comp = float(float(compliance) * ({'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale, 1e-6)))
        source_params = (int(float(end_volt)), float(step_volt),
                             float(hold_time), comp, int(total_time))
    except ValueError:
        print "Please fill in all fields!"
    data = ()
    if source_params is None:
        pass
    else:
        print source_choice
        choice = 0
        if "2657a" in source_choice:
            print "asdf keithley 366"
            choice = 1
        data = curmon(source_params, choice, dataout)
            
    data_out = xlsxwriter.Workbook(filename)
    path = filename
    worksheet = data_out.add_worksheet()
    
    (v, i) = data
    values = []
    pos = v[0]>v[1]
    for x in xrange(0, len(v), 1):
        values.append((v[x], i[x]))
    row = 0
    col = 0
    
    chart = data_out.add_chart({'type':'scatter', 'subtype':'straight_with_markers'})
    
    for volt, cur in values:
        worksheet.write(row, col, volt)
        worksheet.write(row, col + 1, cur)
        row += 1
    
    chart.add_series({'categories': '=Sheet1!$A$1:$A$' + str(row), 'values': '=Sheet1!$B$1:$B$' + str(row)})
    chart.set_x_axis({'name':'Voltage [V]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}, 'reverse':pos})
    chart.set_y_axis({'name':'Current [A]', 'major_gridlines':{'visible':True}, 'minor_tick_mark':'cross', 'major_tick_mark':'cross', 'line':{'color':'black'}, 'reverse':pos})
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
    
    
class ThreadedProgram:
    
    def __init__(self, master):
        self.master = master
        self.inputdata = Queue.Queue()
        self.outputdata = Queue.Queue()
        self.stopqueue = Queue.Queue()
        print "Generating GUI"
        
        self.running = 1
        self.gui = GuiPart(master, self.inputdata, self.outputdata, self.stopqueue)
        
        self.thread1 = threading.Thread(target=self.workerThread1)
        self.thread1.start()
        self.periodicCall()
        self.measuring = False
    
    def periodicCall(self):
        # print "Period"
        self.gui.update()
        if not self.stopqueue.empty():
            print "Exiting program"
            import sys
            self.master.destroy()
            self.running = 0
            sys.exit(0)
            
        self.master.after(200, self.periodicCall)
    
    def workerThread1(self):
        while self.running:
            if self.inputdata.empty() is False and self.measuring is False:
                self.measuring = True
                print "Instantiating Threads"
                (params, type) = self.inputdata.get()
                if type is 0:
                    getvalues(params, self.outputdata)
                elif type is 1:
                    cv_getvalues(params, self.outputdata)
                elif type is 2:
                    spa_getvalues(params, self.outputdata)
                elif type is 3:
                    multiv_getvalues(params, self.outputdata)
                elif type is 4:
                    curmon_getvalues(params, self.outputdata)
                else:
                    pass
                self.measuring = False
    def endapp(self):
        self.running = 0
    
    
if __name__ == "__main__":

    root = Tk()
    root.geometry('800x800')
    root.title('LabMaster')
    client = ThreadedProgram(root)
    root.mainloop() 
