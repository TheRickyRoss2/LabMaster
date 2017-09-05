from Keithley import Keithley2400, Keithley2657a
from Agilent import AgilentE4980a, Agilent4156
from emailbot import sendMail
from Tkinter import Tk, Label, Button, StringVar, Entry, OptionMenu
import ttk
from Tkconstants import LEFT, RIGHT
import matplotlib
matplotlib.use("TkAgg")

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import time
import visa
import tkFileDialog
import xlsxwriter
from multiprocessing.pool import ThreadPool
import Queue



rm = visa.ResourceManager()
print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")

f = Figure(figsize=(5,5), dpi=100)
a = f.add_subplot(111)

def GetIV(sourceparam, sourcemeter=1):
    (start_volt, end_volt, step_volt, delay_time, compliance) = sourceparam
    
    currents = []
    voltages = []
    keithley = 0
    if sourcemeter is 0:
        keithley = Keithley2400()
    else:
        keithley = Keithley2657a()
    keithley.init()
    keithley.configure_measurement(1, 0, compliance)
    last_volt = 0
    badCount = 0
    
    if start_volt>end_volt:
        step_volt = -1*step_volt
    
    for volt in xrange(start_volt, end_volt, step_volt):
        print "in loop"
        
        keithley.set_output(volt)
        time.sleep(delay_time)
        
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
        
        a.plot(voltages, currents)
        
        pb["value"] = volt/(end_volt-step_volt)
        
        #graph point here
        
    while abs(last_volt)>5:
        pass
        #keithley.set_output(last_volt)
        time.sleep(delay_time/2.0)
        if last_volt < 0:
            last_volt +=5
        else:
            last_volt -=5
        
    #keithley.set_output(0)
    #keithley.enable_output(False)
    return (voltages, currents)

def GetCV(voltmeter, lcrmeter, sourcemeter = 0):
    capacitance = []
    keithley = 0
    if sourcemeter is 0:
        keithley = Keithley2400()
    else:
        keithley = Keithley2657a()
    keithley.init()
    keithley.configure_measurement()
    last_volt = 0
    (frequencies, meas_time, avg_factor, signal_type, level, function, impedance) = lcrmeter
    
    (start_volt, end_volt, step_volt, hold_time, delay_time, compliance) = voltmeter
    
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


def getvalues():
    
    filename = tkFileDialog.asksaveasfilename(initialdir = "/",title = "Save data",filetypes = (("Microsoft Excel file","*.xlsx"),("all files","*.*")))
    source_params = ()
    try:
        comp = float(compliance.get())*{'mA':1e-3, 'uA':1e-6, 'nA':1e-9}.get(compliance_scale.get())
        
        source_params = (int(float(start_volt.get())), int(float(end_volt.get())), int(float(step_volt.get())),
                         float(hold_time.get()), comp)
        print source_params
    except ValueError:
        print "Please fill in all fields!"
    data = ()
    if source_params is None:
        pass
    else:
        pool = ThreadPool(processes = 1)
        async_result = pool.apply(GetIV, (source_params, {'Keithley 2400':0, 'Keithley2657a':1}.get(source_choice.get())))     
        data = async_result.get()
        
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
def quit():
    root.destroy()

if __name__=="__main__":
    """
    keithley = Keithley2400()
    keithley._init(22)
    keithley.configure_measurement(1)
    keithley.enable_output(True)
    keithley.configure_source(0, 0, 0.1)

    for i in xrange(0, 50):
        keithley.out_source(i)
        print keithley.read_single_point()
        time.sleep(0.5)
    
    agilent = AgilentE4980a()
    agilent.init(19)
    
    agilent.configure_measurement(0, 4, True)
    for i in (1000, 2000, 10000, 20000):
        agilent.configure_measurement_signal(i, 0, 5)
        print agilent.read_data()
        time.sleep(0.5)
    
    GetIV(0, 20, 1, 0.5, 0.1, 0.1)
    voltmeter = (0, 20, 1, 0.5, 0.1, 0.1)
    frequencies = (1000, 2000, 10000, 20000, 50000, 100000)
    lcrmeter = (frequencies, 1, 1, 0, 1, 0, 1)
    GetCV(voltmeter, lcrmeter)
    print str(float(2.5))
    
    voltmeter = (0, 20, 1, 0.5, 0.1, 0.1)
    X = np.linspace(0, 20, 21)
    plt.ion()
    Y = X*0
    graph = plt.plot(X,Y)[0]
    iv = GetIV(voltmeter, graph)
    plt.plot(iv)
    
    keithley = Keithley2657a()
    keithley.init()
    
    keithley.configure_measurement()
    keithley.output_level()
    keithley.output_limit()
    keithley.enable_output(True)
    
    current = []
    
    for x in xrange(0, 20, 1):
        time.sleep(0.5)
        keithley.output_level(x)
        current.append(keithley.get_current())
    for x in xrange(20, 0, -2):
        time.sleep(0.25)
        keithley.output_level(x)
        
    keithley.output_level(0)
    keithley.enable_output()
    
    print current
    
    agilent = Agilent4156()
    agilent.init()
    agilent.configure_measurement()
    agilent.configure_sampling_measurement()
    agilent.configure_sampling_stop()
    agilent.measurement_actions()
    agilent.wait_for_acquisition()
    print agilent.read_trace_data()
    """
    
    #spa_iv(0, 0)
   
    #sendMail("sample.xlsx", ['rirrodri@ucsc.edu', 'therickyross2@gmail.com'])
    
    #voltmeter = (start_volt, end_volt, step_volt, delay, compliance)
    #voltmeter = (0, -20, 1, 0.1, 1e-6)
    #currents= GetIV(voltmeter, 1)
    
    """
    data_out = xlsxwriter.Workbook('iv_curve.xlsx')
    worksheet = data_out.add_worksheet()
    v = []
    for i in xrange(start_volt, end_volt, step_volt):
        v.append(i)
    values = []
    for x in xrange(0, len(v), 1):
        values.append((v[x], currents[x]))
    row=0
    col=0
    
    for volt, cur in values:
        worksheet.write(row, col, volt)
        worksheet.write(row, col+1, cur)
        row+=1
    data_out.close()
    """
    #for x in xrange(0, 160, 1):
        #print str(x)+","+str(currents[x])

    #spa_iv(0, 0)

    root = Tk()
    root.geometry('600x750')
    root.title('Adap')
    
    start_volt = StringVar()
    end_volt = StringVar()
    step_volt = StringVar()
    hold_time = StringVar()
    compliance = StringVar() 
    recipients = StringVar()   
    
    start_volt.set("0.0")
    end_volt.set("100.0")
    step_volt.set("5.0")
    hold_time.set("1.0")
    compliance.set("1.0")
    
    n = ttk.Notebook(root)
    n.grid(row=1, column=0, columnspan=60, rowspan=60, sticky='NESW')
    f1 = ttk.Frame(n)
    f2 = ttk.Frame(n)
    f3 = ttk.Frame(n)
    f4 = ttk.Frame(n)
    n.add(f1, text='IV')
    n.add(f2, text='CV')
    n.add(f3, text='SPA IV')
    n.add(f4, text='Bot')
    

    
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
    
    s = Label(f1, text="Email data to:")
    s.pack(side=LEFT)
    s.grid(row=5, column=1)
    
    s = Entry(f1, textvariable = recipients)
    s.pack(side=LEFT)
    s.grid(row=5, column=2)
    
    s = Label(f1, text="Separate with ','")
    s.pack(side=LEFT)
    s.grid(row=5, column=3)

    source_choices = {'Keithley 2400', 'Keithley 2657a'}
    source_choice = StringVar()
    source_choice.set('Keithley 2657a')
    s = OptionMenu(f1, source_choice, *source_choices)
    s.pack(side=LEFT)
    s.grid(row=0, column=7)
    
    s = Label(f1, text="Percent Complete:")
    s.pack(side=LEFT)
    s.grid(row=11, column=1)
    
    pb = ttk.Progressbar(f1, orient="horizontal", length=200, mode="determinate")
    pb.pack(side=LEFT)
    pb.grid(row=11, column= 2, columnspan=5)
    pb["maximum"] = 100
    pb["value"] = 0
    
    
    canvas = FigureCanvasTkAgg(f, master=f1)
    canvas.get_tk_widget().grid(row=10, columnspan=10)
    canvas.draw()
    
    s = Button(f1, text="Start IV", command=getvalues)
    s.pack(side=RIGHT)
    s.grid(row=3, column=7)
    
    s = Button(f1, text="Stop", command=quit)
    s.pack(side=RIGHT)
    s.grid(row=4, column=7)
    
    root.mainloop()    
