from Keithley import Keithley2400
from Agilent import AgilentE4980a
from Keithley import Keithley2657a
from Agilent import Agilent4156
import time
import visa
import numpy as np
import matplotlib.pyplot as plt
from Tkconstants import CURRENT

rm = visa.ResourceManager()
print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")

def GetIV(voltmeter, graph):
    (start_volt, end_volt, step_volt, delay_time, hold_time, compliance) = voltmeter
    
    currents = []
    keithley = Keithley2400()
    keithley.init()
    keithley.configure_measurement()
    keithley.configure_source()
    last_volt = 0
    badCount = 0
    if start_volt>end_volt:
        step_volt = -1*step_volt
    
    for volt in xrange(start_volt, end_volt+step_volt, step_volt):
        
        keithley.set_output(volt)
        time.sleep(delay_time)
        
        curr = keithley.read_single_point(hold_time)
        
        if(curr[1] > compliance):
            badCount = badCount + 1
            
        if(badCount>=5):
            print "Compliance reached"
            break
        
        currents.append(curr)
        
        Y = currents
        graph.set_ydata(Y)
        graph.draw()
        
        last_volt = volt
        #graph point here
        
    for volt in xrange(last_volt, start_volt, step_volt*-2):
        keithley.set_output(volt)
        time.sleep(delay_time/2.0)
    
    keithley.enable_output(False)
    print currents
    return currents

def GetCV(voltmeter, lcrmeter):
    capacitance = []
    keithley = Keithley2400()
    keithley.init()
    keithley.configure_measurement()
    keithley.configure_source()
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
    """
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
    """
    agilent = Agilent4156()
    agilent.init()
    agilent.configure_measurement()
    agilent.configure_sampling_measurement()
    agilent.configure_sampling_stop()
    agilent.measurement_actions()
    agilent.wait_for_acquisition()
    print agilent.read_trace_data()
    """