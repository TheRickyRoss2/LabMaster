from Keithley import Keithley2400
from Agilent import AgilentE4980a
import time
import visa

rm = visa.ResourceManager()
print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")

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
    """
    agilent = AgilentE4980a()
    agilent._init(19)
    
    agilent.configure_measurement(0, 4, True)
    for i in (1000, 2000, 10000, 20000):
        agilent.configure_measurement_signal(i, 0, 5)
        print agilent.read_data()
        time.sleep(0.5)
    
    print str(float(2.5))
