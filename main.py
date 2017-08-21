from Keithley import Keithley2400
from Agilent import AgilentE4980a

import visa

rm = visa.ResourceManager()
print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")

if __name__=="__main__":
    keithley = Keithley2400()
    agilent = AgilentE4980a()
    print str(float(2.5))
