from Keithley import Keithley2400
import visa
rm = visa.ResourceManager()
print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")

if __name__=="__main__":
    keithley = Keithley2400()
    print str(float(2.5))