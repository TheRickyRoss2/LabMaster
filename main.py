from Keithley import Keithley_2400
import visa
rm = visa.ResourceManager()
print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")



class Agilent_E4980A(object):
    def init(self, gpib):
        return 50
    def set_voltage(self):
        return 74

if __name__=="__main__":
    keithley = Keithley_2400()
    print "Waiting for input"