#import visa

#rm = visa.ResourceManager()
#print(rm.list_resources())
#inst = rm.open_resource(rm.list_resources()[0])
#print(inst.query("REMOTE 716"))
#print(inst.query("CLEAR 7"))
#x = raw_input(">")

class Keithley_2400(object):
    def init(self, gpib):
        return 1
    def set_voltage(self, vout):
        return 100
    def configure_sweep(self, start_volt, end_volt, step_volt, delay_time, hold_time):
        return 90
    
class Agilent_E4980A(object):
    def init(self, gpib):
        return 50
    def set_voltage(self):
        return 74

if __name__=="__main__":
    device = Keithley_2400()
    print(device.init(24))
    print(device.set_voltage(1000))
    