import visa

class Keithley_2400(object):
    
    def init(self, gpib):
        assert(gpib >= 0), "Please enter a valid gpib address"
        self.gpib_addr = gpib
        
        print "Initializing keithley 2400"
        rm = visa.ResourceManager()
        self.inst = rm.open_resource(rm.list_resources()[0])
        for x in rm.list_resources():
            if str(self.gpib_addr) in x:
                print "found"
                self.inst = rm.open_resource(x)
                
        try:
            print(self.inst.query("*IDN?"))
        finally:
            print "Unable to open visa resource"
        self.inst.query("*RST")
        self.inst.query("*ESE 1;*SRE 32;*CLS;:FUNC:CONC ON;:FUNC:ALL;:TRAC:FEED:CONT NEV;:RES:MODE MAN;")
        print self.inst.query(":SYST:ERR?")
        
        
    def measure_function(self, x):
        return{
            0:"VOLT",
            1:"CURR",
            2:"RES"
        }.get(x, "CURR")
        
    def configure_measurement(self, funct=0):
        
        assert(funct>=0 and funct <3), "Invalid function specified"
        
        
        
        
    def set_voltage(self, vout):
        return 100
    def configure_sweep(self, start_volt, end_volt, step_volt, delay_time, hold_time):
        return 90
    