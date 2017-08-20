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
            print(self.inst.query("*IDN?;"))
        finally:
            print "Unable to open visa resource"
        self.inst.query("*RST;")
        self.inst.query("*ESE 1;*SRE 32;*CLS;:FUNC:CONC ON;:FUNC:ALL;:TRAC:FEED:CONT NEV;:RES:MODE MAN;")
        print self.inst.query(":SYST:ERR?;")
        
        
    def get_function(self, x):
        return{
            0:":VOLT",
            1:":CURR",
            2:":RES"
        }.get(x, "CURR")
        
    def get_source_mode(self, x):
        return{
            0:"VOLT",
            1:"CURR"
        }.get(x, "VOLT")
        
    def configure_measurement(self, funct=0):
        
        assert(funct>=0 and funct <3), "Invalid function specified"
        self.inst.query(self.get_function(funct)+"RANG:AUTO ON;")
    
    def configure_source(self, source_mode, output_level, compliance):
        assert(source_mode>=0 and source_mode<2), "Invalid Source function"
        assert(output_level>-1100 and output_level < 1100), "Voltage out of range"
        assert(compliance<0.5), "Compliance out of range"
        func = self.get_source_mode(source_mode)
        self.inst.query("SOUR:FUNC "+func+";:SOUR:"+func+" "+str(float(output_level))+";:CURR:PROT "+str(float(compliance))+";")
        
    def enable_output(self, output):
        if output is True:
            self.inst.query("OUTP ON;")
            return
        self.inst.query("OUTP OFF;")
        
    def read_single_point(self, timeout = 10000):
        
        
    def set_voltage(self, vout):
        return 100
    def configure_sweep(self, start_volt, end_volt, step_volt, delay_time, hold_time):
        return 90
    