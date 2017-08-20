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
        
    def configure_multipoint(self, arm_count=1, trigger_count=1, mode=0):
        # :ARM:COUN 1;:TRIG:COUN 1;:SOUR:VOLT:MODE FIX;:SOUR:CURR:MODE FIX;
        assert(mode>=0 and mode <3), "Invalid mode"
        source_mode = {0:"FIX", 1:"SWE", 2:"LIST"}.get(mode)
        self.inst.query(":ARM:COUN "+str(arm_count)+";:TRIG:COUN "+str(trigger_count)+";:SOUR:VOLT:MODE "+source_mode+";:SOUR:CURR:MODE "+source_mode)
        
    def configure_trigger(self):
        self.inst.query("ARM:SOUR IMM;:ARM:TIM 0.010000;:TRIG:SOUR IMM;:TRIG:DEL 0.000000;")
        
    def read_single_point(self, timeout = 10000):
        return 1
        
    def set_voltage(self, vout):
        return 100
    def configure_sweep(self, start_volt, end_volt, step_volt, delay_time, hold_time):
        return 90
    