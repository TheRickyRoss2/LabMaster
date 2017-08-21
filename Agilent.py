import visa
import time

class AgilentE4980a(object):
    
    def init(self, gpib=19):
        """Set up gpib controller for device"""
        
        assert(gpib >= 0), "Please enter a valid gpib address"
        self.gpib_addr = gpib
        
        print "Initializing agilent lcr_meter"
        rm = visa.ResourceManager()
        self.inst = rm.open_resource(rm.list_resources()[0])
        for x in rm.list_resources():
            if str(self.gpib_addr) in x:
                print "found"
                self.inst = rm.open_resource(x)
                
        print self.inst.query("*IDN?;")
        
        self.inst.write("*RST;")
        self.inst.write("*ESE 60;*SRE 48;*CLS;")
        self.inst.timeout= 10000
        
    def configure_measurement(self, _function=0, _impedance=3, autorange=True):
        
        function = {0:"CPD", 1:"CPQ",  2:"CPG",   3:"CPRP",  4:"CSD",  5:"CSQ", 6:"CSRS",   7:"LPD",
                 8:"LPQ", 9:"LPG", 10:"LPRP", 11:"LPRD", 12:"LSD", 13:"LSQ", 14:"LSRS", 15:"LSRD",
                 16:"RX", 17:"ZTD", 18:"ZTR", 19:"GB",   20:"YTD", 21:"YTR", 22:"VDID"
                 }.get(_function, "CPD")
        
        impedance = {0:"1E-1;",1:"1E+0;",2:"1E+1;", 3:"1E+2;",4:"3E+2;",5:"1E+3;",6:"3E+3;",7:"1E+4",
                     8:"3E+4", 9:"1E+5"}.get(_impedance, "1E+2")
        
        if autorange is True:             
            self.inst.write(":FUNC:IMP "+function+";:FUNC:IMP:RANG:AUTO ON")
        else:
            self.inst.write(":FUNC:IMP "+function+";:FUNC:IMP:RANG: "+impedance)
        
    def configure_measurement_signal(self, frequency=10000, _signal_type=0, signal_level=5.0):
        signal_type = {0:"VOLT", 1:"CURR"}.get(_signal_type, "VOLT")
        self.inst.write(":FREQ "+str(float(frequency))+";:"+signal_type+" "+str(float(signal_level)))
        
    def configure_aperture(self, _meas_time=1, avg_factor=1):
        meas_time = {0:"SHOR", 1:"MED", 2:"LONG"}.get(_meas_time, "MED")
        self.inst.write(":APER "+meas_time+","+str(float(avg_factor))+";")
    
    def initiate(self):
        self.inst.write(":INIT;")
    
    def fetch_data(self):
        _data_out = self.inst.query(":FETC?")
        #print _data_out
        data_out = _data_out
        parameter1 = data_out.split(",")[0]
        parameter2 = data_out.split(",")[1]
        results = (float(parameter1), float(parameter2))
        #print results
        return results
    
    def read_data(self):
        self.initiate()
        return self.fetch_data()
        
        
        
        
        
        