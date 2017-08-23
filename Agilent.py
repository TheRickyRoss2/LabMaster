import visa
import time
from itertools import chain
from numpy import short

class Agilent4156(object):
    
    def init(self, gpib=6):
        """Set up gpib controller for device"""
        
        assert(gpib >= 0), "Please enter a valid GPIB address"
        self.gpib_addr = gpib
        
        print "Initializing Agilent semiconductor parameter analyzer"
        rm = visa.ResourceManager()
        self.inst = rm.open_resource(rm.list_resources()[0])
        for x in rm.list_resources():
            if str(self.gpib_addr) in x:
                print "found"
                self.inst = rm.open_resource(x)
                
        print self.inst.query("*IDN?")
        
        self.inst.write("*RST")
        self.inst.write("*ESE 60;*SRE 48;*CLS;")
        self.inst.timeout= 10000
        
    def configure_measurement(self, _mode=1):
        mode = {0:":PAGE:CHAN:MODE SWE;", 1:":PAGE:CHAN:MODE SAMP", 2:":PAGE:CHAN:MODE QSCV"}.get(
             _mode,":PAGE:CHAN:MODE SAMP")
         
        self.inst.write(mode)
        
    def configure_sampling_measurement(self, _mode = 0, _filter = False, auto_time=True, 
                                       hold_time=0, interval=4e-3, total_time=0.4,no_samples=101):
        mode = {0:"LIN", 1:"L10", 2:"L25", 3:"L50", 4:"THIN"}.get(_mode, "LIN")
        self.inst.write(":PAGE:MEAS:SAMP:MODE " +mode+";")
        self.inst.write(":PAGE:MEAS:SAMP:HTIM "+str(hold_time)+";")
        self.inst.write(":PAGE:MEAS:SAMP:IINT "+str(interval)+";")
        if _filter is True:
            self.inst.write(":PAGE:MEAS:SAMP:FILT ON;")
        else:
            self.inst.write(":PAGE:MEAS:SAMP:FILT OFF;")
            
        self.inst.write(":PAGE:MEAS:SAMP:PER "+str(total_time)+";")
        if auto_time is True:
            self.inst.write(":PAGE:MEAS:SAMP:PER:AUTO ON;")
        else:
            self.inst.write(":PAGE:MEAS:SAMP:PER:AUTO OFF;")
            
        self.inst.write(":PAGE:MEAS:SAMP:POIN "+str(no_samples)+";")
        
        
    def configure_sampling_stop(self, stop_condition = False, no_events=1,
                                 _event_type=0, delay=0, thresh=0, var ="V2"):
        
        event_type = {0:"LOW", 1:"HIGH", 2:"ABSL", 3:"ABSH"}.get(_event_type, "LOW")
        if stop_condition is True:
            self.inst.write(":PAGE:MEAS:SAMP:SCON ON;")
        else:
            self.inst.write(":PAGE:MEAS:SAMP:SCON OFF;")
            
        self.inst.write(":PAGE:MEAS:SAMP:SCON:ECO "+str(no_events)+";")
        self.inst.write(":PAGE:MEAS:SAMP:SCON:EDEL "+str(delay)+";")
        self.inst.write(":PAGE:MEAS:SAMP:SCON:EVEN "+event_type+";")
        self.inst.write(":PAGE:MEAS:SAMP:SCON:NAME \'"+var+"\';")
        self.inst.write(":PAGE:MEAS:SAMP:SCON:THR "+str(thresh)+";")
        
    def measurement_actions(self, _action=2):
        
        action = {0:"APP", 1:"REP", 2:"SING", 3:"STOP"}.get(_action,"SING")
        self.inst.write(":PAGE:SCON:"+action+";")
        
    def wait_for_acquisition(self):
        return self.inst.query("*OPC?")
        
    def read_trace_data(self, var="I1"):
        return map(lambda x: float(x), self.inst.query(":FORM:BORD NORM;DATA ASC;:DATA? \'"+var+"\';").split(","))
        
    def configure_channel(self, _chan=0, _func=3, _mode = 4, sres=0, standby=False):
        chan = {0:"SMU1", 1:"SMU2", 2:"SMU3", 3:"SMU4"}.get(_chan, "SMU1")
        func = {0:"VAR1", 1:"VAR2", 2:"VARD", 3:"CONS"}.get(_func, "CONS")
        mode = {0:"V", 1:"I", 2:"VPUL", 3:"IPUL", 4:"COMM"}.get(_mode, "COMM")
        self.inst.write(":PAGE:CHAN:"+ chan +":FUNC: "+ func+";")
        iname = {0:"I1", 1:"I2", 2:"I3", 3:"I4"}.get(_chan)
        vname = {0:"V1", 1:"V2", 2:"V3", 3:"V4"}.get(_chan)
        self.inst.write(":PAGE:CHAN:"+ chan + ":INAM\'" + iname + "\';")
        self.inst.write(":PAGE:CHAN:"+ chan + ":MODE "+mode+";")
        self.inst.write(":PAGE:CHAN:"+ chan + ":SRES "+str(sres)+";")
        if standby is True:
            self.inst.write(":PAGE:CHAN:"+ chan + ":STAN ON;")
        else:
            self.inst.write(":PAGE:CHAN:"+ chan + ":STAN OFF;")


        self.inst.write(":PAGE:CHAN:"+ chan + ":VNAM\'" + vname + "\';")
        
    def configure_integration_time(self, NPLC=16, _int_time=1, short_time=640e-6):
        int_time = {0:"SHOR", 1:"MED", 2:"LONG"}.get(_int_time, "MED")
        self.inst.write(":PAGE:MEAS:MSET:ITIME:LONG "+str(NPLC)+";")
        self.inst.write(":PAGE:MEAS:MSET:ITIME "+int_time+";")
        self.inst.write(":PAGE:MEAS:MSET:ITIME:SHOR "+str(short_time)+";")
        
        
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
        