# -*- coding: utf-8 -*-

import pyvisa


class PowerSupplyFactory(object):
    """
    Abstract Class for creating power supply interfaces.
    Call the static method PowerSupply.factory("insert_my_type")
    to get a power supply object instance
    """

    def factory(power_supply_type):
        """
        Factory method for instantiating power supply objects
        Call this function to create a client interface for the
        power supplies
        :param gpib_address: Address of gpib device
        :param power_supply_type: String representation of power
        supply type
        Current supported types are as follows:
        - Keithley Sourcemeter 2657a, called with 'Keithley2657a'
        - Keithley High Voltage Supply 2400/2410, called with 'Keithley2400'
        :return: Power Supply Object which corresponds to type sepcified
        """

        if power_supply_type == "keithley2657a":
            return Keithley2657a()
        elif power_supply_type == "keithley2400":
            return Keithley2400()
        assert 0, "Could not create power supply of type: " + power_supply_type

    factory = staticmethod(factory)


class Keithley2657a(PowerSupplyFactory):

    def __init__(self, gpib_address=24):
        """
        Initializer for Keithley power supply
        Checks to make sure that the power supply is not on
        :param gpib_address: GPIB Address of power supply
        """

        assert (gpib_address >= 0), "Please enter a valid GPIB address"

        resource_manager = pyvisa.ResourceManager()

        # Temporarily set the power supply to point at the first address
        self.supply = resource_manager.\
            open_resource(resource_manager.list_resources()[0])

        # Search through intrument cluster for the gpib address
        # TODO Implement the same thing using a filter or reduce
        for address in resource_manager.list_resources():
            print("Searching for Keithley @" + str(gpib_address))
            if str(gpib_address) in str(address):
                print("Keithley Found")
                self.supply = resource_manager.open_resource(address)
            else:
                print("Keithley not found; Please check GPIB Address")

        print("Initializing Keithley 2657A")
        # Verify sanity of device
        print(self.supply.query("*IDN?"))

        # Reset state of device
        self.supply.write("setup.recall(1)")

        # Display current readings on front screen
        self.supply.write("display.smua.measure.func = display.MEASURE_DCAMPS")

        # Set the front end ADC to integrate for more accurate measurements
        self.supply.write("smua.measure.adc=smua.ADC_INTEGRATE")

        # Set the number of power line cycles for integration
        # The integration aperture is based on the number of power line cycles (NPLC),
        # where 1 PLC for 60 Hz is 16.67 ms (1/60) and 1 PLC for 50 Hz is 20 ms (1/50).
        self.supply.write("smua.measure.nplc=25")

        # Autorange for easy front panel readings
        self.supply.write("smua.measure.autorangei=smua.AUTORANGE_ON")

        # Reasonable timeout, really dont have to worry about this
        self.supply.timeout = 10000

    def __configure_source(self, _source=1):
        """
        Set what we are outputting, whether that is voltage or current
        :param _source: 0 for current source, 1 for voltage source
        :return: None
        """

        source = {0: "OUTPUT_DCAMPS", 1: "OUTPUT_DCVOLTS"}.get(_source, "OUTPUT_DCVOLTS")
        self.supply.write("smua.source.func = smua." + source)

    def set_output(self, level=0):
        """
        Sets the output level of the supply
        :param level: Level in volts/amps
        :return: None
        :note: Will turn the power supply on as a side effect
        """

        self.supply.write("smua.source.levelv = " + str(level))
        self.enable_output(True)

    def __configure_compliance(self, limit=0.1):
        """
        Sets compliance level of device
        :param limit: Compliance level of device
        :return: None
        """

        self.supply.write("smua.source.limiti = " + str(limit))

    def configure_measurement(self, _func=1, output_level=0, compliance=0.1):
        """
        Sets up the measurement parameters in one fell swoop
        Will turn the power supply on!
        :param _func: 0 for sourcing current, 1 for voltage
        :param output_level: Set what level the supply should output
        :param compliance: limiting parameter of supply
        :return: None
        """

        self.__configure_source(_func)
        self.__configure_compliance(compliance)
        self.set_output(output_level)

    def enable_output(self, out=False):
        """
        Turns power supply output on or off
        :param out: True for on, False for off
        :return: None
        """

        if out:
            self.supply.write("smua.source.output = 1")
            return
        self.supply.write("smua.source.output = 0")

    def get_current(self):
        """
        Query the device for a current reading
        :return: float representation of the current measured
        """

        return float(self.supply.query("printnumber(smua.measure.i())").split("\n")[0])


class Keithley2400(PowerSupplyFactory):
    """
    Object which control the Keithley 2400 series high voltage supplies
    """

    def __init__(self, gpib_address=22):
        """
        Initializer for keithley power supply
        Checks to make sure that the power supply is not on
        :param gpib_address: GPIB Address of power supply
        """

        assert (gpib_address >= 0), "Please enter a valid gpib_address address"

        print("Initializing Keithley 2400")
        resource_manager = pyvisa.ResourceManager()
        self.supply = resource_manager.open_resource(resource_manager.list_resources()[0])
        for address in resource_manager.list_resources():
            if str(gpib_address) in address:
                print("Keithley 2400 found")
                self.supply = resource_manager.open_resource(address)
            else:
                print("Keithley not found; Please check GPIB Address")

        # Sanity check on device
        print(self.supply.query("*IDN?;"))

        # Reset the state of the device
        self.supply.write("*RST;")

        # Device Defaults
        self.supply.write("*ESE 1;*SRE 32;*CLS;:FUNC:CONC ON;:FUNC:ALL;:TRAC:FEED:CONT NEV;:RES:MODE MAN;")

        # Check for errors and see if we're ready to rock 'n roll
        print(self.supply.query(":SYST:ERR?;"))

        # Display current on front panel
        self.supply.write("SYST:KEY 22;")
        # self.supply.write(":DISP:CND;")

        # Timeout, as usual isnt too big of a deal
        self.supply.timeout = 10000

    def configure_measurement(self, _meas=1, output_level=0, compliance=5e-6):
        """
        Set what the supply is measuring, its output level, and compliance limit
        :param _meas: Whether we are measuring voltage or current
         *** IMPORTANT TO NOTE this is opposite of what we are sourcing ***

        :param output_level: Output level of source
        :param compliance: Maximum value of measured function.
        :return: None
        """

        meas = {0: ":VOLT", 1: ":CURR", 2: ":RES"}.get(_meas, ":CURR")
        # Enable autoranging
        self.supply.write(meas + ":RANG:AUTO ON;")
        self.__configure_source(0, output_level, compliance)

    def __configure_source(self, _func=0, output_level=0, compliance=5e-6):
        """
        This function describes methods of setting the output function
        and limiting voltage and current factors of the power supply source

        :param _func: 0 for sourceing voltage, 1 for sourcing current
        :param output_level: output level of sourcing device in volts/amps
        :param compliance: Current or voltage limit for the device
        :return: None
        """

        assert (0 <= _func < 2), "Invalid Source function"
        assert (-1100 < output_level < 1100), "Voltage out of range"
        assert (compliance < 0.5), "Compliance out of range"

        func = {0: "VOLT", 1: "CURR"}.get(_func, "VOLT")

        self.supply.write("SOUR:FUNC " + func +
                          ";:SOUR:" + func + " " + str(float(output_level)) +
                          ";:CURR:PROT " + str(float(compliance)) + ";")

    def set_output(self, level=0):
        """
        Sets the output level of the device
        Also changes the output state of the supply to ON
        :param level: Output level of device
        :return: None
        """

        self.supply.write(":SOUR:VOLT " + str(float(level)))
        self.enable_output(True)

    def enable_output(self, output=False):
        """
        Sets the output state of the power supply
        :param output: True for on
        :return: None
        """

        if output is True:
            self.supply.write("OUTP ON;")
            return
        self.supply.write("OUTP OFF;")

    def __configure_multipoint(self, arm_count=1, trigger_count=1, mode=0):
        """
        Arms the trigger of the device

        :param arm_count: How many arms for the trigger
        :param trigger_count: How many triggers
        :param mode: 0 for fixed arming, 1 for sweep arming, and 2 for list
        :return: None
        """

        assert (0 <= mode < 3), "Invalid mode"
        source_mode = {0: "FIX", 1: "SWE", 2: "LIST"}.get(mode)
        self.supply.write(":ARM:COUN " + str(arm_count) +
                          ";:TRIG:COUN " + str(trigger_count) +
                          ";:SOUR:VOLT:MODE " + source_mode +
                          ";:SOUR:CURR:MODE " + source_mode)

    def __configure_trigger(self, delay=0.0):
        """
        Set the delay of the armed trigegr
        :param delay: Delay in seconds
        :return: None
        """

        self.supply.write("""ARM:SOUR IMM;
                            :ARM:TIM 0.010000;
                            :TRIG:SOUR IMM;
                            :TRIG:DEL """ + str(delay) + ";")

    def __initiate_trigger(self):
        """
        Arms the supply's trigger count
        :return: None
        """

        self.supply.write(":TRIG:CLE;:INIT;")

    def __wait_operation_complete(self):
        """
        Blocks for return of computer values
        :return: None
        """

        self.supply.write("*OPC;")

    def __fetch_measurements(self):
        """
        Gets the read measurements from the supply
        :return: Current in float format
        """

        read_bytes = self.supply.query(":FETC?")
        return float(read_bytes.split(",")[1])

    def get_current(self, delay=0):
        """
        High level trigger arming and execution
        :param delay: Not needed, but seconds to delay the trigger by
        :return: Single current measurement
        """

        self.__configure_multipoint()
        self.__configure_trigger(delay)
        self.__initiate_trigger()
        self.__wait_operation_complete()
        return self.__fetch_measurements()
