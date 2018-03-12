class LCRMeterFactory(object):
    """
    Abstract Class for creating LCR meter interfaces.
    Call the static method LCRMeter.factory("insert_my_type")
    to get a LCR meter object instance
    """

    def factory(lcr_meter_type, gpib_address):
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

        if lcr_meter_type.lower() == "agilente4980a":
            pass
        else:
            pass
