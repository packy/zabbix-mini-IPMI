import importlib.util
import subprocess
import sys
import re

###########################################################################
#
# If the WMI module isn't available, we can simulate it through running
# Get-CimInstance through Powershell.  It's slower, but it means the end
# user doesn't need to install package that doesn't come with Python.
#

# class to simulate the objects returned by wmi.WMI()
class FauxStructure:
    _dict = {}
    def __init__(self, **entries):
        self._dict = entries
    def __getattr__(self, name):
        if name in self._dict:
            return self._dict[name]
        else:
            raise AttributeError(
                "'{}' object has no attribute "
                "or key '{}'".format(
                    self.__class__.__name__, name
                )
            )
    def setSMBiosData(self, data):
        self._dict['SMBiosData'] = data

# class to get the data returned by wmi.WMI()
# and stuff it into FauxStructure objects
class FauxWMI(object):
    def __init__(self, namespace):
        self.namespace = 'root\\' + namespace
        self.cmdlet = 'Get-CimInstance'
        self.cmd = 'powershell -c'
    def run(self, wmiclass):
        cmd = '%s %s -Namespace "%s" -Class %s' % (
            self.cmd, self.cmdlet, self.namespace, wmiclass
        )
        p = subprocess.check_output(cmd, universal_newlines=True)
        lines = p.splitlines()
        records = []
        thisRec = {}
        for line in lines:
            if thisRec and not line:
                records.append(FauxStructure(**thisRec))
                thisRec = {}
            elif re.search(r'^\S+\s+:\s+', line):
                fields = re.split(r'\s+:\s+', line)
                thisRec[fields[0]] = fields[1]
        return records
    def Hardware(self):
        return self.run('Hardware')
    def Sensor(self):
        return self.run('Sensor')
    def WMINET_Instrumentation(self):
        return self.run('WMINET_Instrumentation')
    def MSSMBios_RawSMBiosTables(self):
        records = self.run('MSSMBios_RawSMBiosTables')
        # run a command to get the data under SMBiosData, since the
        # command above only gets the first four bytes SMBiosData and
        # then elides the rest with an ellipsis.
        cmd = '%s $x = %s -Namespace "%s" -Class %s; $x.SMBiosData' % (
            self.cmd, self.cmdlet, self.namespace, 'MSSMBios_RawSMBiosTables'
        )
        p = subprocess.check_output(cmd, universal_newlines=True)
        # turn the data into bytes and stuff it into
        # the 'SMBiosData' accessor on the record
        records[0].setSMBiosData([int(n) for n in p.splitlines()])
        return records

# factory class that will look to see if the WMI module
# is available, and, if so load it and return wmi.WMI()
# for the requested namespace. If the WMI module isn't
# available, return an instance of our FauxWMI() for
# the requested namespace.
class MyWMI(object):
    @classmethod
    def getObject(cls, namespace):
        if 'wmi' in sys.modules:
            # the wmi module is already loaded
            return sys.modules['wmi'].WMI(namespace=namespace)
        elif (spec := importlib.util.find_spec('wmi')) is not None:
            # the wmi module isn't loaded yet, but it's available
            module = importlib.util.module_from_spec(spec)
            sys.modules['wmi'] = module
            spec.loader.exec_module(module)
            return sys.modules['wmi'].WMI(namespace=namespace)
        else:
            # simulate the wmi module
            return FauxWMI(namespace=namespace)

###########################################################################
#
# start of code translated from Hardware/SMBIOS.cs
# in https://github.com/openhardwaremonitor/openhardwaremonitor
#

class Structure(object):
    def __init__(self, type, handle, data, strings):
        self.type = type;
        self.handle = handle;
        self.data = data;
        self.strings = strings;
    def GetByte(self, offset):
        if (offset < len(self.data) and self.data[offset] >= 0):
            return self.data[offset]
        else:
            return 0
    def GetWord(self, offset):
        if (offset + 1 < len(self.data) and self.data[offset] >= 0):
            return((self.data[offset + 1] << 8) | self.data[offset])
        else:
            return 0
    def GetString(self, offset):
        if (offset < len(self.data) and
            self.data[offset] > 0 and
            self.data[offset] <= len(self.strings)):
            return self.strings[self.data[offset] - 1]
        else:
            return ""
    def Type(self):
        return self.type
    def Handle(self):
        return self.handle
    def __str__(self):
        strings = "'\n                '".join(self.strings)
        if len(self.strings) > 0:
            strings = f"[\n                '{strings}'\n              ]"
        else:
            strings = '[]'

        return(
            f"instance of Structure:\n"
            f"    Type    = '{self.type}'\n"
            f"    Handle  = '{self.handle}'\n"
            f"    Strings = {strings}"
        )

class BIOSInformation(Structure):
    def __init__(self, p1, p2, data=None, strings=None):
        if data and strings:
            super().__init__(p1, p2, data, strings)
            self.vendor  = self.GetString(0x04)
            self.version = self.GetString(0x05)
        else:
            super().__init__(0x00, 0, None, None)
            self.vendor  = type
            self.version = handle
    def Vendor(self):
        return self.vendor
    def Version(self):
        return self.version
    def __str__(self):
        return(
            f"instance of BIOSInformation:\n"
            f"    Vendor  = '{self.vendor}'\n"
            f"    Version = '{self.version}'"
        )

class SystemInformation(Structure):
    def __init__(self, p1, p2, p3=None, p4=None, p5=None):
        if p3 and p4 and p5:
            super().__init__(0x00, 0, None, None)
            self.manufacturerName = p1
            self.productName      = p2
            self.version          = p3
            self.serialNumber     = p4
            self.family           = p5
        else:
            super().__init__(p1, p2, p3, p4)
            self.manufacturerName = self.GetString(0x04)
            self.productName      = self.GetString(0x05)
            self.version          = self.GetString(0x06)
            self.serialNumber     = self.GetString(0x07)
            self.family           = self.GetString(0x1A)
    def ManufacturerName(self):
        return self.manufacturerName
    def ProductName(self):
        return self.productName
    def Version(self):
        return self.version
    def SerialNumber(self):
        return self.serialNumber
    def Family(self):
        return self.family
    def __str__(self):
        return(
            f"instance of SystemInformation:\n"
            f"    ManufacturerName = '{self.manufacturerName}'\n"
            f"    ProductName      = '{self.productName}'\n"
            f"    Version          = '{self.version}'\n"
            f"    SerialNumber     = '{self.serialNumber}'\n"
            f"    Family           = '{self.family}'"
        )

class BaseBoardInformation(Structure):
    def __init__(self, p1, p2, p3, p4):
        if isinstance(p3, str) and isinstance(p4, str):
            super().__init__(0x00, 0, None, None)
            self.manufacturerName = p1
            self.productName      = p2
            self.version          = p3
            self.serialNumber     = p4
        else:
            super().__init__(p1, p2, p3, p4)
            self.manufacturerName = self.GetString(0x04).strip()
            self.productName      = self.GetString(0x05).strip()
            self.version          = self.GetString(0x06).strip()
            self.serialNumber     = self.GetString(0x07).strip()
    def ManufacturerName(self):
        return self.manufacturerName
    def ProductName(self):
        return self.productName
    def Version(self):
        return self.version
    def SerialNumber(self):
        return self.serialNumber
    def __str__(self):
        return(
            f"instance of BaseBoardInformation:\n"
            f"    ManufacturerName = '{self.manufacturerName}'\n"
            f"    ProductName      = '{self.productName}'\n"
            f"    Version          = '{self.version}'\n"
            f"    SerialNumber     = '{self.serialNumber}'"
        )

class ProcessorInformation(Structure):
    def __init__(self, p1, p2, p3, p4):
        super().__init__(p1, p2, p3, p4)
        self.ManufacturerName = self.GetString(0x07).strip()
        self.Version          = self.GetString(0x10).strip()
        self.CoreCount        = self.GetByte(0x23)
        self.CoreEnabled      = self.GetByte(0x24)
        self.ThreadCount      = self.GetByte(0x25)
        self.ExternalClock    = self.GetWord(0x12)
    def __str__(self):
        return(
            f"instance of ProcessorInformation:\n"
            f"    ManufacturerName = '{self.ManufacturerName}'\n"
            f"    Version          = '{self.Version}'\n"
            f"    CoreCount        = '{self.CoreCount}'\n"
            f"    CoreEnabled      = '{self.CoreEnabled}'\n"
            f"    ThreadCount      = '{self.ThreadCount}'\n"
            f"    ExternalClock    = '{self.ExternalClock}'"
        )

class MemoryDevice(Structure):
    def __init__(self, p1, p2, p3, p4):
        super().__init__(p1, p2, p3, p4)
        self.deviceLocator    = self.GetString(0x10).strip()
        self.bankLocator      = self.GetString(0x11).strip()
        self.manufacturerName = self.GetString(0x17).strip()
        self.serialNumber     = self.GetString(0x18).strip()
        self.partNumber       = self.GetString(0x1A).strip()
        self.speed            = self.GetWord(0x15)
    def DeviceLocator(self):
        return self.deviceLocator
    def BankLocator(self):
        return self.bankLocator
    def ManufacturerName(self):
        return self.manufacturerName
    def SerialNumber(self):
        return self.serialNumber
    def PartNumber(self):
        return self.partNumber
    def Speed(self):
        return self.speed
    def __str__(self):
        return(
            f"instance of MemoryDevice:\n"
            f"    DeviceLocator    = '{self.deviceLocator}'\n"
            f"    BankLocator      = '{self.bankLocator}'\n"
            f"    ManufacturerName = '{self.manufacturerName}'\n"
            f"    SerialNumber     = '{self.serialNumber}'\n"
            f"    PartNumber       = '{self.partNumber}'\n"
            f"    Speed            = '{self.speed}'"
        )

class SMBios(object):
    def __init__(self):
        self.table = []
        self.biosInformation = None
        self.systemInformation = None
        self.baseBoardInformation = None
        self.processorInformation = None
        self.memoryDeviceList = []
        c = MyWMI.getObject(namespace="WMI")
        for bios in c.MSSMBios_RawSMBiosTables():
            self.Version = (
                str(bios.SmbiosMajorVersion) + '.' + str(bios.SmbiosMinorVersion)
            )
            self.decodeRawSMBiosData(bios)
    def decodeRawSMBiosData(self, bios):
        raw = bios.SMBiosData
        if not raw:
            return
        offset = 0
        type = raw[offset]
        while (offset + 4 < len(raw) and type != 127):
            type = raw[offset]
            length = raw[offset + 1]
            handle = ((raw[offset + 2] << 8) | raw[offset + 3])
            if (offset + length > len(raw)):
                break
            data = raw[offset:offset+length]
            offset += length
            stringsList = []
            # if string is null terminated
            if (offset < len(raw) and raw[offset] == 0):
                offset += 1
            while (offset < len(raw) and raw[offset] != 0):
                sb = ''
                while (offset < len(raw) and raw[offset] != 0):
                    sb += chr(raw[offset])
                    offset += 1
                offset += 1
                stringsList.append(sb)
            offset += 1

            if (type == 0x00):
                self.biosInformation = BIOSInformation(type, handle, data, stringsList)
                self.table.append(self.biosInformation)
            elif (type == 0x01):
                self.systemInformation = SystemInformation(type, handle, data, stringsList)
                self.table.append(self.systemInformation)
            elif (type == 0x02):
                self.baseBoardInformation = BaseBoardInformation(type, handle, data, stringsList)
                self.table.append(self.baseBoardInformation)
            elif (type == 0x04):
                self.processorInformation = ProcessorInformation(type, handle, data, stringsList)
                self.table.append(self.processorInformation)
            elif (type == 0x11):
                m = MemoryDevice(type, handle, data, stringsList)
                self.memoryDeviceList.append(m)
                self.table.append(m)
            else:
                self.table.append(Structure(type, handle, data, stringsList))
        return

#
# end of code translated from Hardware/SMBIOS.cs
# in https://github.com/openhardwaremonitor/openhardwaremonitor
#
###########################################################################

# an item of Hardware data
class HardwareItem(object):
    def __init__(self, obj):
        self.HardwareType = obj.HardwareType
        self.Identifier = obj.Identifier
        self.InstanceId = obj.InstanceId
        self.Name = obj.Name
        self.Parent = obj.Parent
    def __str__(self):
        return(
            f"instance of HardwareItem:\n"
            f"    HardwareType = '{self.HardwareType}'\n"
            f"    Identifier   = '{self.Identifier}'\n"
            f"    Parent       = '{self.Parent}'\n"
            f"    Name         = '{self.Name}'\n"
            f"    InstanceId   = '{self.InstanceId}'"
        )

# an item of Sensor data
class SensorValue(object):
    # SensorType Units used and Recommended display format (Suffix) from
    # http://openhardwaremonitor.org/wordpress/wp-content/uploads/2011/04/OpenHardwareMonitor-WMI.pdf
    typeMap = {
        'Voltage':     { 'Unit': 'Volt',                   'Suffix': ' V'   },
        'Clock':       { 'Unit': 'Megahertz',              'Suffix': ' MHz' },
        'Temperature': { 'Unit': 'Celsius',                'Suffix': ' °C'   },
        'Load':        { 'Unit': 'Percentage',             'Suffix': '%'    },
        'Fan':         { 'Unit': 'Revolutions per minute', 'Suffix': ' RPM' },
        'Flow':        { 'Unit': 'Liters per hour',        'Suffix': ' L/h' },
        'Control':     { 'Unit': 'Percentage',             'Suffix': '%'    },
        'Level':       { 'Unit': 'Percentage',             'Suffix': '%'    },
        # not in the documentation, guessed!
        'Power':       { 'Unit': 'Watts',                  'Suffix': ' W'   },
        # no guesses for units for Data, SmallData, and Throughput
    }
    def __init__(self, obj):
        self.Identifier = obj.Identifier
        self.Index = obj.Index
        self.InstanceId = obj.InstanceId
        self.Max = obj.Max
        self.Min = obj.Min
        self.Name = obj.Name
        self.Parent = obj.Parent
        self.SensorType = obj.SensorType
        self.Value = obj.Value
        if obj.SensorType in self.typeMap:
            self.Unit   = self.typeMap[obj.SensorType]['Unit']
            self.Suffix = self.typeMap[obj.SensorType]['Suffix']
        else:
            self.Unit   = 'unknown'
            self.Suffix = ''
    def format(self, field, precision):
        return precision.format(getattr(self, field))+self.Suffix
    def __str__(self):
        return(
            f"instance of SensorValue:\n"
            f"    Identifier = '{self.Identifier}'\n"
            f"    Parent     = '{self.Parent}'\n"
            f"    SensorType = '{self.SensorType}'\n"
            f"    Index      = '{self.Index}'\n"
            f"    Name       = '{self.Name}'\n"
            f"    Min        = '{self.Min}{self.Suffix}'\n"
            f"    Value      = '{self.Value}{self.Suffix}'\n"
            f"    Max        = '{self.Max}{self.Suffix}'\n"
            f"    Unit       = '{self.Unit}'\n"
            f"    InstanceId = '{self.InstanceId}'"
        )

# object to get the version of OpenHardwareMonitor used
class OpenHardwareMonitor(object):
    def __init__(self):
        self.version = None
        c = MyWMI.getObject(namespace="OpenHardwareMonitor")
        for item in c.WMINET_Instrumentation():
            OHMRver = re.search(r'Version=([^,]+)', item.FullName, re.I | re.M)
            if OHMRver:
                self.version = OHMRver.group(1).strip()
    def Version(self):
        return self.version
    def __str__(self):
        return(
        f"instance of OpenHardwareMonitor:\n"
        f"    Version = '{self.version}'"
    )

def zfill_match(match):
    # pad matches with zeros to three digits
    return match.group(1).zfill(3)

# base class for Hardware and Sensor data objects
class OpenHardwareMonitorData(object):
    itemClass = {
        'Hardware': 'HardwareItem',
        'Sensor':   'SensorValue'
    }
    def __init__(self):
        self.table = []
        self.tree = {}
        c = MyWMI.getObject(namespace="OpenHardwareMonitor")
        cls = self.__class__.__name__
        # loop over c.Hardware() or c.Sensor()
        for item in getattr(c, cls)():
            # instantiate either HardwareItem() or SensorValue()
            i = globals()[self.itemClass[cls]](item)
            self.table.append(i)
            self.tree[i.Identifier] = i
    def sort(self): # sort the table by identifier
        # pad numbers with zeros to sort "1, 10, 100, 2" as "001, 002, 010, 100"
        self.table = sorted(self.table, key=lambda k: re.sub(r'(?<=\/)(\d+)$', zfill_match, k.Identifier))
        return self # so we can chain
    def keys(self):
        return [ i.Identifier for i in self.table ]

class Hardware(OpenHardwareMonitorData):
    pass # all the work is done in the base class

class Sensor(OpenHardwareMonitorData):
    pass # all the work is done in the base class
