import sys
import re
import platform
from operator import itemgetter
from mini_ipmi_wmi import SMBios, Hardware, Sensor, OpenHardwareMonitor


WMI_NAMESPACE = 'root\OpenHardwareMonitor'
WMI_CMD = 'powershell -c Get-WmiObject'
ACTION = 'get'


if __name__ == '__main__':
    hardware = Hardware().sort()
    sensor = Sensor().sort()
    bios = SMBios()
    maxSensorNameWidth = 0
    maxSensorValWidth  = 0
    for item in sensor.table:
        maxSensorNameWidth = max(maxSensorNameWidth, len(item.Name))
        maxSensorValWidth  = max(maxSensorValWidth,
            len(item.format('Min',   '{:.2f}')),
            len(item.format('Value', '{:.2f}')),
            len(item.format('Max',   '{:.2f}'))
        )

    cpus   = [ k for k in hardware.table if 'cpu' in k.Identifier]
    gpus   = [ k for k in hardware.table if 'gpu' in k.Identifier]
    hdds   = [ k for k in hardware.table if 'hdd' in k.Identifier]
    others = [ k for k in hardware.table if
                 not 'cpu'        in k.Identifier and
                 not 'gpu'        in k.Identifier and
                 not 'hdd'        in k.Identifier and
                 not '/mainboard' in k.Identifier and
                 not '/ram'       in k.Identifier
             ]
    hardwareitems = [hardware.tree['/mainboard']] + cpus + gpus + [hardware.tree['/ram']] + others + hdds

    memory = []
    for item in bios.table:
        if item.__class__.__name__ == 'BaseBoardInformation':
            baseboardinfo = item
        elif item.__class__.__name__ == 'BIOSInformation':
            biosinfo = item
        elif item.__class__.__name__ == 'ProcessorInformation':
            cpuinfo = item
        elif item.__class__.__name__ == 'MemoryDevice':
            memory.append(item)

    hardwarerow='+- {} ({})'
    biosrow='|  +- {:<%s} :  {}' % (maxSensorNameWidth)
    sensorrow='|  +- {:<%s} :  {:>%s}  {:>%s}  {:>%s} ({})' % (maxSensorNameWidth, maxSensorValWidth, maxSensorValWidth, maxSensorValWidth)
    print("Sensors")
    for hw in hardwareitems:
        print("|")
        print(hardwarerow.format(hw.Name, hw.Identifier))
        if hw.Identifier == '/mainboard':
            print(biosrow.format('Manufacturer',  baseboardinfo.manufacturerName))
            print(biosrow.format('Product',       baseboardinfo.productName))
            print(biosrow.format('Version',       baseboardinfo.version))
            print(biosrow.format('Serial Number', baseboardinfo.serialNumber))
            print(biosrow.format('BIOS Vendor',   biosinfo.vendor))
            print(biosrow.format('BIOS Version',  biosinfo.version))
            print(biosrow.format('CPU',           cpuinfo.Version))
            print(biosrow.format('CPU Cores',     cpuinfo.CoreEnabled))
            print(biosrow.format('CPU Threads',   cpuinfo.ThreadCount))
            for dimm in memory:
                if dimm.manufacturerName != 'NO DIMM':
                    print(biosrow.format(dimm.deviceLocator, dimm.manufacturerName + ' ' + dimm.partNumber + ' ('+ str(dimm.speed)+' MHz)'))
                else:
                    print(biosrow.format(dimm.deviceLocator, dimm.manufacturerName))
        else:
            items = [item for item in sensor.table if item.Identifier.startswith(hw.Identifier)]
            for item in items:
                print(sensorrow.format(
                    item.Name,
                    item.format('Min',   '{:.2f}'),
                    item.format('Value', '{:.2f}'),
                    item.format('Max',   '{:.2f}'),
                    item.Identifier
                ))
