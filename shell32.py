from System import Type, Activator
com = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"))
windows = com.Windows()
for ie in windows:
    print(ie.LocationURL())
    ie.Navigate('file:///')