import clr
clr.AddReference("Interop.shdocvw")
from interop.shdocvw import IShellWindows
#com = Activator.CreateInstance(Type.GetTypeFromProgID("Interop.shdocvw"))
com = IShellWindows()
windows = com.Windows()
for ie in windows:
    print(ie.LocationURL())
    ie.quit()