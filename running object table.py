import pythoncom
import win32com.client

# must be 0
context = pythoncom.CreateBindCtx(0)

# Get Running Object Table
running_coms = pythoncom.GetRunningObjectTable()

# Create an enumerator that can list all the monikers in our table
monikers = running_coms.EnumRunning()

# Loop through all the monikers
for moniker in monikers:

    print('-'*100)

    # print the display name
    print(moniker.GetDisplayName(context, moniker))

    # print the hash
    print(moniker.Hash())

    # print is System
    print(moniker.IsSystemMoniker())

# now use win32com to see what you can do with the com objects on the list
# ThisApplication = win32com.client.gencache.EnsureDispatch('{}')
# help(ThisApplication)