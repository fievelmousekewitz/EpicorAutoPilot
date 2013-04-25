#Epicor AutoPilot

Small toolbox to:
* backup live database
* restore backup to pilot database
* run & clean progress tabanalys output into CSV
* We shall seeeeeeeeeeeeeeee.......

![alt text](http://puu.sh/2GRwN.png "AutoPilot screenshot")

This is intended to be a fully automated tool. Restoring from a 'live' backup to pilot will also reset data directories, proc & task agent ports, and set the company name to your choosing.

The tabananalys feature runs a progress tabanalys then cleans the output into a CSV ready to sort with Excel.

###ProUtil Tabanalys Output:

![alt text](http://puu.sh/2GRDB.png "Messy tabanalys output")

###AutoPilot Tabanalys Cleaned output:

![alt text](http://puu.sh/2GRKD.png "Cleaned up output")


Depends: 
* EasyGUI - Because tkinter gives me a headache 
* pypyodbc - Connect to ODBC DSN's

Known Issue: When installing this on a 64bit server the odbc drivers included in the OpenEdge install are 64bit and will not work with this script. You must run the Windows\SysWow64\odbcad32.exe setup and install drivers from your 32 bit admin client. Note that you cannot just run 'odbcad32', you MUST include the full path when running.

ODBC Driver must be set to Serializable connection type.

###YOU MUST CHANGE THE PATHS AND PORTS IN THE SETTINGS.TXT FILE TO FIT YOUR ENVIRONMENT.

Do whatever you like with this script. Clean it up, share it with friends, whatever. Just don't sell it; Hire me instead. :)
