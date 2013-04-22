Epicor AutoPilot
-----------------
Small toolbox to: 
* backup live database and restore to pilot database
* clean tabanalys output into CSV
* We shall seeeeeeeeeeeeeeee.......


ODBC Driver must be set to Serializable connection type.

Depends:
EasyGUI - Because tkinter gives me a headache
pypyodbc - Connect to ODBC DSN's


Known Issue:
When installing this on a 64bit server the odbc drivers included in the OpenEdge install are 64bit and will not work with this script. You must run the Windows\SysWow64\odbcad32.exe setup and install drivers from your 32 bit admin client. Note that you cannot just run 'odbcad32', you MUST include the full path when running.