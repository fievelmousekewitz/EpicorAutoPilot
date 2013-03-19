#OpenEdgeApp
import re
import time
import command

class OpenEdgeApp:
    def __init__(self,SF):
        #Settings file must retain 'SettingName: Value' format
        #First load defaults:
        self.Settings = {
            'OpenEdgeDir' : r'c:\epicor\oe102a',
            'EpicorDBDir' : 'c:\\epicor\\epicor905\\db',
            'EpicorBackupDir' : 'c:\\epicor\epicor905\\db',
            'DSN' : 'pilot',
            'Default Company Name' : '--TEST SERVER--1',
            'Database Name' : 'mfgsys',
            'AppServerURL' : 'AppServerDC://localhost:9433',
            'MfgSysAppServerURL' : 'AppServerDC://localhost:9431',
            'FileRootDir' : '\\apollo\epicor\epicor905,C:\epicor\epicordataPilot,\\apollo\epicor\epicor905\server',
            'DBName' : 'EpicorPilot905',
            'AppName' : 'EpicorPilot905',
            'ProcName' : 'EpicorPilot905ProcessServer',
            'TaskName' : 'EpicorPilot905TaskAgent',
            'ProBkup' : '\\bin\\probkup.bat',
            'ProRest' : '\\bin\\prorest.bat',
            'ASBMan' : '\\bin\\asbman.bat',
            'DBMan' : '\\bin\\dbman.bat',
            'Sleepytime' : '3'
            }
        #Now attempt to load from file:
        SettingsFile = open(SF, "r")
        for line in SettingsFile:
            linetranslate = re.match(r'(.*):\s(.*)',line,re.I) #Regex, split everything on the left of : to group1, everything to the right of : to group2
            self.Settings[linetranslate.group(1)] = linetranslate.group(2) #this is hilariously sloppy

    def GetSettings(self,filename):
        #Settings file must retain 'SettingName: Value' format
        SettingsFile = open(filename, "r")
        for line in SettingsFile:
            linetranslate = re.match(r'(.*):\s(.*)',line,re.I) #Regex, split everything on the left of : to group1, everything to the right of : to group2
            self.Settings[linetranslate.group(1)] = linetranslate.group(2) #this is hilariously sloppy

    def Backup(self,source,target,online): #online = "" or "online"
        log = Command(self.Settings['OpenEdgeDir'] + self.Settings['ProBkup'] + " " + online + " " + source + " " + target).run()
        return log.output
    
    def Restore(self,filename):
        log = Command(self.Settings['OpenEdgeDir'] + self.Settings['ProRest'] + " " + self.Settings['EpicorDBDir'] + "\\" + self.Settings['Database Name'] + " " + filename).run()
        return log.output
    
    def Shutdown(self):
        log = ""
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['ProcName'] + " -stop").run(); log+=cmdlog.output
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['TaskName'] + " -stop").run() + "\n"; log+=cmdlog.output
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['AppName']  + " -stop").run() + "\n"; log+=cmdlog.output
        return log

    def ShutdownDB(self):
        log = ""
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['DBMan'] + " -db " + self.Settings['DBName'] + " -stop").run() + "\n"; log+=cmdlog.output
        return log
       
    def StartupDB(self,retries):
        log = ""
        while retries > 0:
            log += "Trying to Start DB...\n"
            time.sleep(int(self.Settings['Sleepytime']))
            cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['DBMan'] + " -db " + self.Settings['DBName'] + " -start").run(); log+=cmdlog.output
            if log == None:
                retries -= 1
                log == "Database Startup Failed.\n"
            else:
                return log
        
    def Startup(self):
         log = ""
         cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['TaskName'] + " -start").run() + "\n"; log+=cmdlog.output
         cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['ProcName'] + " -start").run(); log+=cmdlog.output
         cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['AppName']  + " -start").run() + "\n"; log+=cmdlog.output
         return log
        
        
