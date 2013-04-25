#########################################################################
#
#
#                Epicor AutoPilot
#
#  Simple script to automate backing up Epicor Live &
#  Restoring backups into Pilot.
#  Automates changing appserver URLs and DB name via pypyodbc,
#  you will need an ODBC DSN setup on your machine (and working).
#  All settings can be found in the settings.txt file. Please configure
#  for your environment before running this script. Use this at your own
#  risk. I accept zero responsability for any damage this may do to your
#  unique environment. It works for mine, that's all I can tell you.
#  If you don't understand how any of this works, I would recommend you not use
#  it. Misconfigured, this script could do some very bad things.
#
#
# * NONE of the Pilot settings should EVER directly reference your live environment,
#   OR any DSN's pointing to your live environment. EVER.
#
# * This will only work in single company environments. I do not use a multicompany
#   environment, so I don't have reason to write this specifically for that, nor do
#   I have the resources/time to test it. 
#
#
#
#       Feel free to use this, modify it, distribute it, whatever.
#       Just don't sell it.
#       -Jeff Johnson, Angeles Composite Technologies, Inc.
#
#  Credit goes to A Mercer Sisson from the EUG for writing the original
#  WinBatch version of the Pilot restore process.
#
#
#
#   Dependancies you will need: (both can be found on SourceForge)
#   * EasyGUI
#   * pypyodbc
#  
#########################################################################

import easygui
import pypyodbc
import re
import time
import sys
import subprocess
import datetime

def CleanTabAnalys(sourcefile,destfile):
    f = open(sourcefile,'r')
    o = open(destfile,'w')
    linenum = 1
    o.write('\"Tables\",\"Records\",\"Size(bytes)\",\"Record Size Min\",\"Record Size Max\",\"Record Size Mean\",\"Fragment Count\",\"Fragment Factor\",\"Scatter Factor\"')
    for line in f:
            if (line.find("PUB") == -1 and linenum > 0):
                linenum = linenum + 1 
                line =''
                if (linenum > 200):
                    print "Doesn't seem to be a tabanalys"
                    break
            elif (line.find("-------") != -1 and linenum == -1): #end of table data
                    break       
            else:
                linenum = -1
                #My RegEx Foo is weak, forgive me:
                line = re.sub(r'([\s*]{1,25})',',',line.rstrip()) #strip spaces and replace with commas
                line = re.sub(r',,',',',line.rstrip()) #strip double commas
                line = re.sub(r'PUB\.','\nPUB.',line.rstrip()) #insert newlines before PUBs
                line = re.sub(r'(?P<nm>[0-9])(?P<dot>[\.])(?P<nx>[0-9])(?P<sz>[BKMG])','\g<nm>\g<nx>\g<sz>',line.rstrip()) #strip decimals from values, but NOT tables
                line = re.sub(r'(?P<nb>[0-9])B','',line.rstrip()) # strip B's
                line = re.sub(r'(?P<nk>[0-9])K','\g<nk>00',line.rstrip()) #strip K's and add 00 (250.4K becomes 250400 and so on...)
                line = re.sub(r'(?P<nM>[0-9])M','\g<nM>00000',line.rstrip()) #strip M's and add 00000
                line = re.sub(r'(?P<nG>[0-9])G','\g<nG>00000000',line.rstrip()) #strip G's and add 00000000
                #sys.stdout.write(line) #debug
                o.write(line)   
    o.close()
    
#Command Class from DaniWeb written by Gribouillis:
class Command(object):
    """Run a command and capture it's output string, error string and exit status"""
    def __init__(self, command):
        self.command = command 
    def run(self, shell=True):
        import subprocess as sp
        process = sp.Popen(self.command, shell = shell, stdout = sp.PIPE, stderr = sp.PIPE)
        self.pid = process.pid
        self.output, self.error = process.communicate()
        self.failed = process.returncode
        return self
    @property
    def returncode(self):
        return self.failed

class EpicorDatabase:
    def __init__(self,dsn,usr,pwd):
        self.conn = pypyodbc.connect('DSN='+dsn+';UID='+usr+';PWD='+pwd+';')
    
    def Connect(self,dsn,pwd):
        self.conn = pypyodbc.connect('DSN='+dsn+';UID='+usr+';PWD='+pwd+';')

    def Sql(self,statement):
        self.cur = self.conn.cursor()
        self.cur.execute(statement)

    def  Commit(self):
        self.conn.commit()

    def Close(self):
        self.conn.close()

    def Rollback(self):
        self.conn.rollback()

class OpenEdgeApp:
    def __init__(self,SF):
        #Settings file must retain 'SettingName: Value' format
        #First load defaults:
        self.Settings = {
            'OpenEdgeDir' : r'c:\epicor\oe102a',
            'EpicorDBDir' : r'c:\epicor\epicor905\\db',
            'EpicorPilotDBDir' : r'c:\epicor\epicor905\db\pilot',
            'EpicorBackupDir' : r'c:\\epicor\epicor905\\db',
            'DSN' : 'pilot',
            'Default Company Name' : '--TEST SERVER--',
            'Database Name' : 'mfgsys',
            'AppServerURL' : r'AppServerDC://localhost:9433',
            'MfgSysAppServerURL' : r'AppServerDC://localhost:9431',
            'FileRootDir' : r'\\apollo\epicor\epicor905,C:\epicor\epicordataPilot,\\apollo\epicor\epicor905\server',
            'DBName' : 'EpicorPilot905',
            'AppName' : 'EpicorPilot905',
            'ProcName' : 'EpicorPilot905ProcessServer',
            'TaskName' : 'EpicorPilot905TaskAgent',
            'ProBkup' : r'\bin\probkup.bat',
            'ProRest' : r'\bin\prorest.bat',
            'ASBMan' : r'\bin\asbman.bat',
            'DBMan' : r'\bin\dbman.bat',
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
            linetranslate = re.match(r'(.*):\s(.*)',line,re.I) 
            self.Settings[linetranslate.group(1)] = linetranslate.group(2) 

    def Backup(self,source,target,online): #online = "" or "online"
        log = Command(self.Settings['OpenEdgeDir'] + self.Settings['ProBkup'] + " " + online + " " + source + " " + target).run()
        return log.output
    
    def Restore(self,filename):
        #Initiate ProRest. Note the 'echo y' pipe is necessary to overwrite the existing db.
        log = Command('echo y | ' + self.Settings['OpenEdgeDir'] + self.Settings['ProRest'] + " " + self.Settings['EpicorPilotDBDir'] + "\\" + self.Settings['Database Name'] + " " + filename).run()
        return log.output
    
    def Shutdown(self):
        log = ""
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['ProcName'] + " -stop").run(); log+=cmdlog.output
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['TaskName'] + " -stop").run(); log+="\n" + cmdlog.output
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['AppName']  + " -stop").run(); log+="\n" + cmdlog.output
        return log

    def ShutdownDB(self):
        log = ""
        cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['DBMan'] + " -db " + self.Settings['DBName'] + " -stop").run(); log+="\n" + cmdlog.output
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
         cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['TaskName'] + " -start").run(); log+=cmdlog.output
         cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['ProcName'] + " -start").run(); log+="\n" + cmdlog.output
         cmdlog = Command(self.Settings['OpenEdgeDir'] + self.Settings['ASBMan'] + " -name " + self.Settings['AppName']  + " -start").run(); log+="\n" + cmdlog.output
         return log
        
def GetDSNPassword():
    msg = "Enter logon information"
    title = "DSN Info"
    fieldNames = ["User ID", "Password"]
    fieldValues = []  # we start with blanks for the values
    fieldValues = easygui.multpasswordbox(msg,title, fieldNames)
    # make sure that none of the fields were left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
            if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
        if errmsg == "": break # no problems found
        fieldValues = multpasswordbox(errmsg, title, fieldNames, fieldValues)
    return fieldValues

def Choice_Backup_Live():
    rightnow = datetime.datetime.now()
    defaultfilename =  rightnow.strftime("%B")
    defaultfilename = defaultfilename[:3] 
    
    if easygui.ynbox('Online Backup?') == 1:
         ynonline = "online"
    else:
        ynonline = ""
    easygui.msgbox(msg='After selecting Source, Destination & Online please wait, this can take a while and you will see no progress bar.')
    PilotApp.Backup(
        easygui.fileopenbox(
            'Select database to backup',
            default=PilotApp.Settings['EpicorDBDir']+'\\'),
        easygui.filesavebox(
            'Filename to backup to',
            default=PilotApp.Settings['EpicorDBDir']+'\\'+ defaultfilename + 'live' + str(rightnow.day ) ),
        ynonline)

def Choice_Set_Params():
    #Connect
    usrpw = GetDSNPassword()
    PilotDB = EpicorDatabaseClass.EpicorDatabase(PilotApp.Settings['DSN'],usrpw[0], usrpw[1] ); del usrpw
    #Update to Pilot Settings
    compname = easygui.enterbox(msg='New Pilot Company Name?',default=PilotApp.Settings['Default Company Name'])
    PilotDB.Sql("UPDATE pub.company set name = \'" +compname + "\'")
    PilotDB.Sql("UPDATE pub.SysAgent set AppServerURL = \'" + PilotApp.Settings['AppServerURL'] + "\'")
    PilotDB.Sql("UPDATE pub.SysAgent set MfgSysAppServerURL = \'" + PilotApp.Settings['MfgSysAppServerURL'] + "\'")
    PilotDB.Sql("UPDATE pub.SysAgent set FileRootDir = \'" + PilotApp.Settings['FileRootDir'] + "\'")
    #Remove Global Alerts / Task Scheduler
    PilotDB.Sql("UPDATE pub.glbalert set active=0 where active=1")
    PilotDB.Sql("UPDATE pub.chglogGA set SendEmail=0 where SendEmail=1")
    PilotDB.Sql("DELETE from pub.SysAgentTask")
    #Commit changes and close connection
    PilotDB.Commit()
    PilotDB.Close()

def Choice_Restore_Pilot():
    #Restore From File
    RestoreFile = easygui.fileopenbox("Select File To Restore From","Select File To Restore From",PilotApp.Settings['EpicorBackupDir'])
    usrpw = GetDSNPassword()
    rightnow = datetime.datetime.now()
    compname =  rightnow.strftime("%B")
    compname = compname[:3]
    compname = '--- TEST ' + compname + str(rightnow.day) + ' ---'
    compname = easygui.enterbox(msg='New Pilot Company Name?',default=compname)
    easygui.msgbox(msg='This will take some time, please wait until you see the main app dialog.')
    #Shutdown Pilot
    PilotApp.Shutdown()
    PilotApp.ShutdownDB()
    PilotApp.Restore(RestoreFile)
    PilotApp.StartupDB(5) #Number of retries
    #Connect to db
    PilotDB = EpicorDatabase(PilotApp.Settings['DSN'],usrpw[0], usrpw[1] );del usrpw
    #Update to Pilot Settings
    PilotDB.Sql("UPDATE pub.company set name = \'" + compname + "\'")
    PilotDB.Sql("UPDATE pub.SysAgent set AppServerURL = \'" + PilotApp.Settings['AppServerURL'] + "\'")
    PilotDB.Sql("UPDATE pub.SysAgent set MfgSysAppServerURL = \'" + PilotApp.Settings['MfgSysAppServerURL'] + "\'")
    PilotDB.Sql("UPDATE pub.SysAgent set FileRootDir = \'" + PilotApp.Settings['FileRootDir'] + "\'")
    #Remove Global Alerts / Task Scheduler
    PilotDB.Sql("UPDATE pub.glbalert set active=0 where active=1")
    PilotDB.Sql("UPDATE pub.chglogGA set SendEmail=0 where SendEmail=1")
    PilotDB.Sql("DELETE from pub.SysAgentTask")
    #Commit changes and close connection
    PilotDB.Commit()
    PilotDB.Close()
    PilotApp.Startup()
   
def Choice_Test_Connection():
    #Connect
    usrpw = GetDSNPassword()
    PilotDB = EpicorDatabase(PilotApp.Settings['DSN'],usrpw[0], usrpw[1] )
    CurComp = PilotDB.Sql("SELECT name FROM pub.company")
    print CurComp[0]
    easygui.msgbox("Connected to: " + str(CurComp[0]))
    PilotDB.Rollback()
    PilotDB.Close()

def Choice_Run_Tabanalys():
    #Run Tabanalys and clean into CSV file
    now = datetime.datetime.now()
    TabAnalysFile = PilotApp.Settings['EpicorDBDir'] + r'\TabAnalys.tmp'
    TabCSVFile = PilotApp.Settings['EpicorDBDir'] + r'\TabAnalys ' + str(now.year) + '-' + str(now.month) + '-' + str(now.day) + '.csv'
    easygui.msgbox('This may take a few, please be patient.')
    Command(PilotApp.Settings['OpenEdgeDir'] + r'\bin\proutil.bat ' + PilotApp.Settings['EpicorDBDir'] + '\\' + PilotApp.Settings['Database Name'] + ' -C tabanalys > ' + TabAnalysFile).run()
    #print PilotApp.Settings['OpenEdgeDir'] + '\\bin\\proutil.bat ' + PilotApp.Settings['EpicorDBDir'] + '\\' + PilotApp.Settings['Database Name'] + ' -C tabanalys > ' + TabAnalysFile
    tabtmp = open(TabAnalysFile)
    tabcontents = tabtmp.readlines()
    easygui.codebox('Contents','Contents',tabcontents)
    tabtmp.close()
    CleanTabAnalys(TabAnalysFile,TabCSVFile)
    easygui.msgbox('Saved csv (w/ headers) file to ' + TabCSVFile)
    
#------------------------ BEGIN ----------------------

PilotApp = OpenEdgeApp("settings.txt")

while True:
    ServerSettingsTitles = ["Company Name:", "Password"]
    ServerSettings = [PilotApp.Settings['Default Company Name'],"***"]
    choice = easygui.buttonbox('','Epicor AutoPilot',['Backup LIVE now',
                                                      'Restore PILOT from backup',
                                                      'Display Default Settings',
                                                      'Test DB Connection',
                                                      'SetParams',
                                                      'Run Tabanalys',
                                                      'Quit'],image='epicor.gif')
    if choice == 'Backup LIVE now':
        Choice_Backup_Live()
        
    elif choice == 'SetParams':
        Choice_Set_Params()
        
    elif choice == 'Restore PILOT from backup':
        Choice_Restore_Pilot()
        
    elif choice == 'Display Default Settings':
        #TODO
        print 'Not Implemented'

    elif choice == 'Test DB Connection':
        Choice_Test_Connection()
        
    elif choice ==  'Run Tabanalys':
        Choice_Run_Tabanalys()
        
    elif choice == 'Quit':
        break
       #Quit
    else:
        print "Not Implemented"
        break
        #Quit
