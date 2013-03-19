#########################################################################
#
#
#                Epicor Auto Pilot
#
# Simple script to automate backing up Epicor Live &
# Restoring backups into Pilot.
# Automates changing appserver URLs and DB name via pypyodbc,
# you will need an ODBC DSN setup on your machine (and working).
# All settings can be found in the settings.txt file. Please configure
# for your environment before running this script. Use this at your own
# risk. I accept zero responsability for any damage this may do to your
# unique environment. It works for mine, that's all I can tell you.
# If you don't understand how any of this works, I would recommend you not use
# it. Misconfigured, this script could do some baaaaad things.
#
#
# * NONE of these settings should EVER directly reference your live environment,
#   OR any DSN's pointing to your live environment. EVER.
#
# * This will only work in single company environments. I do not use a multicompany
#   environment, so I don't have reason to write this specifically for that, nor do
#   I have the resources/time to test it. Updating for multicompany should only
#   be a matter of updating SQL statemetns though.
#
#
#
#       Feel free to use this, modify it, distribute it, whatever.
#       Just don't sell it.
#       -Jeff Johnson, Angeles Composite Technologies, Inc.
#
#  Credit goes to A Mercer Sisson from the EUG for writing the original
#  WinBatch version.
#
#
#
#   Dependancies you will need: (both can be found on SourceForge)
#   * EasyGUI
#   * pypyodbc
#  
#########################################################################

import easygui
import EpicorDatabaseClass
import OpenEdge

PilotApp = OpenEdge.OpenEdgeApp("settings.txt")
PilotDB = EpicorDatabaseClass.EpicorDatabase()

def GetDSNPassword():
    msg = "Enter logon information"
    title = "DSN Info"
    fieldNames = ["User ID", "Password"]
    fieldValues = []  # we start with blanks for the values
    fieldValues = multpasswordbox(msg,title, fieldNames)
    
    # make sure that none of the fields was left blank
    while 1:
        if fieldValues == None: break
        errmsg = ""
        for i in range(len(fieldNames)):
            if fieldValues[i].strip() == "":
                errmsg = errmsg + ('"%s" is a required field.\n\n' % fieldNames[i])
        if errmsg == "": break # no problems found
        fieldValues = multpasswordbox(errmsg, title, fieldNames, fieldValues)
    
while True:
    ServerSettingsTitles = ["Company Name:", "Password"]
    ServerSettings = [PilotApp.Settings['Default Company Name'],"***"]
    choice = easygui.buttonbox('','Epicor AutoPilot',['Backup LIVE now',
                                                      'Restore PILOT from backup',
                                                      'Display Default Settings',
                                                      'Test DB Connection',
                                                      'SetParams',
                                                      'Quit'],image='epicor.gif')
    if choice == 'Backup LIVE now':
        print 'Not implemented'
        #--------BACKUP LIVE----------
    elif choice == 'SetParams':
        #Connect
        usrpw = GetDSNPassword()
        PilotDB.Connect(usrpw["User ID"],usrpw["Password"]); del usrpw
        #Update to Pilot Settings
        PilotDB.Sql("UPDATE pub.company set name = \'" + PilotApp.Settings['Default Company Name'] + "\'")
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
    elif choice == 'Restore PILOT from backup':
        #Shutdown Pilot
        PilotApp.Shutdown()
        PilotApp.ShutdownDB()
        #Restore From File
        RestoreFile = easygui.fileopenbox("Select File To Restore From","Select File To Restore From",PilotApp.Settings['EpicorBackupDir'])
        PilotApp.Restore(RestoreFile)
        PilotApp.StartupDB(5) #Number of retries
        #Connect
        usrpw = GetDSNPassword()
        PilotDB.Connect(usrpw["User ID"],usrpw["Password"]); del usrpw
        #Update to Pilot Settings
        PilotDB.Sql("UPDATE pub.company set name = \'" + PilotApp.Settings['Default Company Name'] + "\'")
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
        
     # Need to add func to delete all mfgsys* in pilot dir, otherwise restore fails/prompts for yn
    elif choice == 'Display Default Settings':
        print 'Not Implemented'
        #Disp
    elif choice == 'Test DB Connection':
         #Connect
        usrpw = GetDSNPassword()
        PilotDB.Connect(usrpw["User ID"],usrpw["Password"]); del usrpw
        PilotDB.Rollback()
        PilotDB.Close()
    elif choice == 'Quit':
        print 'Not Implemented'
        #Quit
        break
    else:
        print "Not Implemented"
        break
        #Quit
