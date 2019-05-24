import win32com.client
from selenium.webdriver import Firefox
from selenium.common.exceptions import *
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time, datetime, sys
from ProcessError import ProcessingError
from Navigation import Navigation
from config_parser import cfg_parser

"""
Title: OSCAR Manual Amendments Script
Compatibility: Python 3.6 or later
Status: Draft
Author: Danny Ruttle
Version Date: 2019-05-14
Project: ODS 3rd PArty data Automation
Internal Ref: v0.8
Copyright Health and Social Care Information Centre (c) 2019

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.


WHAT THIS CODE DOES:

The module defined below carries out the following procedures:

1. Opens a URL to a page that contains records requiring manual amendments (post import)
2. It authenticates using the Selenium Web Driver and the win32 shell scripting python library
3. It always acts on the record at the top of the list, processing that and then refreshing the page containing manual amendments
4. It searches for the element "MainContent_gvImportGroup_aNotes_0" (this will be in the first entry only)
    if this element is present
        a. the checkbox is clicked to identify that this record should be deferred
        b. the defer button is pressed, which then refreshes the page and the next record is then at index "_0"
    else
        a. searches for the class "MainContent_gvImportGroup_hlProcess_0"
        b. clicks the link within that table cell e.g. href="ImportManualChange.aspx?ImportGroupOrganisationID=3955786"
        c. calls enterTextToClassID() which searches for the audit field and enters the audit text
        d. searches for a button on the target page with the ID "btnSave"
        e. btnSave is then clicked button which then returns the process back to the original URL
        f. it handles for the presence of any errors reported via the OSCAR Error pane and implements a fix for simple errors
        g. it detects closing records and handles the operational closure in the task screen.
4. A log is provided to audit what the script did and report any errors


CHANGES:

0.1 - initial POC with half a dozen Social HQ/Providers
0.2 - modified to action the defer on the same screen
0.3 - version checked in ***without*** TOO MANY REDIRECTS issue fixed
0.4 - Firefox version with first stab at removing too many redirects handlers
0.5 - Truncate short name if that error appears. Webdriver Exception first stab
0.6 - simplified logging and moved creation of the log file to __init__, refactored record processing
      so new methods handleDefer and handle Audit added
0.7 - Added handling to defer re-open requests for manual review.  Refactored Manualamend to run the process with Navigation and 
      ProcessError as separate modules
      Handles records that display errors (as in the broken record detector)
      Logging class/method which opens, appends then closes so progress can be monitored
0.8 - added self.amendCount to logs - this is the loop that keeps a track of the records processed/attempted to process
0.9 - added functionality to read params from a config file

TO DO:

1. Ensure logging is consistent and relevant - done 0.8, needs testing
2. Add logic to prevent infinite call to process() from Webdriver exception - done in v0.6 but not tested
3. Switch to using a config file for running the application
4. Add a crypto type method for storing domain passwords
5. Add support for locked records
6. Move WriteLog into a separate module as it is required by ManualAmend and ProcessError



"""

class ManAmend:

    def __init__(self, configFile):
        """
        :param configFile: path to text file with config
        """
        # process args here ### will be in a config file ###

        self.config_dict = self.read_config(configFile)

        self.env = self.config_dict.get("env")
        self.user = self.config_dict.get("user")
        self.pwd = self.config_dict.get("pwd")
        self.amendID = self.config_dict.get("AmendID")
        self.auditText = self.config_dict.get("AuditText")
        self.processLimit = int(self.config_dict.get("limit"))
        self.driverPath = self.config_dict.get('driverPath')
        self.binaryPath = self.config_dict.get('firefoxBinary')
        self.binary = FirefoxBinary(self.binaryPath)

        self.sleepDuration = 1
        self.sleepDurationLong = 2
        self.navigateLimit = 3 #used to set loops when searching for elements/IDs

        #
        self.deferNextRecord = False # used to defer after discovering a ReOpen request
        self.amendCount = 1

        # logging vars
        self.fileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.logFileName = "ScriptLog_ManualAmends_%s_%s_%s.csv" % (self.env, self.amendID, self.fileTime)
        self.logFile = open(self.logFileName, "w")

        # used to support debugging - so exceptions are reported rather than being handled
        # by the default behaviour to restart the driver
        self.driverRestart = True
        self.restartOutputMessage = ""

    def WriteLog(self, msg):
        """
        Opens up the logfile appends an updated and closes the file
        This behaviour allows the progress to be reviewed during processing
        :param msg: string to be written
        :return: none
        """
        self.logFile = open(self.logFileName, 'a')
        self.logFile.write(msg)
        self.logFile.close()


    def read_config(self, config):
        """
        Read from config file into a dictionary
        :param config: text file with
        :return: dictionary containing config properties
        """
        cp = cfg_parser.config_parser()
        return cp.read(config)

    def process(self):
        """
        """

        self.driver = Firefox(firefox_binary=self.binary, executable_path=self.driverPath)

        # create instance of Navigation class here
        self.nav = Navigation(self.driver)

        #test connect:
        connectURL = "https://%s" % ( self.env )
        print("Connecting to %s......" % connectURL)
        try:
            self.driver.get(connectURL)
            self.handleAuth()
            self.driver.get(connectURL)
            print("Successfully connected to %s......" % connectURL)
        except UnexpectedAlertPresentException:
            print("need to authenticate when accessing home")
            self.handleAuth()
            self.driver.get(connectURL)

        time.sleep(self.sleepDuration) #allow time to login manually

        self.fileWriteTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.WriteLog('"0","0","processing started/restarted","@ %s"\n' % self.fileWriteTime)

        # connect to the environment (not part of the test) to prompt authentication

        self.processState = True

        if self.driverRestart: # supports debug
            try:
                while self.amendCount <= self.processLimit and self.processState == True:
                    self.processAmendment()
                    self.amendCount += 1
            except WebDriverException:
                self.restartOutputMessage = "WebDriverException in process() > Loop"
                self.restartDriver(self.restartOutputMessage)
        else:
            while self.amendCount <= self.processLimit and self.processState == True:
                self.processAmendment()
                self.amendCount += 1

        # CLOSE THE SCRIPT LOGFILE:
        self.endFileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.WriteLog('"%s","0","processing completed","@ %s"\n' % (self.amendCount, self.endFileTime))
        self.driver.quit()

    def restartDriver(self, outputMessage):
        print(outputMessage)
        time.sleep(self.sleepDurationLong)
        self.driver.quit()
        self.restartOutputMessage = "Quiting driver due to %s" % outputMessage
        print(self.restartOutputMessage)
        outputMessage = "Driver disposed now calling process() due to %s" % outputMessage
        print(outputMessage)
        time.sleep(self.sleepDurationLong)
        self.process()

    def processAmendment(self):
        """
        """
        OrganisationID = 0 # Org being processed - required for audit
        AmendmentsURL = "https://%s/Exchange/ImportGroupManualChanges.aspx?ID=%s" % (self.env, self.amendID)

        try:
            self.driver.get(AmendmentsURL)
        except UnexpectedAlertPresentException:
            self.handleAuth()
            self.driver.get(AmendmentsURL)
        except WebDriverException:
            self.restartOutputMessage = "WebDriverException in processAmendment()"
            self.restartDriver(self.restartOutputMessage)

        time.sleep(self.sleepDurationLong)

        try:
            # if we can't get the organisationID then exit
            OrganisationID = self.nav.getOrganisationIDfromProcessLink("MainContent_gvImportGroup_hlProcess_0")
            print("Now processing %s" % OrganisationID)
        except:
            msg = '"0","0","System Exit!","Unable to retrieve OrganisationID"\n'
            self.WriteLog( msg )
            self.procesState = False

        if self.nav.checkNotes("MainContent_gvImportGroup_aNotes_0") or self.deferNextRecord is True:
            # too complex to automate so just defer it
            self.handleDefer( OrganisationID )
            self.deferNextRecord = False #process for orgs that require reviewing as part of the re-open task
            msg = '"%s","%s","Defer","Processing Notes Found"\n' % ( self.amendCount, OrganisationID )
            self.WriteLog( msg )

        else: #attempt to process
            if self.nav.navigateToClassID("MainContent_gvImportGroup_hlProcess_0"):
                time.sleep( self.sleepDuration )
                try:
                    # handle save alert ### need to see more examples of these so scenarios can be tested ###
                    self.driver.switch_to.alert.dismiss()
                except:
                    pass # it never happened!

                if self.nav.navigateToClassID("MainContent_btnSave"):
                    time.sleep(self.sleepDurationLong)
                    """
                    Need to detect if there's an error on the page and handle or bypass 
                    """
                    procErr = ProcessingError(OrganisationID, self.driver, self.logFile, self.logFileName, self.amendCount )
                    if procErr.hasErrors(): # may need to put a conditional here
                        procErr.handleProcessingError()
                    else:
                        self.WriteLog(msg = '"%s","%s","Processed","Routine process completed"\n' % ( self.amendCount, OrganisationID ))
                    # Audit panel and buttons have the same ID and can appear at this point in the process
                    try:
                        self.handleAudit()
                    except:
                        pass

                    # For records that save changes and close tasks
                    self.handleTasks(OrganisationID)
                else:
                    self.handleTasks(OrganisationID)

        return self.processState

    def handleTasks(self, OrganisationID):
        # task screen related process
        try:
            self.handleClose()
            msg = '"%s","%s","ProcessClose","Closure process completed"\n' % (self.amendCount, OrganisationID)
            print(msg)
            self.WriteLog(msg)
            time.sleep(self.sleepDuration)
        except:
            # assume nothing to do here
            pass

        try:
            self.handleReOpenDefer()
            msg = '"%s","%s","ProcessReOpen","ReOpen process flagged to defer"\n' % (self.amendCount, OrganisationID)
            print(msg)
            self.WriteLog(msg)
            time.sleep(self.sleepDuration)
        except:
            # assume nothing to do here
            pass

    def handleAuth(self):
        """
        Simulates what AutoIt did previously
        :return:
        """
        time.sleep(self.sleepDuration)
        print("authentication process called")
        shell = win32com.client.Dispatch("WScript.Shell")
        login = "corp\\%s" % (self.user)
        shell.Sendkeys(login.lower())
        time.sleep(self.sleepDurationLong)
        shell.Sendkeys("{TAB}")
        time.sleep(self.sleepDurationLong)
        shell.Sendkeys(self.pwd)
        time.sleep(self.sleepDurationLong)
        shell.Sendkeys("{ENTER}")
        time.sleep(self.sleepDurationLong)

    def handleClose(self):
        self.nav.navigateToClassID("MainContent_wizTasks_ctl17_chkCloseOperationally")
        time.sleep(self.sleepDurationLong)
        self.nav.navigateToClassID("MainContent_wizTasks_ctl17_btnNextStart")
        time.sleep(self.sleepDurationLong)
        self.nav.navigateToClassID("MainContent_wizTasks_ctl18_btnFinishClose")
        time.sleep(self.sleepDurationLong)
        self.handleAudit()

    def handleReOpenDefer(self):
        # see if the openlegally option is available
        if self.nav.searchforClassID("MainContent_wizTasks_ctl17_chkReOpenLegally"):
            self.deferNextRecord = True
            time.sleep(self.sleepDurationLong)
            self.nav.navigateToClassID("MainContent_wizTasks_ctl17_btnCancelStart")

    def handleDefer(self, OrganisationID):
        self.nav.navigateToClassID("MainContent_gvImportGroup_chkDefer_0")
        time.sleep(self.sleepDuration)
        self.nav.navigateToClassID("btnDefer")
        time.sleep(self.sleepDuration)
        # there's also an alert dialog "Are you sure you want to defer all the selected manual amendments?"
        # with options Yes or Cancel - I so alert.accept() is probably the way forward
        self.driver.switch_to.alert.accept()

    def handleAudit(self):
        self.nav.enterTextToClassID("MainContent_AuditReason1_txtAuditReason", self.auditText)
        time.sleep(self.sleepDuration)
        # Click the Audit reason Button
        self.nav.navigateToClassID("MainContent_AuditReason1_btnAuditReasonOk")
        time.sleep(self.sleepDuration)

if __name__ == "__main__":
    args = sys.argv
    if len(args) == 2:
        configFile = sys.argv[1]
        MA = ManAmend(configFile)
        MA.process()
    else:
        msg = "Please provide the following arguments:\n"
        msg+= "config file name\n\n"
        print( msg )

