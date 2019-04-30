import win32com.client
from selenium.webdriver import Firefox
from selenium.common.exceptions import *
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
import time, datetime, sys
from ProcessError import ProcessingError
from Navigation import Navigation

"""
Title: OSCAR Manual Amendments Script
Compatibility: Python 3.6 or later
Status: Draft
Author: Danny Ruttle
Version Date: 2019-04-17
Project: ODS 3rd PArty data Automation
Internal Ref: v0.6
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
0.6 - simplified logging and mmoved creation of the log file to __init__, refactored record processing
      so new methods handleDefer and handle Audit added

TO DO:

1. Add logic to prevent infinite call to process() from Webdriver exception - done in v0.6 but not tested
2. Handle records that display errors (as in the broken record detector) - if required, to be determined by testing
3. Switch to using a config file for running the application
4. Add a crypto type method for storing domain passwords



"""

class ManAmend:

    def __init__(self, env, user, pwd, AmendID, AuditText, limit):
        """

        :param env: OSCAR environment
        :param user: user network name (domain not reauired)
        :param pwd: network password
        :param AmendID: import group to be processed
        :param AuditText: Appropriate message
        :param limit: allows batching of records to be controlled
        """
        # process args here
        self.env = env
        self.user = user
        self.pwd = pwd
        self.amendID = AmendID
        self.auditText = AuditText
        self.processLimit = limit
        self.fileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames

        self.binary = FirefoxBinary('C:\\Program Files\\Mozilla Firefox\\firefox.exe')
        self.sleepDuration = 1
        self.sleepDurationLong = 2
        self.processLimit = 0
        self.navigateLimit = 3 #used to set loops when searching for elements/IDs

        # logging vars
        self.logFileName = "ScriptLog_ManualAmends_%s_%s_%s.csv" % (self.env, self.amendID, self.fileTime)
        self.logFile = open(self.logFileName, "w")



    def process(self):
        """
        """

        self.driver = Firefox(firefox_binary=self.binary, executable_path='C:\\selenium\geckodriver.exe')

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

        self.logFile.write("processing started: %s\n" % self.fileTime)

        # connect to the environment (not part of the test) to prompt authentication
        amendCount = 0
        self.processState = True
        try:
            while amendCount < self.processLimit and self.processState == True:
                self.processAmendment()
                amendCount += 1
        except WebDriverException:
            self.restartDriver("WebDriverException in process() > Loop")

        # CLOSE THE SCRIPT LOGFILE:
        self.endFileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.logFile.write("processing completed: %s\n" % self.endFileTime)
        self.logFile.close()
        self.driver.quit()

    def restartDriver(self, outputMessage):
        print(outputMessage)
        time.sleep(self.sleepDurationLong)
        self.driver.quit()
        outputMessage = "Quiting driver due to %s" % outputMessage
        print(outputMessage)
        outputMessage = "Driver disposed now calling process() due to %s" % outputMessage
        print(outputMessage)
        time.sleep(self.sleepDurationLong)
        self.process()

    def processAmendment(self):
        """
        """
        OrganisationID = 0 # Org being processed - required for audit
        AmendmentsURL = "https://%s/Exchange/ImportGroupManualChanges.aspx?ID=%s" % (self.env, self.amendID)
        print("Import Group to be processed = %s" % AmendmentsURL)

        # if we can't get the organisationID then exit
        try:
            self.driver.get(AmendmentsURL)
        except UnexpectedAlertPresentException:
            self.handleAuth()
            self.driver.get(AmendmentsURL)
        except WebDriverException:
            outputMessage = "WebDriverException in processAmendment()"
            self.restartDriver(outputMessage)

        time.sleep(self.sleepDurationLong)

        try:
            OrganisationID = self.nav.getOrganisationIDfromProcessLink("MainContent_gvImportGroup_hlProcess_0")
        except:
            msg = '"System Exit!","Unable to retrieve OrganisationID"\n'
            self.logFile.write(msg)
            self.procesState = False

        if self.nav.checkNotes("MainContent_gvImportGroup_aNotes_0"):
            # too complex to automate so just defer it
            self.handleDefer()
            msg = '"%s","Defer","Processing Notes Found\n"' % OrganisationID
            self.logFile.write(msg)

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
                    Need to detect if there's an error on the page and by pass for now
                    """
                    procErr = ProcessingError(OrganisationID, self.driver, self.logFile)
                    if procErr.hasErrors(): # may need to put a conditional here
                        procErr.handleProcessingError()
                    # Audit panel and buttons have the same ID
                    self.handleAudit()

                    try:
                        self.handleClose()
                        msg = '"%s","ProcessClose","Closure process completed"\n' % OrganisationID
                        print( msg )
                        self.logFile.write(msg)
                        time.sleep( self.sleepDuration )
                    except:
                        # assume nothing to do here
                        pass

            else:
                msg = "End of records reached"
                self.logFile.write(msg)
                time.sleep( self.sleepDuration )
                self.procesState = False
        return self.processState

    def handleAuth(self):
        """
        Simulates what AutoIt did previously
        :return:
        """
        time.sleep(self.sleepDuration)
        print("authentication process called")
        shell = win32com.client.Dispatch("WScript.Shell")
        shell.Sendkeys("corp\\%s" % self.user)
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
    if len(args) == 7:
        env = sys.argv[1]
        user = sys.argv[2]
        pwd = sys.argv[3]
        amendID = sys.argv[4]
        auditText = sys.argv[5] # "CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE"
        recordsToProces = sys.argv[6]
        MA = ManAmend(env, user, pwd, amendID, auditText, int(recordsToProces))
        MA.process()
    else:
        msg = "Please provide the following arguments:\n"
        msg+= "1. environment, e.g. OSCAR-TESTSP\n"
        msg+= "2. user (shortcode only, Domain not required)\n"
        msg+= "3. password\n"
        msg+= "4. amendID (The Import to be procesed - e.g. 627)\n"
        msg+= "5. auditText - e.g.CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE\n"
        msg+= "6. Number of recordsToProces - e.g. 5\n\n"
        print( msg )

