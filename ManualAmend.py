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
Version Date: 2019-04-15
Project: ODS 3rd PArty data Automation
Internal Ref: v0.4
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
2. It authenticates using the Selenium chrome Web Driver
3. It always acts on the record at the top of the list, processing that and then refreshing the page containing manual amendments
4. It searches for the element "MainContent_gvImportGroup_aNotes_0" (this will be in the first entry only)
    if this element is present
        a. the checkbox is clicked to identify that this record should be deferred
        b. the defer button is pressed, which then refreshes the page and the next record is then at index "_0"
    else
        a. searches for the class "MainContent_gvImportGroup_hlProcess_0"
        b. clicks the link within that table cell e.g. href="ImportManualChange.aspx?ImportGroupOrganisationID=3955786"
        c. calls enterTextToClassID() which searches for the audit field and enters the audit text
        d. 
        searches for a button on the target page with the ID "btnSave"
        e. btnSave is then clicked button which then returns the process back to the original URL
4. A log is provided to audit what the script did and report any errors


CHANGES:

0.1 - initial POC with half a dozen Social HQ/Providers
0.2 - modified to action the defer on the same screen
0.3 - version checked in ***without*** TOO MANY REDIRECTS issue fixed
0.4 - Firefox version with first stab at removing too many redirects handlers

TO DO:

1. Handle records that display errors (as in the broken record detector) - if required, to be determined by testing
2. Truncate short name if that error appears.
3. Identify "TOO MANY REDIRECTS" message and bypass
- search for ID/text
- refresh and search until not present
- return control back to the processing script

NEXT THING TO DO:



"""

class ManAmend:

    def __init__(self):
        """
        """
        self.fileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames

        binary = FirefoxBinary('C:\\Program Files\\Mozilla Firefox\\firefox.exe')
        self.driver = Firefox(firefox_binary=binary, executable_path='C:\\selenium\geckodriver.exe')
        cap = self.driver.capabilities
        self.auditText = ""
        self.sleepDuration = 1
        self.sleepDurationLong = 2
        self.processLimit = 0
        self.navigateLimit = 3 #used to set loops when searching for elements/IDs


    def process(self, env, user, pwd, AmendID, AuditText, limit):
        """

        :param env:
        :param user:
        :param pwd:
        :param AmendID:
        :param AuditText:
        :param limit:
        :return:
        """
        self.env = env
        self.user = user
        self.pwd = pwd
        self.amendID = AmendID
        self.auditText = AuditText
        self.processLimit = limit

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

        # logging vars
        self.logFileName = "ScriptLog_ManualAmends_%s_%s_%s.csv" % (self.env, self.amendID, self.fileTime)
        self.logFile = open(self.logFileName, "w")
        self.logFile.write("processing started: %s\n" % self.fileTime)

        # connect to the environment (not part of the test) to prompt authentication
        amendCount = 0
        self.processState = True
        while amendCount < self.processLimit and self.processState == True:
            self.processAmendment()
            amendCount += 1

        # CLOSE THE SCRIPT LOGFILE:
        self.endFileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.logFile.write("processing completed: %s\n" % self.endFileTime)
        self.logFile.close()
        self.driver.quit()


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
        time.sleep(self.sleepDurationLong)

        try:
            OrganisationID = self.nav.getOrganisationIDfromProcessLink("MainContent_gvImportGroup_hlProcess_0")
        except:
            msg = '"System Exit!","Unable to retrieve OrganisationID"\n'
            self.logFile.write(msg)
            self.procesState = False

        if self.nav.checkNotes("MainContent_gvImportGroup_aNotes_0"):
            msg = '"%s","Processing Notes Found","processing deferred"\n' % OrganisationID
            print(msg)
            self.logFile.write(msg)
            time.sleep(self.sleepDuration)
            if self.nav.navigateToClassID("MainContent_gvImportGroup_chkDefer_0"):
                msg = '"%s","Defer checkbox clicked and marked to be deferred"\n' % OrganisationID
                print(msg)
                self.logFile.write(msg)
                time.sleep(self.sleepDuration)
                # for next step selenium claims that ctl00$MainContent$btnDefer is a list so changed to ID and btnDefer
                if self.nav.navigateToClassID( "btnDefer" ):
                    msg = '"%s","Defer button clicked and successfully selected for confirmation to be deferred"\n' % OrganisationID
                    print(msg)
                    self.logFile.write(msg)
                    time.sleep(self.sleepDuration)
                    # there's also an alert dialog "Are you sure you want to defer all the selected manual amendments?"
                    # with options Yes or Cancel - I so alert.accept() is probably the way forward
                    self.driver.switch_to.alert.accept()
                    msg = "%s, deferred alert box successfully accepted\n" % OrganisationID
                    print(msg)
                    self.logFile.write(msg)

        else:
            if self.nav.navigateToClassID("MainContent_gvImportGroup_hlProcess_0"):
                msg = '"%s","navigate to record processed successfully"\n' % OrganisationID
                print( msg )
                self.logFile.write(msg)
                time.sleep( self.sleepDuration )
                try:
                    # handle save alert ### need to see more examples of these so sceanrios can be tested ###
                    self.driver.switch_to.alert.dismiss()
                except:
                    pass # it never happened!


                if self.nav.navigateToClassID("MainContent_btnSave"):
                    msg = '"%s","accepting amendment processed successfully"\n' % OrganisationID
                    print(msg)
                    self.logFile.write(msg)
                    time.sleep(self.sleepDurationLong)
                    """
                    Need to detect if there's an error on the page and by pass for now
                    """
                    procErr = ProcessingError(OrganisationID, self.driver, self.logFile)
                    if procErr.hasErrors(): # may need to put a conditional here
                        procErr.handleProcessingError()
                        msg = '"%s","Errors resolved."\n' % OrganisationID
                        print(msg)
                        self.logFile.write(msg)
                    if self.nav.enterTextToClassID("MainContent_AuditReason1_txtAuditReason", self.auditText):
                        msg = '"%s","audit text entry entered successfully"\n' % OrganisationID
                        print( msg )
                        self.logFile.write(msg)
                        time.sleep( self.sleepDuration )
                        # Click the Audit reason Button
                        if self.nav.navigateToClassID("MainContent_AuditReason1_btnAuditReasonOk"):
                            msg = '"%s","audit button press processed successfully"\n' % OrganisationID
                            print( msg )
                            self.logFile.write(msg)
                            time.sleep( self.sleepDuration )
                        else:
                            msg = '"%s","audit button press failed"\n' % OrganisationID
                            print( msg )
                            self.logFile.write(msg)
                            time.sleep( self.sleepDuration )
                    else:
                        msg = '"%s","audit text reason entry failed"\n' % OrganisationID
                        print( msg )
                        self.logFile.write(msg)
                        time.sleep( self.sleepDuration )
                else:
                    msg = '"%s","audit text entry processing failed"\n' % OrganisationID
                    print( msg )
                    self.logFile.write(msg)
                    time.sleep( self.sleepDuration )
            else:
                msg = '"%s","navigate to record processing failed"\n' % OrganisationID
                print( msg )
                self.logFile.write(msg)
                time.sleep( self.sleepDuration )
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

if __name__ == "__main__":
    args = sys.argv
    if len(args) == 7:
        env = sys.argv[1]
        user = sys.argv[2]
        pwd = sys.argv[3]
        amendID = sys.argv[4]
        auditText = sys.argv[5] # "CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE"
        recordsToProces = sys.argv[6]
        MA = ManAmend()
        MA.process(env, user, pwd, amendID, auditText, int(recordsToProces))
    else:
        msg = "Please provide the following arguments:\n"
        msg+= "1. environment, e.g. OSCAR-TESTSP\n"
        msg+= "2. user (shortcode only, Domain not required)\n"
        msg+= "3. password\n"
        msg+= "4. amendID (The Import to be procesed - e.g. 627)\n"
        msg+= "5. auditText - e.g.CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE\n"
        msg+= "6. Number of recordsToProces - e.g. 5\n\n"
        print( msg )

