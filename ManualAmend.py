import win32com.client
from selenium import webdriver
from selenium.webdriver import Firefox
from selenium.webdriver import FirefoxOptions
from selenium.common.exceptions import *
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary

import subprocess


import time, datetime, sys
import csv


"""
Title: OSCAR Manual Amendments Script
Compatibility: Python 3.6 or later
Status: Draft
Author: Danny Ruttle
Version Date: 2019-04-04
Project: ODS 3rd PArty data Automation
Internal Ref: v0.3
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
        d. searches for a button on the target page with the ID "btnSave"
        e. btnSave is then clicked button which then returns the process back to the original URL
4. A log is provided to audit what the script did and report any errors


CHANGES:

0.1 - initial POC with half a dozen Social HQ/Providers
0.2 - modified to action the defer on the same screen
0.3 - version checked in ***without*** TOO MANY REDIRECTS issue fixed

TO DO:

1. Handle records that display errors (as in the broken record detector) - if required, to be determined by testing
2. Truncate short name if that error appears.
3. Identify "TOO MANY REDIRECTS" message and bypass
- search for ID/text
- refresh and search until not present
- return control back to the processing script

NEXT THING TO DO:

Create a new browser instance on each call! - NOPE
Identify when the 


"""

class ManAmend:

    def __init__(self):
        """
        """
        self.fileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames

        binary = FirefoxBinary('C:\\Program Files\\Mozilla Firefox\\firefox.exe')
        self.driver = Firefox(firefox_binary=binary, executable_path='C:\\selenium\geckodriver.exe')
        self.auditText = ""
        self.sleepDuration = 1
        self.processLimit = 0
        self.refreshLimit = 50
        self.navigateLimit = 3


    def process(self, env, user, pwd, AmendID, AuditText, limit):
        """
        :param env:
        :param user:
        :param pwd:
        """

        self.env = env
        self.user = user
        self.pwd = pwd
        self.amendID = AmendID
        self.auditText = AuditText
        self.processLimit = limit

        #test connect:
        connectURL = "https://%s" % ( self.env )
        print("Connecting to %s......" % connectURL)
        try:
            self.driver.get(connectURL)
            print("Successfully connected to %s......" % connectURL)
        except UnexpectedAlertPresentException:
            print("need to authenticate when accessing home")
            print("### self.user = %s ###" % self.user)
            time.sleep(self.sleepDuration)
            print("need to authenticate when accessing amendments page")
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.Sendkeys("corp\\%s" % self.user)
            time.sleep(3)
            shell.Sendkeys("{TAB}")
            time.sleep(3)
            shell.Sendkeys(self.pwd)
            time.sleep(3)
            shell.Sendkeys("{ENTER}")
            time.sleep(3)
        time.sleep(self.sleepDuration) #allow time to login manually

        # logging vars
        self.logFileName = "ScriptLog_ManualAmends_%s_%s_%s.csv" % (self.env, self.amendID, self.fileTime)
        self.logFile = open(self.logFileName, "w")
        self.logFile.write("processing started: %s\n" % self.fileTime)

        # connect to the environment (not part of the test) to prompt authentication
        amendCount = 0
        while amendCount < self.processLimit:
            #if self.detectTooManyRedirects():
            self.processAmendment()
            self.driver.delete_all_cookies()

            amendCount += 1

        # CLOSE THE SCRIPT LOGFILE:
        self.endFileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.logFile.write("processing completed: %s\n" % self.endFileTime)
        self.logFile.close()

    def processAmendment(self):
        """
        """
        OrganisationID = 0 # Org being processed - required for audit

        #AmendmentsURL = "https://%s:%s@%s/Exchange/ImportGroupManualChanges.aspx?ID=%s" % (self.user, self.pwd, self.env, self.amendID)
        AmendmentsURL = "https://%s/Exchange/ImportGroupManualChanges.aspx?ID=%s" % (self.env, self.amendID)

        print("Import Group to be processed = %s" % AmendmentsURL)

        # if we can't get the organisationID then exit
        try:
            self.driver.get(AmendmentsURL)
        except UnexpectedAlertPresentException:
            print("### self.user = %s ###" % self.user)
            time.sleep(self.sleepDuration)
            print("need to authenticate when accessing amendments page")
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.Sendkeys("corp\\%s" % self.user)
            time.sleep(3)
            shell.Sendkeys("{TAB}")
            time.sleep(3)
            shell.Sendkeys(self.pwd)
            time.sleep(3)
            shell.Sendkeys("{ENTER}")
            time.sleep(3)
            self.driver.get(AmendmentsURL)

            #alert = self.driver.switch_to.window(self.driver.window_handles[1])
            #alert = self.driver.switch_to.alert()
            #time.sleep(self.sleepDuration)
            #alert.send_keys("corp\\%s" % self.user)
            #alert.send_keys(Keys.TAB)
            #alert.send_keys("%s" % self.pwd)
            #alert.accept()

        time.sleep(self.sleepDuration * 2)
        #self.detectTooManyRedirects()


        try:
            OrganisationID = self.getOrganisationIDfromProcessLink("MainContent_gvImportGroup_hlProcess_0")
        except:
            msg = '"System Exit!","Unable to retrieve OrganisationID"\n'
            self.logFile.write(msg)
            #sys.exit()


        if self.checkNotes("MainContent_gvImportGroup_aNotes_0"):
            msg = '"%s","Processing Notes Found","processing deferred"\n' % OrganisationID
            print(msg)
            self.logFile.write(msg)
            time.sleep(self.sleepDuration)
            if self.navigateToClassID("MainContent_gvImportGroup_chkDefer_0"):
                msg = '"%s","Defer checkbox clicked and marked to be deferred"\n' % OrganisationID
                print(msg)
                self.logFile.write(msg)
                time.sleep(self.sleepDuration)
                # for next step selenium claims that ctl00$MainContent$btnDefer is a list so changed to ID and btnDefer
                if self.navigateToClassID( "btnDefer" ):
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
            if self.navigateToClassID("MainContent_gvImportGroup_hlProcess_0"):
                msg = '"%s","navigate to record processed successfully"\n' % OrganisationID
                print( msg )
                self.logFile.write(msg)
                time.sleep( self.sleepDuration )
                try:
                    # handle save alert
                    self.driver.switch_to.alert.dismiss()
                except:
                    pass # it nerer happened!


                if self.navigateToClassID("MainContent_btnSave"):
                    msg = '"%s","accepting amendment processed successfully"\n' % OrganisationID
                    print(msg)
                    self.logFile.write(msg)
                    time.sleep(self.sleepDuration * 2)
                    """
                    Need to detect if there's an error on the page and by pass for now
                    """
                    if self._handlePrcoessingError(): # may need to put a conditional here
                        # abort and log
                        msg = '"%s","unable to save record as Errors exist."\n' % OrganisationID
                        print(msg)
                        self.logFile.write(msg)
                    else:
                        if self.enterTextToClassID("MainContent_AuditReason1_txtAuditReason"):
                            msg = '"%s","audit text entry entered successfully"\n' % OrganisationID
                            print( msg )
                            self.logFile.write(msg)
                            time.sleep( self.sleepDuration )
                            # Click the Audit reason Button
                            if self.navigateToClassID("MainContent_AuditReason1_btnAuditReasonOk"):
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

    def detectTooManyRedirects(self):
        try:
            self.target = self.driver.find_element_by_class_name("error-code")
            if self.clearTooManyRedirects() == False:
                msg = '"Error"unable to clear TOO MANY REDIRECTS issue"\n'
                print(msg)
                self.logFile.write(msg)
                return False
                sys.exit()  # for now - may just log it - or quit and restart the driver!!!
            else:
                msg = '"Resolved","Successfully worked around TOO MANY REDIRECTS issue"\n'
                print(msg)
                self.logFile.write(msg)
                return True
        except NoSuchElementException:
            return True

    def clearTooManyRedirects(self, classToFind):
        """

        :return: boolean (True if the TOO MANY REDIRECTS page has gone)
        """

        result = False

        attempts = 0
        while attempts < self.refreshLimit:
            try:
                self.target = self.driver.find_element_by_class_name( classToFind )  # this is present in the error page in Chrome
                # maybe check the text if ambiguous results are seen
                self.driver.refresh()
                attempts += 1
                time.sleep(1)

            except NoSuchElementException:
                result = True
                break

        return result




    def navigateToClassID(self, ID):
        """
        Navigates and clicks the element supplied in ID
        :param ID: String = ID of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id( ID )
                self.target.click()
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
        return result

    def searchforClassID(self, ID):
        """
        Navigates and clicks the element supplied in ID
        :param ID: String = ID of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id( ID )
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
        return result


    def navigateByName(self, Name):
        """
        Navigates and clicks the element supplied in Name

        :param Name: String = Name of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_elements_by_name( Name )
                self.target.click()
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
        return result

    def enterTextToClassID(self, ID):
        """
        :param ID: String = ID of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id( ID )
                self.target.send_keys(self.auditText)
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
            except ElementNotVisibleException:
                attempts += 1
                time.sleep(1)
        return result

    def checkNotes(self, ID):   # ID = MainContent_gvImportGroup_aNotes_0
        """
        Class only appears if Notes are present
        Notes are only present if it's not a straightforward click
        So need to look for this class and:
        if present: Defer
        else: Accept
        """

        result = False
        attempts = 0
        while attempts < 3:
            try:
                self.target = self.driver.find_element_by_id(ID)
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
            except NoSuchElementException:
                attempts += 1
                time.sleep(1)
                result = False
                break
        return result

    def getOrganisationIDfromProcessLink(self, ID):
        attempts = 0
        organisationID = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id( ID )
                linkText = self.target.get_attribute("href")
                OrganisationID = linkText.split("=")[1]
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
        return OrganisationID



# not implemented yet!!!!


    def _handlePrcoessingError(self):
        '''
        This method is here to identify errors and take appropriate steps - eg. short name error
        :return:  boolean
        '''
        # simple binary return...
        return False # stubbed for now
        """
        try:
            errors = self.driver.find_element_by_id("MainContent_CustomValidationSummary1")
            return True
        except:
            return False
        """
#        uls = errors.find_elements_by_tag_name("ul")
#        if len(uls) > 0:
#            for ul in uls:
#                li = ul.find_elements_by_tag_name("li")
#                for item in li:
#                    outputString = '"%s","\n' % (item.text)
#                    # self.outputFile.write(outputString)
#                    # print errors to console
#                    print(outputString)
#                    self.logFile.write("Error reported for row %s\n" % outputString)

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

