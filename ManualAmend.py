from selenium import webdriver
from selenium.common.exceptions import *
import subprocess

import time, datetime, sys
import csv


"""
Title: OSCAR Manual Amendments Script
Compatibility: Python 3.6 or later
Status: Draft
Author: Danny Ruttle
Version Date: 2019-02-13
Project: ODS Reconfiguration
Internal Ref: v0.2
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
3. It searches for the element "MainContent_gvImportGroup_hlProcess_0" 
    a. clicks the link within that table cell e.g. href="ImportManualChange.aspx?ImportGroupOrganisationID=3955786"
    b. searches for a button on the target page with the ID "btnSave"
    c. Clicks the button which then returns the user back to the original URL
4. A log is provided to audit what the script did and report any errors


CHANGES:

0.1 - initial POC with half a dozen Social HQ/Providers

TO DO:

1. Handle that have been recorded as autoload failed - defer these and log that the action for the support team to follow up
2. Handle records that display errors (as in the broken record detector)


"""

class ManAmend:

    def __init__(self):
        """
        """
        self.fileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.driver = webdriver.Chrome('C:\\selenium\chromedriver.exe')
        self.auditText = ""
        self.sleepDuration = 2


    def connect(self, env, user, pwd):
        """
        :param env:
        :param user:
        :param pwd:
        """

        self.env = env
        # connect to the environment (not part of the test) to prompt authentication
        self.driver.get('https://%s:%s@%s/' % ( user, pwd, self.env ) )


    def processAmendments(self, AmendID, AuditText):
        """
        :param file:
        :param driver:
        :return:
        """
        self.logFileName = "ScriptLog_ManualAmends_%s_%s_%s.csv" % ( self.env, AmendID, self.fileTime )
        self.logFile = open(self.logFileName, "w")
        self.logFile.write("processing started: %s\n" % self.fileTime )

        self.auditText = AuditText
        testCount = 0
        endOfRecords = False
        while not endOfRecords:
            if testCount == 3:
                sys.exit()
            # flag to decide whether to Accept or Defer the change
            defer = False
            # work around for too many redirects issue - browse to X09:
            X09Org = "/OrganisationScreens/OrganisationMaintenance.aspx?ID=111933"
            url = "https://%s/%s" % (self.env, X09Org)
            self.driver.get(url)
            print("naviagated to X09 = %s" % url)
            time.sleep( self.sleepDuration )

            url = "https://%s/Exchange/ImportGroupManualChanges.aspx?ID=%s" % (self.env, AmendID)
            print("Import Group = %s" % url)
            self.driver.get(url)
            time.sleep( self.sleepDuration )

            try:
                # first step is to get the the OSCAR organisation ID by returning the text from the element below and assigning this to self.target
                # <input type="hidden" name="ctl00$MainContent$gvImportGroup$ctl04$hidIGOID" id="MainContent_gvImportGroup_hidIGOID_1" value="3955943">
                if self.checkNotes("MainContent_gvImportGroup_aNotes_0"):
                    msg = '"%s","Processing Notes Found","processing deferred"\n' % self.target
                    print(msg)
                    time.sleep(self.sleepDuration)
                    if self.navigateToClassID("MainContent_gvImportGroup_chkDefer_0"):
                        msg = '"%s","Defer checkbox clicked","Record marked to be deferred"\n' % self.target
                        print(msg)
                        time.sleep(self.sleepDuration)
                        if self.navigateByName( "ctl00$MainContent$btnDefer" ):
                            msg = '"%s","Defer button clicked","Record successfully deferred"\n' % self.target
                            print(msg)
                            time.sleep(self.sleepDuration)
                else:
                    if self.navigateToClassID("MainContent_gvImportGroup_hlProcess_0"):
                        msg = '"%s","navigate to record","processed successfully"\n' % self.target
                        print( msg )
                        time.sleep( self.sleepDuration )
                        if self.navigateToClassID("MainContent_btnSave"):
                            msg = '"%s","accepting amendment","processed successfully"\n' % self.target
                            print(msg)
                            if self.enterTextToClassID("MainContent_AuditReason1_txtAuditReason"):
                                msg = '"%s","audit text entry","processed successfully"\n' % self.target
                                print( msg )
                                time.sleep( self.sleepDuration )
                                # Click the Audit reason Button
                                if self.navigateToClassID("MainContent_AuditReason1_btnAuditReasonOk"):
                                    msg = '"%s","audit button press","processed successfully"\n' % self.target
                                    print( msg )
                                    time.sleep( self.sleepDuration )
                                else:
                                    msg = '"%s","audit button press","processing failed"\n' % self.target
                                    print( msg )
                                    time.sleep( self.sleepDuration )
                            else:
                                msg = '"%s","audit button press","processing failed"\n' % self.target
                                print( msg )
                                time.sleep( self.sleepDuration )
                        else:
                            msg = '"%s","audit text entry","processing failed"\n' % self.target
                            print( msg )
                            time.sleep( self.sleepDuration )
                    else:
                        msg = '"%s","navigate to record","processing failed"\n' % self.target
                        print( msg )
                        time.sleep( self.sleepDuration )
            except NoSuchElementException as e:
                endOfRecords = True
                e.msg
                msg = "End of records reached (hopefully via NoSuchElementException MainContent_gvImportGroup_hlProcess_0) - see below:\n%s\n" % e.msg
                print( msg )
                time.sleep( self.sleepDuration )
            testCount += 1
            self.logFile.write( msg )


        # CLOSE THE SCRIPT LOGFILE:
        self.endFileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.logFile.write("processing completed: %s\n" % self.endFileTime)
        self.logFile.close()
        # quit the driver
        self.driver.quit()

    def navigateToClassID(self, ID):
        """
        :param ID: String = ID of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < 3:
            try:
                self.target = self.driver.find_element_by_id( ID )
                self.target.click()
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
        return result

    def navigateByName(self, Name):
        """
        :param Name: String = Name of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < 3:
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
        while attempts < 3:
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



if __name__ == "__main__":
    MA = ManAmend()
    MA.connect("oscar-testsp", "corp%5cdaru", "$Spring2019")
    MA.processAmendments("627", "CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE")


"""
selenium.common.exceptions.UnexpectedAlertPresentException: Alert Text: None
Message: unexpected alert open: {Alert text : This record is currently locked by CORPdaru as of 04/02/2019 14:37}
"""