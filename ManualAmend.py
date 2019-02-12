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
Version Date: 2019-02-11
Project: ODS Reconfiguration
Internal Ref: v0.1
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

0.1 - initial POC

TO DO:

1. Handle Pop-ups - Done for "Laeve saved record" alert
2. Deal with Locked Records - need navigate to that part of the system and release
3. CSV to ignore header row - Done
4. Handle alerts - bypassbut collect the error message/notification
#  selenium.common.exceptions.UnexpectedAlertPresentException: Alert Text: None


"""

class ManAmend:

    def __init__(self):
        """
        """
        self.fileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.logFileName = "ScriptLog_ManualAmends%s.csv" % self.fileTime
        self.driver = webdriver.Chrome('C:\\selenium\chromedriver.exe')
        self.auditText = ""


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

        #self.logFile = open(self.logFileName, "w")
        #self.logFile.write("processing started: %s\n" % self.fileTime )

        self.auditText = AuditText

        endOfRecords = False
        while not endOfRecords:
            # work around for too many redirects issue - browse to X09:
            X09Org = "/OrganisationScreens/OrganisationMaintenance.aspx?ID=111933"
            url = "https://%s/%s" % (self.env, X09Org)
            self.driver.get(url)
            print("naviagated to X09 = %s" % url)
            time.sleep(2)

            url = "https://%s/%s" % (self.env, AmendID)
            print("Import Group = %s" % url)
            self.driver.get(url)
            time.sleep(2)

            try:
                if self.navigateToClassID("MainContent_gvImportGroup_hlProcess_0"):
                    msg = '"%s","navigate to record","processed successfully"' % self.target
                    print( msg )
                    time.sleep(2)
                    if self.navigateToClassID("MainContent_btnSave"):
                        msg = '"%s","accepting amendment","processed successfully"' % self.target
                        print(msg)
                        if self.enterTextToClassID("MainContent_AuditReason1_txtAuditReason"):
                            msg = '"%s","audit text entry","processed successfully"' % self.target
                            print(msg)
                            time.sleep(2)
                            # Click the Audit reason Button
                            if self.navigateToClassID("MainContent_AuditReason1_btnAuditReasonOk"):
                                msg = '"%s","audit button press","processed successfully"' % self.target
                                print(msg)
                                time.sleep(2)
                            else:
                                msg = '"%s","audit button press","processing failed"' % self.target
                                print(msg)
                        else:
                            msg = '"%s","audit button press","processing failed"' % self.target
                            print(msg)
                    else:
                        msg = '"%s","audit text entry","processing failed"' % self.target
                        print( msg )
                else:
                    msg = '"%s","navigate to record","processing failed"' % self.target
            except NoSuchElementException:
                endOfRecords = True


        # CLOSE THE SCRIPT LOGFILE:
        #self.endFileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        #self.logFile.write("processing completed: %s\n" % self.endFileTime)
        #self.logFile.close()
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

if __name__ == "__main__":
    MA = ManAmend()
    MA.connect("oscar-testsp", "corp%5cdaru", "$Spring2019")
    MA.processAmendments("Exchange/ImportGroupManualChanges.aspx?ID=623", "CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE")


"""
selenium.common.exceptions.UnexpectedAlertPresentException: Alert Text: None
Message: unexpected alert open: {Alert text : This record is currently locked by CORPdaru as of 04/02/2019 14:37}
"""