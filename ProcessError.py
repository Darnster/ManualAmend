import time
from Navigation import Navigation
import WriteLog

"""
Changes

amdndCount now included in logs

"""

class ProcessingError:
    def __init__(self, OrganisationID, driver, logfile, logFileName, amendCount):
        self.driver = driver
        self.logFile = logfile
        self.logFileName = logFileName
        self.OrganisationID = OrganisationID
        self.errList = []
        self.sleepDuration = 2
        self.nav = Navigation(self.driver)
        self.amendCount = amendCount

    def hasErrors(self):
        """
        Confirms whether the error panel is present
        :return: boolean
        """
        try:
            self.errors = self.driver.find_element_by_id("MainContent_CustomValidationSummary1")
            if len(self.getErrors()) > 0:
                return True
            else:
                return False
        except:
            return False

    def getErrors(self):
        """
        Parses all errors reported
        :return: list
        """

        uls = self.errors.find_elements_by_tag_name("ul")
        if len(uls) > 0:
            for ul in uls:
                li = ul.find_elements_by_tag_name("li")
                for item in li:
                    outputString = '"%s","%s","Error to Correct","%s"\n' % ( self.amendCount, self.OrganisationID, item.text)
                    self.errList.append(item.text) # don't require ID for processing
                    self.WriteLog( outputString )
                    # print errors to console

        return self.errList

    def getFixableErrors(self):
        """
        Provide a list of strings of containing the types of errors that can be fixed by this process
        :return:list
        """
        return ["Organisation Short Name required when organisation name longer than 40 characters",
                "Organisation Short Name must be absent unless organisation name longer than 40 characters"]

    def handleProcessingError(self):
        '''
        This method is here to identify errors and take appropriate steps - eg. short name error
        '''
        for err in self.errList:
            if err in self.getFixableErrors():
                if err == 'Organisation Short Name required when organisation name longer than 40 characters':
                    self.addShortName()
                    self.nav.navigateToClassID("MainContent_btnSave")
                    msg = '"%s","%s","ProcessErrorFixed","Complex error on Org Maintenance Screen(%s)"\n' % ( self.amendCount, self.OrganisationID, err )
                    self.WriteLog( msg )
                    time.sleep(self.sleepDuration)
                else:
                    self.removeShortName()
                    self.nav.navigateToClassID("MainContent_btnSave")
                    msg = '"%s","%s","ProcessErrorFixed","Complex error on Org Maintenance Screen(%s)"\n' % (
                    self.amendCount, self.OrganisationID, err)
                    self.WriteLog(msg)
                    time.sleep(self.sleepDuration)
            else:
                if self.nav.navigateToClassID("MainContent_btnDefer"):
                    msg = '"%s","%s","ProcessErrorDeferred","Complex error on Org Maintenance Screen(%s)"\n' % ( self.amendCount, self.OrganisationID, err )
                    self.WriteLog( msg )
                    time.sleep( self.sleepDuration )

    def addShortName(self):
        orgNameID = "MainContent_tcDetails_tpOrganisation_ucOrganisation_txtName"
        orgShortNameID = "MainContent_tcDetails_tpOrganisation_ucOrganisation_txtShortName"
        """
        get orgName and truncate to 40 chars
        set orgShortName to shortened value
        """
        try:
            nav = Navigation(self.driver)
            orgName = nav.getTextbyID(orgNameID)
            time.sleep( self.sleepDuration )
            orgShortName = orgName[0:40]
            nav.enterTextToClassID(orgShortNameID, orgShortName)
            return True
        except:
            return False

    def removeShortName(self):
        orgShortNameID = "MainContent_tcDetails_tpOrganisation_ucOrganisation_txtShortName"
        """
        get orgName and truncate to 40 chars
        set orgShortName to shortened value
        """
        try:
            nav = Navigation(self.driver)
            time.sleep( self.sleepDuration )
            orgShortName = ""
            nav.enterTextToClassID(orgShortNameID, orgShortName)
            return True
        except:
            return False

    def WriteLog(self, msg):
        """
        Opens up the logfile appends an updated and closes the file
        This behaviour allows the progress to be reviewed during processing
        :param msg: string to be written
        :return: none
        """
        wl = WriteLog
        wl = WriteLog.WL()
        wl.WriteLog(self.logFileName, msg)
