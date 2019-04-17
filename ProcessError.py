import time
from Navigation import Navigation

class ProcessingError:
    def __init__(self, OrganisationID, driver, logfile):
        self.driver = driver
        self.logFile = logfile
        self.OrganisationID = OrganisationID
        self.errList = []
        self.sleepDuration = 2
        self.nav = Navigation(self.driver)

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
                    outputString = '%s,%s"\n' % (self.OrganisationID, item.text)
                    self.errList.append(item.text) # don't require ID for processing
                    self.logFile.write(outputString)
                    # print errors to console
                    print(outputString)

        return self.errList

    def getFixableErrors(self):
        """
        Provide a list of strings of containing the types of errors that can be fixed by this process
        :return:list
        """
        return ["Organisation Short Name required when organisation name longer than 40 characters"]

    def handleProcessingError(self):
        '''
        This method is here to identify errors and take appropriate steps - eg. short name error
        '''
        for err in self.errList:
            if err in self.getFixableErrors():
                if err == 'Organisation Short Name required when organisation name longer than 40 characters':
                    if self.addShortName():
                        self.nav.navigateToClassID("MainContent_btnSave")
                    return True
            else:
                if self.nav.navigateToClassID("MainContent_btnDefer"):
                    msg = '"%s","Defer Button clicked after discovering complex error on Org Maintenance Screen"\n' % self.OrganisationID
                    print(msg)
                    self.logFile.write(msg)
                    time.sleep( self.sleepDuration )
                return True

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



