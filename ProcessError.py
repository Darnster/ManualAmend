import time
class ProcessingError:
    def __init__(self, OrganisationID, driver, logfile):
        self.errors = ""  # set as a string here but will be cast to an object later
        self.driver = driver
        self.logFile = logfile
        self.OrganisationID = OrganisationID
        self.errList = []

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
                    self.errList.append(outputString)
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
        for err in self.errors:
            if err in self.getFixableErrors():
                if err == 'Organisation Short Name required when organisation name longer than 40 characters':
                    self.addShortName()
                    return True
            else:
                if self.navigateToClassID("MainContent_btnDefer"):
                    msg = '"%s","Defer Button clicked after discovering complex error on Org Maintenance Screen"\n' % OrganisationID
                    print(msg)
                self.logFile.write(msg)
                time.sleep(1)
                return True

    def addShortName(self):

