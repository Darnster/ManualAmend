from selenium.common.exceptions import *
import time

"""
Abstracts search, click, element text content and entering of text into fields.
"""

class Navigation:

    def __init__(self, driver):
        self.driver = driver
        self.navigateLimit = 3
        self.sleepDurationLong = 2


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
                time.sleep( self.sleepDurationLong )
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
                time.sleep( self.sleepDurationLong )
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
                time.sleep( self.sleepDurationLong )
        return result

    def enterTextToClassID(self, ID, text):
        """
        :param ID: String = ID of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id( ID )
                self.target.send_keys(text)
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep( self.sleepDurationLong )
            except ElementNotVisibleException:
                attempts += 1
                time.sleep( self.sleepDurationLong )
            except WebDriverException:
                attempts += 1
                time.sleep(self.sleepDurationLong )
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
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id(ID)
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep( self.sleepDurationLong )
            except NoSuchElementException:
                attempts += 1
                time.sleep( self.sleepDurationLong )
                result = False
                break
        return result

    def getOrganisationIDfromProcessLink(self, ID):
        """
        :param ID: this is the ID of the element
        :return: parsed OrganisationID
        """
        attempts = 0
        OrganisationID = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id( ID )
                linkText = self.target.get_attribute("href")
                OrganisationID = linkText.split("=")[1]
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep( self.sleepDurationLong )
        return OrganisationID


    def getTextbyID(self, ID):
        """
        :param ID: this is the ID of the element
        :return: parsed OrganisationID
        """
        attempts = 0
        text = ""
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_id( ID )
                text = self.target.get_attribute("value")
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep( self.sleepDurationLong )
        return text