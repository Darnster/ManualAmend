from selenium.webdriver import Chrome
from selenium.common.exceptions import *

import time, sys

"""
Quick POC for Clearing Chrome settings
"""

class SelTest:

    def __init__(self):
        """
        """

        self.driver = Chrome( executable_path='C:\\selenium\ChromeDriver.exe')
        self.sleepDuration = 1
        self.connectURL = ""
        self.navigateLimit = 3


    def connect(self, url):

        #test connect:
        self.connectURL = url
        self.driver.get(self.connectURL)
        print("Successfully connected to = %s" % self.connectURL)
        time.sleep(self.sleepDuration)


    def refresh(self):
        self.driver.refresh()
        print("Successfully refreshed %s" % self.connectURL)
        time.sleep(self.sleepDuration)

    def clearAllCookies(self):
        self.driver.get("chrome://settings/siteData")
        if self.navigateByCSS('* /deep/ #removeShowingSites'):
            print("Successfully deleted all cookies")
            time.sleep(self.sleepDuration * 30)
        else:
            print("Unable to delete all cookies")
            sys.exit()

    def navigateByCSS(self, CSSRef):
        """
        Navigates and clicks the element supplied in ID
        :param ID: String = ID of widget to be clicked
        :return: Boolean
        """
        result = False
        attempts = 0
        while attempts < self.navigateLimit:
            try:
                self.target = self.driver.find_element_by_css_selector(CSSRef)
                self.target.click()
                result = True
                break
            except StaleElementReferenceException:
                attempts += 1
                time.sleep(1)
            except NoSuchElementException:
                attempts += 1
                time.sleep(1)
        return result



if __name__ == "__main__":
    poc = SelTest()
    poc.connect("https://corp%5cdaru:$Spring2019@oscar-testsp")
    poc.refresh()
    poc.clearAllCookies()
    # should present login dialog
    poc.connect("https://oscar-testsp")


