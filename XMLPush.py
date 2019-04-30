from ManualAmend import *
import csv

"""
This software will push orgs identified by OSCAR organisation ID to the next XML Partial file.

tested with X09 - OSCAR Org ID = 111933

TO DO
1. Handle records that don't exist!

"""

class xmlPush(ManAmend):
    def __init__(self, env, user, pwd, file, AuditText):
        self.env = env
        self.user = user
        self.pwd = pwd
        self.file = file
        self.auditText = AuditText
        self.fileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames

        self.binary = FirefoxBinary('C:\\Program Files\\Mozilla Firefox\\firefox.exe')
        self.sleepDuration = 1
        self.sleepDurationLong = 2
        self.processLimit = 0
        self.navigateLimit = 3  # used to set loops when searching for elements/IDs

        # input file
        try:
            self.inputFile = open(self.file, "r")
            print("Processing: %s" % self.inputFile)
        except:
            print("Could not locate the specified file: %s" % self.file)
            sys.exit()

        # logging vars
        self.logFileName = "ScriptLog_pushToXML_%s_%s.csv" % (self.env, self.fileTime)
        self.logFile = open(self.logFileName, "w")

    def process(self):

        self.driver = Firefox(firefox_binary=self.binary, executable_path='C:\\selenium\geckodriver.exe')

        # create instance of Navigation class here
        self.nav = Navigation(self.driver)

        # test connect:
        connectURL = "https://%s" % (self.env)
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

        time.sleep(self.sleepDuration)  # allow time to login manually

        self.logFile.write("processing started: %s\n" % self.fileTime)

        self.processed = []
        reader = csv.reader(self.inputFile)
        self.rowNum = 0
        for record in reader:
            if self.rowNum == 0:
                self.rowNum = 1
                pass  # skip header row
            else:
                # loop over content and processID (avoid any that have been processed already in case of driver restart)
                if self.rowNum not in self.processed:
                    ID = record[0]
                    try:
                        self.pushRecord(ID)
                    except WebDriverException as e:
                        # Attempt retry
                        time.sleep(self.sleepDurationLong)
                        print("Web Driver Exception raised - %s" % e.msg)
                        self.restartDriver("Web Driver Exception raised")
                        self.logFile.write("WebDriver restart due to %s...." % e.msg[0:50])
                self.processed.append(self.rowNum)
                self.rowNum += 1
                self.logFile.write("Successfully Processed record %s\n" % ID)


        # close the input file
        self.inputFile.close()
        # close the output file

        # CLOSE THE SCRIPT LOGFILE:
        self.endFileTime = datetime.datetime.now().strftime("%Y-%m-%dT%H%M%S%Z")  # can't include ":" in filenames
        self.logFile.write("processing completed: %s\n" % self.endFileTime)
        self.logFile.close()
        self.driver.quit()

    def pushRecord(self, ID):
        self.navigateToOrg(ID)
        """
        MainContent_btnTasks
        """
        self.nav.navigateToClassID("MainContent_btnTasks")
        time.sleep(self.sleepDurationLong)
        """
        MainContent_wizTasks_ctl17_chkPushOrganisationToXml
        """
        self.nav.navigateToClassID("MainContent_wizTasks_ctl17_chkPushOrganisationToXml")
        time.sleep(self.sleepDurationLong)
        self.nav.navigateToClassID("MainContent_wizTasks_ctl17_btnNextStart")
        time.sleep(self.sleepDurationLong)
        # MainContent_wizTasks_ctl22_btnFinishPushOrganisationToXml
        self.nav.navigateToClassID("MainContent_wizTasks_ctl22_btnFinishPushOrganisationToXml")
        time.sleep(self.sleepDurationLong)
        self.handleAudit()


    def navigateToOrg(self, ID):
        """
        :param ID:
        :return: Boolean
        """
        result = True

        self.driver.get('https://%s/OrganisationScreens/OrganisationMaintenance.aspx?ID=%s' % ( self.env, ID ) )

        time.sleep( self.sleepDuration )
        try:
            # if present the org isn't in this env
            pageTitle = self.driver.find_element_by_id("MainContent_pnlOrgDoesntExist")
            result =  False
        except NoSuchElementException:
            result =  True

        except UnexpectedAlertPresentException as e:
            time.sleep( self.sleepDuration )
            messageText = e.msg
            if messageText != "":
                if messageText.find( "This record is currently locked" ) >= 0:
                    # handle l;ocked record
                    self.driver.switch_to.alert.dismiss() # need to test this yet
                    msg = "Locked record issue bypassed %s" % ID
                    outputString = '"%s","%s","%s"\n' % (self.rowNum, ID, msg)
                    self.logFile.write("%s for ID %s on row %s\n" % (msg, ID, self.rowNum))
                else:
                    # assume it's the unsaved record bug
                    self.driver.switch_to.alert.accept() #.dismiss() did not work
                    msg = "Unsaved record issue bypassed on record preceeding %s" % ID
                    outputString = '"%s","%s","%s"\n' % (self.rowNum, ID, msg)
                    self.logFile.write("%s for ID %s on row %s\n" % (msg, ID, self.rowNum))
                    self.navigateToOrg( ID )

        return result

if __name__ == "__main__":
    args = sys.argv
    if len(args) == 6:
        env = sys.argv[1]
        user = sys.argv[2]
        pwd = sys.argv[3]
        file = sys.argv[4]
        auditText = sys.argv[5] # "CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE"

        xp = xmlPush(env, user, pwd, file, auditText)
        xp.process()
    else:
        msg = "Please provide the following arguments:\n"
        msg+= "1. environment, e.g. OSCAR-TESTSP\n"
        msg+= "2. user (shortcode only, Domain not required)\n"
        msg+= "3. password\n"
        msg+= "4. file (The file to be procesed - e.g. xyz.csv)\n"
        msg+= "5. auditText - e.g.CQC ALIGNMENT - SELENIUM AUTOMATED ACCEPTANCE\n"
        print( msg )