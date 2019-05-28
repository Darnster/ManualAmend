# module to share the logging functionality


class WL:

    def WriteLog(self, logFileName, msg):
        logFile = open(logFileName, 'a')
        logFile.write(msg)
        logFile.close()
