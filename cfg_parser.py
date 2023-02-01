"""
Utility class to read a config file into a dictionary

Needs to be updated in site_packages whenever an update is made
"""

import sys, os, string

def configParser():
    cp = config_parser()
    return cp


class config_parser(object):
    def __init__(self):
        pass #nothing to do

    def read(self, config_file):
        """
        IN - config_file
        OUT - dictionary of entries in config file
        """
        _config = config_file #"\\".join( (os.path.abspath(os.path.curdir), config_file) )

        try:
            fh = open(_config,'r')
        except:
            sys.stderr.write("unable to open config file %s.\n\n" % ( _config ) )
            sys.exit()
        self.config = [[x] for x in fh.readlines()]
        self.config = map(lambda x: str.split(str.strip(x[0]),':',1), self.config)
        self.config = map(lambda x: [str.strip(x[0]),str.strip(x[1])], self.config)
        _config_dict = {}
        [_config_dict.setdefault(*n) for n in self.config]
        fh.close()
        return _config_dict

if __name__ == "__main__":
    sys.stderr.write("This module cannot be run directly")
    sys.stderr.write("""syntax for import:\n
    from config_parser import config_parser\n
    cp = config_parser().read('config.txt')""")