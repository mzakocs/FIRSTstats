############################################################################################################################
# FirstConfig                                                                                                              #
# Used to grab config files from config.ini                                                                                #
# TODO: Make config editable from Google Sheets                                                                            #
############################################################################################################################

import os
import configparser
import base64

class FirstConfig:
    def __init__(self):
        # Opens the config file or creates it if it doesn't exist
        self.config = configparser.ConfigParser()
        if not os.path.exists('config.ini'):
            self.config ['Event Config'] = {'eventid':'CALN', 'season':'2019'}
            self.config ['FIRST API'] = {'Host':'https://frc-api.firstinspires.org', 'Username':'', 'Token':''}
            self.config ['Google Sheets'] = {'sheetid': ''}
            with open ('config.ini', 'w') as configfile:
                self.config.write(configfile)
        # Reads config file for values
        self.config.read('config.ini')
        self.eventid = self.config['Event Config']['eventid']
        self.season = self.config['Event Config']['season']
        self.sheetid = self.config['Google Sheets']['sheetid']
        self.host = self.config['FIRST API']['Host']
        self.username = self.config['FIRST API']['Username']
        self.password = self.config['FIRST API']['Token']
        self.authString = base64.b64encode('%s:%s' % (self.username, self.password))
    # TODO: Make config changable from sheets file