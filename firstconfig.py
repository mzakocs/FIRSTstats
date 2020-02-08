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
            self.config ['Event Config'] = {'eventid':'', 'season':''}
            self.config ['Customization'] = {'teamsort':'TRUE', 'displayunplayedmatches':'TRUE', 'matchteamfilter':'ALL'}
            self.config ['FIRST API'] = {'Host':'https://frc-api.firstinspires.org', 'Username':'', 'Token':''}
            self.config ['Google Sheets'] = {'sheetid': '', 'oauthjsonpath':'FIRST Python Stats-c64a29c90ec3.json'}
            with open ('config.ini', 'w') as configfile:
                self.config.write(configfile)
            print ("Please fill out the config file and restart the application!")
            exit()
        # Reads config file for values
        self.config.read('config.ini')
        self.eventid = self.config['Event Config']['eventid']
        self.season = self.config['Event Config']['season']
        self.teamsort = self.config['Customization']['teamsort']
        self.displayunplayedmatches = self.config['Customization']['displayunplayedmatches']
        self.matchteamfilter = self.config['Customization']['matchteamfilter']
        self.sheetid = self.config['Google Sheets']['sheetid']
        self.oauthjsonpath = self.config['Google Sheets']['oauthjsonpath']
        self.host = self.config['FIRST API']['Host']
        self.username = self.config['FIRST API']['Username']
        self.password = self.config['FIRST API']['Token']
        self.authString = base64.b64encode(('%s:%s' % (self.username, self.password)).encode('utf-8'))

    def checkSheetsConfig(self, worksheet, csqp):
        ## Checks to see if the config matches the config on sheets
        # If it doesn't, update the values on the local config and returns False
        # If it does, simply returns True
        tempResult = True
        sheets_eventid = csqp.readCell(2, 1)
        sheets_season = int(csqp.readCell(2, 2))
        sheets_teamsort = csqp.readCell(2, 6)
        sheets_displayunplayedmatches = csqp.readCell(2, 7)
        sheets_matchteamfilter = csqp.readCell(2, 8)
        # Compares the values
        if not sheets_eventid == self.eventid:
            self.config['Event Config']['eventid'] = sheets_eventid
            tempResult == False
        if not sheets_season == self.season:
            self.config['Event Config']['season'] = sheets_season
            tempResult == False
        if not sheets_teamsort == self.teamsort:
            self.config['Customization']['teamsort'] = sheets_teamsort
            tempResult == "Filter"
        if not sheets_displayunplayedmatches == self.displayunplayedmatches:
            self.config['Customization']['displayunplayedmatches'] = sheets_displayunplayedmatches
            tempResult == "Filter"
        if not sheets_matchteamfilter == self.matchteamfilter:
            self.config['Customization']['matchteamfilter'] = sheets_matchteamfilter
            tempResult == "Filter"
        if tempResult != True:
            with open ('config.ini', 'w') as configfile:
                self.config.write(configfile)
        return tempResult