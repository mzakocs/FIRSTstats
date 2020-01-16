import requests
import json
import base64
import configparser
import os
import collections


# TODO SECTION
# TODO: IMPLEMENT ALL UNFINISHED FUNCTIONS
# TODO: CREATE FAKE SCOREDETAIL JSON FOR 2020 TESTING (PAIGE)
# TODO: CREATE THE ALGORITHM MitchRating FOR RANKING TEAMS

# pip gspread, gspread_formatting, pyopenssl and oauth2client for google sheets functionality
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import gspread_formatting

class Config:
    def __init__(self):
        # Opens the config file or creates it if it doesn't exist
        self.config = configparser.ConfigParser()
        if not os.path.exists('config.ini'):
            self.config ['Event Config'] = {'eventid':'CMPMO'}
            self.config ['FIRST API'] = {'Host':'', 'Username':'', 'Token':'frc-api.firstinspires.org'}
            self.config ['Google Sheets'] = {'sheetid': ''}
            with open ('config.ini', 'w') as configfile:
                self.config.write(configfile)
        # Reads config file for values
        self.config.read('config.ini')
        self.eventid = self.config['Event Config']['eventid']
        self.sheetid = self.config['Google Sheets']['sheetid']
        self.host = self.config['FIRST API']['Host']
        self.username = self.config['FIRST API']['Username']
        self.password = self.config['FIRST API']['Token']
        self.authString = base64.b64encode('%s:%s' % (self.username, self.password))
    # TODO: Maybe make config changable from sheets file?

class MatchData:
    def __init__(self, config):
        self.config = config
    def getScheduleData (self):
        response = requests.get('%s/v2.0/2019/schedule/%s/qual/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.scheduleData = response.json()
        return self.scheduleData
    def getScoreData (self):
        response = requests.get('%s/v2.0/2019/scores/%s/qual' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.scoreData = response.json()
        return self.scoreData

class Sheets:
    def __init__(self, config, data):
        self.config = config
        # Google Drive OAuth2 Setup
        # gspread.readthedocs.io/en/latest/
        scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('FIRST Python Stats-c64a29c90ec3.json', scope)
        self.gc = gspread.authorize(credentials)
        self.sh = self.gc.open_by_key(self.config.sheetid)
    def worksheetSetup(self):
        # Creates the worksheet if it doesn't exist and sets it up with the match schedule
        try:
            self.ws = self.sh.worksheet(str(self.config.eventid))
        except Exception as e:
            self.ws = self.sh.add_worksheet(title=str(self.config.eventid), rows="200", cols="30")
            print("Worksheet Created: " + str(e))
        # TODO: Creates the template for all the matches
    def getMatchDict(self, matchtype, matchnum):
        # TODO: Returns the match's dictionary based on matchtype and matchid
        #       This is for setting up the Match objects
        print("nothing")
    def getScoreDict(self, matchtype, matchnum):
        # TODO: Returns the match's score dictionary based on matchtype and matchid
        #       This is for setting up the Match objects
        print("nothing")

        print("nothing")
    def createMatchObjects(self, scheduleData, scoreData):
        # TODO: Loop through all of the matches and feeds the matchdict and scoredict into a Match class
        #       Create a dictionary of objects named Matches
        #       Example Layout: {matchnum+1: Match(matchdict, scoredict) for matchnum in scheduleData}
        print("nothing")

class Match:
    def __init__(self, matchdict, scoredict):
        # Sets up origin and matchdict setup
        self.match = matchdict
        self.score = scoredict
        self.matchnum = self.match["matchNumber"]
        self.o_x = 0
        self.o_y = 0
        # TODO: Find entry position in sheets or makes the entry if it doesn't exist, then sets origin to entry position
    def isDataValid(self):
        # TODO: Checks to see if the data is valid(int or ~), returns true or false
        print("nothing")
    def pushData(self):
        # TODO: Push all info from dict into sheet if data valid
        #       Also add data to sheet
        print("nothing")
    def pushNullandPredict(self):
        # TODO: Push ~ into fields and predict data if data not valid
        print("nothing")
    def checkEntry(self, matchtype, matchnum):
        # TODO: Checks the entry in sheets to see if it matches the dict data
        #       If it doesn't match and is null, push new data (usually means match was played)
        #       If it doesn't match and isn't null, push new data and throw exception (means program fucked up somehow)
        #       If it does match and is null, do nothing as match hasn't occured yet
        #       If it does match and isn't null, do nothing as match data is there & match has occured
        print("nothing")
    def predictScores(self):
        # TODO: Only used when a match hasn't been played yet, predicts data based off team data
        #       Maybe check to see if like 20 games have been played first?
        print("nothing")

    
def main():
    config = Config()
    data = MatchData(config)
    sheets = Sheets(config, data)

    data.getScheduleData()
    data.getScoreData()
    print(json.dumps(data.scheduleData, sort_keys=True, indent=4, default=str))
    print(json.dumps(data.scoreData, sort_keys=True, indent=4, default=str))

    sheets.worksheetSetup()

if __name__ == "__main__":
    main()