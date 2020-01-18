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
from gspread_formatting import *

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
        # Class Imports
        self.config = config
    def getScheduleData (self):
        response = requests.get('%s/v2.0/2019/schedule/%s/qual/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.qualScheduleData = response.json()["Schedule"]
        response = requests.get('%s/v2.0/2019/schedule/%s/playoff/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.playoffPlayoffData = response.json()["Schedule"]
    def getScoreData (self):
        response = requests.get('%s/v2.0/2019/scores/%s/qual' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.qualScoreData = response.json()["MatchScores"]
        response = requests.get('%s/v2.0/2019/scores/%s/playoff' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.playoffScoreData = response.json()["MatchScores"]
    def getEventData (self):
        response = requests.get('%s/v2.0/2019/events?eventCode=%s' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.eventData = response.json()["Events"][0]

class Sheets:
    def __init__(self, config, data):
        # Class Imports
        self.config = config
        self.data = data

        # Google Drive OAuth2 Setup
        # gspread.readthedocs.io/en/latest/
        scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('FIRST Python Stats-c64a29c90ec3.json', scope)
        self.gc = gspread.authorize(credentials)
        self.sh = self.gc.open_by_key(self.config.sheetid)

        # Formatting Setup
        title_format = CellFormat(textFormat = TextFormat(fontSize = 30, bold=True))
        venue_format = CellFormat(textFormat = TextFormat(fontSize = 20, bold=True))
        category_title_format = CellFormat(textFormat = TextFormat(fontSize = 14, bold=True))
        box_thick_format = CellFormat(borders = Borders(top = Border(style = SOLID_MEDIUM), bottom = Border(style = SOLID_MEDIUM), left = Border(style = SOLID_MEDIUM), right = Border(style = SOLID_MEDIUM)))
        box_format = CellFormat(borders = Borders(top = Border(style = SOLID), bottom = Border(style = SOLID), left = Border(style = SOLID), right = Border(style = SOLID)))
        centered_teamname_format = CellFormat(horizontalAlignment = CENTER, verticalAlignment = MIDDLE)
        blue_color_format = CellFormat(backgroundColor = Color(1, 0, 0, 0.3))
        red_color_format = CellFormat(backgroundColor = Color(0, 1, 0, 0.3))

        # Checks to see if the sheet exists, if not it gets created
        try:
            self.ws = self.sh.worksheet(str(self.config.eventid))
        except Exception as e:
            self.ws = self.sh.add_worksheet(title=str(self.config.eventid), rows="2000", cols="30")
            # Formatting the Cells
            title_format = CellFormat(textFormat = TextFormat(fontSize = 30, bold=True))
            venue_format = CellFormat(textFormat = TextFormat(fontSize = 20, bold=True))
            format_cell_range(self.ws, 'A1:A1', title_format)
            format_cell_range(self.ws, 'A2:A2', venue_format)
            # Adding the event info to the first 2 rows
            self.ws.update_cell(1, 1, self.data.eventData["name"])
            self.ws.update_cell(2, 1, self.data.eventData["venue"])
            print("Worksheet Created: " + str(e))

    def createMatchObjects(self):
        tempMatch = Match(self.data.qualScheduleData[0], self.data.qualScoreData[0], self)
        self.qualMatchList = [tempMatch] # Starts the list with a fake match so the match numbers match the index and aren't -1
        self.playoffMatchList = [tempMatch] # Starts the list with a fake match so the match numbers match the index -1
        for x in range(len(self.data.qualScheduleData)):
            tempMatch = Match(self.data.qualScheduleData[x], self.data.qualScoreData[x], self)
            self.qualMatchList.extend(tempMatch)
        for x in range(len(self.data.playoffScheduleData)):
            tempMatch = Match(self.data.playoffScheduleData[x], self.data.qualScoreData[x], self)
            self.playoffMatchList.extend(tempMatch)

class Match:
    # A class that represents a match entry in the spreadsheet
    # One is created for each match in an event
    def __init__(self, scheduledict, scoredict, ws):
        # Class Imports
        self.schedule = scheduledict
        self.score = scoredict
        self.ws = ws

        # Info about the Match Entry
        self.matchnum = self.schedule["matchNumber"]
        self.matchtype = self.schedule["tournamentLevel"]
        self.matchtitle = ("%s Match #%s" % self.matchtype, self.matchnum)
        self.o_x = 0 # X Location of the Entry
        self.o_y = 0 # Y Location of the Entry
        
        # Finds the coordinates of the entry, if it fails it creates a new entry below everything
        try:
            # Try to find the cell
            cell = ws.find (self.matchtitle)
            self.o_y = cell.row - 1
        except Exception as e:
            if (matchnum == 1 and matchtype == "Qualification")
            else:
                # If the search fails, get the location of the last entry, set the origin below it, and create the entry
                locatorcell_list = ws.findall("Auto Score")
                self.o_y = int(locatorcell_list.pop().row) + 5 # 5 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
                # Creates the entry's template
                self.createEntry()
                print ("Entry Created: %s" % e)
                
    def isDataValid(self):
        # TODO: Checks to see if the data is valid(int or ~), returns true or false
        print("nothing")
    def pushData(self):
        # TODO: Push all info from dict into sheet if data valid
        #       Also add data to team sheet
        print("nothing")
    def pushNullandPredict(self):
        # TODO: Push ~ into fields and predict data if data not valid
        print("nothing")
    def verifyEntry(self):
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
    def convertLocal(self, x, y):
        # TODO: Calculate pos based on origin and converts to A1
        temp_x = self.o_x + x
        temp_y = self.o_y + y
        return gspread.utils.rowcol_to_a1(temp_x, temp_y)
    def createEntry(self):
        # Match Title
        self.ws.update_acell(self.convertLocal(1, 1), self.matchtitle)
        self.
    
def main():
    config = Config()
    data = MatchData(config)
    sheets = Sheets(config, data)

    data.getScheduleData()
    data.getScoreData()
    data.getEventData()

    sheets.worksheetSetup()
    sheets.createMatchObjects()

if __name__ == "__main__":
    main()