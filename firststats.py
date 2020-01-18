import requests
import json
import base64
import configparser
import os
import collections
import time


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
    def getScheduleData (self, dprint = False):
        response = requests.get('%s/v2.0/2019/schedule/%s/qual/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.qualScheduleData = response.json()["Schedule"]
        response = requests.get('%s/v2.0/2019/schedule/%s/playoff/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.playoffScheduleData = response.json()["Schedule"]
        if dprint == True:
            print(json.dumps(self.qualScheduleData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScheduleData, sort_keys=True, indent=4))

    def getScoreData (self, dprint = False):
        response = requests.get('%s/v2.0/2019/scores/%s/qual' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.qualScoreData = response.json()["MatchScores"]
        response = requests.get('%s/v2.0/2019/scores/%s/playoff' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.playoffScoreData = response.json()["MatchScores"]
        if dprint == True:
            print(json.dumps(self.qualScoreData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScoreData, sort_keys=True, indent=4))

    def getEventData (self, dprint = False):
        response =requests.get('%s/v2.0/2019/events?eventCode=%s' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.eventData = response.json()["Events"][0]
        if dprint == True:
            print(json.dumps(self.eventData, sort_keys=True, indent=4))

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
        self.title_format = CellFormat(textFormat = TextFormat(fontSize = 30, bold=True))
        self.venue_format = CellFormat(textFormat = TextFormat(fontSize = 20, bold=True))
        self.matchtitle_format = CellFormat(textFormat = TextFormat(fontSize = 24, bold=True))
        self.category_title_format = CellFormat(textFormat = TextFormat(fontSize = 14, bold=True))
        self.box_thick_format = CellFormat(borders = Borders(top = Border(style = 'SOLID_MEDIUM'), bottom = Border(style = 'SOLID_MEDIUM'), left = Border(style = 'SOLID_MEDIUM'), right = Border(style = 'SOLID_MEDIUM')))
        self.box_format = CellFormat(borders = Borders(top = Border(style = 'SOLID'), bottom = Border(style = 'SOLID'), left = Border(style = 'SOLID'), right = Border(style = 'SOLID')))
        self.centered_teamname_format = CellFormat(horizontalAlignment = 'CENTER', verticalAlignment = 'MIDDLE')
        self.blue_color_format = CellFormat(backgroundColor = Color(1, 0, 0, 0.3))
        self.red_color_format = CellFormat(backgroundColor = Color(0, 1, 0, 0.3))

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
            # Setting the cell size to 135 for all cells
            body = {
                "requests": [
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": self.ws._properties['sheetId'],
                                "dimension": "COLUMNS",
                                "startIndex": 0,
                                "endIndex": 30
                            },
                            "properties": {
                                "pixelSize": 136
                            },
                            "fields": "pixelSize"
                        }
                    }
                ]
            }
            self.sh.batch_update(body)
            print("Worksheet Created: " + str(e))

    def updateCell(self, cella1, cellvalue):
        for cell in self.cell_list:
            if (cell.row == a1_to_rowcol(cella1)[0] and cell.col == a1_to_rowcol(cella1)[1]):
                cell.value = cellvalue
                return cell

    def createEntry(self, match):
        self.cell_list = self.ws.range('%s:%s' % (self.convertLocal(1, 1, match), self.convertLocal(6, 12, match))) # Grabs the whole section where it should be
        # Match Title
        self.formatCell(1, 1, match, self.matchtitle_format)
        # self.ws.update_acell(self.convertLocal(1, 1, match), match.matchtitle)
        self.updateCell(self.convertLocal(1, 1, match), match.matchtitle)
        # Section Names
        self.formatCell(1, 2, match, self.category_title_format)
        self.formatCell(4, 2, match, self.category_title_format)
        self.formatCell(1, 8, match, self.category_title_format)
        self.formatCell(4, 8, match, self.category_title_format)
        self.updateCell(self.convertLocal(1, 2, match), "Match Info")
        self.updateCell(self.convertLocal(4, 2, match), "Predictions")
        self.updateCell(self.convertLocal(1, 8, match), "Auto Score")
        self.updateCell(self.convertLocal(4, 8, match), "Manual Score")
        # Sub-Sections
        # Match Info
        self.updateCell(self.convertLocal(1, 3, match), "Team 1")
        self.updateCell(self.convertLocal(1, 4, match), "Team 2")
        self.updateCell(self.convertLocal(1, 5, match), "Team 3")
        self.updateCell(self.convertLocal(1, 6, match), "Switch Level")
        self.updateCell(self.convertLocal(1, 7, match), "Ranking Points")
        # Predictions
        self.updateCell(self.convertLocal(4, 3, match), "Team 1 MR")
        self.updateCell(self.convertLocal(4, 4, match), "Team 2 MR")
        self.updateCell(self.convertLocal(4, 5, match), "Team 3 MR")
        self.updateCell(self.convertLocal(4, 6, match), "Average MR")
        self.updateCell(self.convertLocal(4, 7, match), "Score Predict")
        # Auto Score
        self.updateCell(self.convertLocal(1, 9, match), "Inner")
        self.updateCell(self.convertLocal(1, 10, match), "Outer")
        self.updateCell(self.convertLocal(1, 11, match), "Lower")
        self.updateCell(self.convertLocal(1, 12, match), "Total")
        # Manual Score
        self.updateCell(self.convertLocal(4, 9, match), "Inner")
        self.updateCell(self.convertLocal(4, 10, match), "Outer")
        self.updateCell(self.convertLocal(4, 11, match), "Lower")
        self.updateCell(self.convertLocal(4, 12, match), "Total")
        # Red and Blue Team Sections
        
        self.ws.update_cells(self.cell_list)
    
    def formatCell(self, x, y, match, cellformat):
        tempcell = self.convertLocal(x, y, match)
        cellrange = '%s:%s' % (tempcell, tempcell)
        format_cell_range(self.ws, cellrange, cellformat)

    def convertLocal(self, x, y, match):
        # Calculate pos based on origin and converts to A1
        temp_x = match.o_x + x
        temp_y = match.o_y + y
        return gspread.utils.rowcol_to_a1(temp_y, temp_x)

    def createMatchObjects(self):
        # Creates lists to store the qualifier and playoff matches
        tempMatch = Match(self.data.qualScheduleData[0], self.data.qualScoreData[0])
        self.qualMatchList = [tempMatch] # Starts the list with a placeholder match so the match numbers match the index and aren't -1
        self.playoffMatchList = [tempMatch] # Starts the list with a placeholder match so the match numbers match the index -1

        # Qualifier Match Object Creation
        for x in range(0, len(self.data.qualScheduleData) - 1):
            tempMatch = Match(self.data.qualScheduleData[x], self.data.qualScoreData[x])
            try:
                # Try to find the cell
                cell = self.ws.find(tempMatch.matchtitle)
                tempMatch.o_y = cell.row - 1
            except Exception as e:
                if (x == 0):
                    tempMatch.o_y = 3 # If it's the very first match, set the origin to the third row
                else:
                    # Search fails, get the location of the last entry, set the origin below it, and create the entry
                    locatorcell_list = self.ws.findall("Auto Score")
                    tempMatch.o_y = int(locatorcell_list.pop().row) + 5 # 5 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
                # Creates the entry's template
                self.createEntry(tempMatch)
                print ("Entry Created: %s" % e)
            self.qualMatchList.insert(len(self.qualMatchList), tempMatch)

        # Playoff Match Object Creation
        for x in range(0, len(self.data.playoffScheduleData) - 1):
            tempMatch = Match(self.data.qualScheduleData[x], self.data.qualScoreData[x])
            try:
                # Try to find the cell
                cell = self.ws.find (self.qualMatchList[x].matchtitle)
                tempMatch.o_y = cell.row - 1
            except Exception as e:
                # Search fails, get the location of the last entry, set the origin below it, and create the entry
                locatorcell_list = self.ws.findall("Auto Score")
                tempMatch.o_y = int(locatorcell_list.pop().row) + 5 # 5 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
            # Creates the entry's template
            self.createEntry(tempMatch)
            print ("Entry Created: %s" % e)
            self.playoffMatchList.insert(len(self.playoffMatchList), tempMatch)

class Match:
    # A class that represents a match entry in the spreadsheet
    # One is created for each match in an event
    def __init__(self, scheduledict, scoredict):
        # Class Imports
        self.schedule = scheduledict
        self.score = scoredict

        # Info about the Match Entry
        self.matchnum = self.schedule["matchNumber"]
        self.matchtype = self.schedule["tournamentLevel"]
        self.matchtitle = ("%s Match #%s" % (self.matchtype, self.matchnum))
        self.o_x = 0 # X Location of the Entry
        self.o_y = 0 # Y Location of the Entry
        
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

def main():
    config = Config()
    data = MatchData(config)
    data.getScheduleData()
    data.getScoreData()
    data.getEventData()

    sheets = Sheets(config, data)

    sheets.createMatchObjects()

    sheets.createMatchObjects()

if __name__ == "__main__":
    main()
