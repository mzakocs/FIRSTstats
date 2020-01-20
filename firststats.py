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
    def __init__(self, config, testing = False):
        # Class Imports
        self.config = config
        self.testing = testing

    ### Data Retrieval Functions ###
    ## Gets Data from the FIRST API
    # Testing parameter allows usage of a local JSON file for testing instead of the FIRST API
    # DPrint parameter formats the dictionary into a readable format and prints it
    def getScheduleData (self, dprint = False):
        if (self.testing == True):
            with open('qualscheduledata.json') as json_file:
                self.qualScheduleData = json.load(json_file)
            with open('playoffscheduledata.json') as json_file:
                self.playoffScheduleData = json.load(json_file)
        else:    
            qualresponse = requests.get('%s/v2.0/2019/schedule/%s/qual/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            playoffresponse = requests.get('%s/v2.0/2019/schedule/%s/playoff/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.qualScheduleData = qualresponse.json()["Schedule"]
            self.playoffScheduleData = playoffresponse.json()["Schedule"]
        if dprint == True:
            print(json.dumps(self.qualScheduleData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScheduleData, sort_keys=True, indent=4))

    def getScoreData (self, dprint = False):
        if (self.testing == True):
            with open('qualscoredata.json') as json_file:
                self.qualScheduleData = json.load(json_file)
            with open('playoffscoredata.json') as json_file:
                self.playoffScheduleData = json.load(json_file)
        response = requests.get('%s/v2.0/2019/scores/%s/qual' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.qualScoreData = response.json()["MatchScores"]
        response = requests.get('%s/v2.0/2019/scores/%s/playoff' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
        self.playoffScoreData = response.json()["MatchScores"]
        if dprint == True:
            print(json.dumps(self.qualScoreData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScoreData, sort_keys=True, indent=4))

    def getEventData (self, dprint = False):
        if (self.testing == True):
            with open('qualeventdata.json') as json_file:
                self.qualScheduleData = json.load(json_file)
            with open('playoffeventdata.json') as json_file:
                self.playoffScheduleData = json.load(json_file)
        response = requests.get('%s/v2.0/2019/events?eventCode=%s' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
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
        self.centered_teamname_format = CellFormat(horizontalAlignment = 'CENTER', verticalAlignment = 'MIDDLE')
        self.red_color_format = CellFormat(backgroundColor = Color(0.94, 0.5, 0.5))
        self.blue_color_format = CellFormat(backgroundColor = Color(0.68, 0.85, 0.9))
        self.match_played_format = CellFormat(backgroundColor = Color(0.56, 0.94, 0.56))
        self.match_scheduled_format = CellFormat(backgroundColor = Color(0.83, 0.83, 0.83))
        self.box_thick_format = CellFormat(borders = Borders(top = Border(style = 'SOLID_THICK'), bottom = Border(style = 'SOLID_THICK'), left = Border(style = 'SOLID_THICK'), right = Border(style = 'SOLID_THICK')))
        self.box_format = CellFormat(borders = Borders(top = Border(style = 'SOLID'), bottom = Border(style = 'SOLID'), left = Border(style = 'SOLID'), right = Border(style = 'SOLID')))
        self.box_left_format = CellFormat(borders = Borders(left = Border(style = 'SOLID')))
        self.box_right_format = CellFormat(borders = Borders(right = Border(style = 'SOLID')))
        self.box_top_format = CellFormat(borders = Borders(top = Border(style = 'SOLID')))
        self.box_bottom_format = CellFormat(borders = Borders(bottom = Border(style = 'SOLID')))
        self.box_left_thick_format = CellFormat(borders = Borders(left = Border(style = 'SOLID_THICK')))
        self.box_right_thick_format = CellFormat(borders = Borders(right = Border(style = 'SOLID_THICK')))
        self.box_top_thick_format = CellFormat(borders = Borders(top = Border(style = 'SOLID_THICK')))
        self.box_bottom_thick_format = CellFormat(borders = Borders(bottom = Border(style = 'SOLID_THICK')))

        # Checks to see if the sheet exists, if not it gets created
        try:
            self.ws = self.sh.worksheet(str(self.config.eventid))
        except Exception as e:
            self.ws = self.sh.add_worksheet(title=str(self.config.eventid), rows="3000", cols="30")
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
    

    def createEntry(self, match):
        ### Compressed Sheets Query Protocol Setup ###
        csqp = CSQP(self, match)

        ### Match Title ###
        csqp.updateCellValue(1, 1, match, match.matchtitle)
        csqp.updateCellFormatting(1, 1, match, self.matchtitle_format)

        ### Section Names ###
        csqp.updateCellFormatting(1, 2, match, self.category_title_format)
        csqp.updateCellFormatting(4, 2, match, self.category_title_format)
        csqp.updateCellFormatting(1, 8, match, self.category_title_format)
        csqp.updateCellFormatting(4, 8, match, self.category_title_format)
        csqp.updateCellValue(1, 2, match, "Match Info")
        csqp.updateCellValue(4, 2, match, "Predictions")
        csqp.updateCellValue(1, 8, match, "Auto Score")
        csqp.updateCellValue(4, 8, match, "Teleop Score")

        ### Sub-Sections ###
        ## Match Info
        csqp.updateCellValue(1, 3, match, "Team 1")
        csqp.updateCellValue(1, 4, match, "Team 2")
        csqp.updateCellValue(1, 5, match, "Team 3")
        csqp.updateCellValue(1, 6, match, "Switch Level")
        csqp.updateCellValue(1, 7, match, "Ranking Points")
        ## Predictions
        csqp.updateCellValue(4, 3, match, "Team 1 MR")
        csqp.updateCellValue(4, 4, match, "Team 2 MR")
        csqp.updateCellValue(4, 5, match, "Team 3 MR")
        csqp.updateCellValue(4, 6, match, "Average MR")
        csqp.updateCellValue(4, 7, match, "Score Predict")
        ## Auto Score
        csqp.updateCellValue(1, 9, match, "Inner")
        csqp.updateCellValue(1, 10, match, "Outer")
        csqp.updateCellValue(1, 11, match, "Lower")
        csqp.updateCellValue(1, 12, match, "Total")
        ## Teleop Score
        csqp.updateCellValue(4, 9, match, "Inner")
        csqp.updateCellValue(4, 10, match, "Outer")
        csqp.updateCellValue(4, 11, match, "Lower")
        csqp.updateCellValue(4, 12, match, "Total")

        ### Team Setup ###
        ## Red and Blue Colors
        csqp.updateCellRangeFormatting(2, 2, 2, 12, match, self.red_color_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 12, match, self.blue_color_format)
        csqp.updateCellRangeFormatting(5, 2, 5, 12, match, self.red_color_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 12, match, self.blue_color_format)
        ## Team Color Labels
        csqp.updateCellValue(2, 2, match, "Red")
        csqp.updateCellValue(2, 8, match, "Red")
        csqp.updateCellValue(5, 2, match, "Red")
        csqp.updateCellValue(5, 8, match, "Red")
        csqp.updateCellValue(3, 2, match, "Blue")
        csqp.updateCellValue(3, 8, match, "Blue")
        csqp.updateCellValue(6, 2, match, "Blue")
        csqp.updateCellValue(6, 8, match, "Blue")
        csqp.updateCellRangeFormatting(2, 2, 3, 2, match, self.centered_teamname_format)
        csqp.updateCellRangeFormatting(2, 8, 3, 8, match, self.centered_teamname_format)
        csqp.updateCellRangeFormatting(5, 2, 6, 2, match, self.centered_teamname_format)
        csqp.updateCellRangeFormatting(5, 8, 6, 8, match, self.centered_teamname_format)

        ### Lines ###
        ## Top Category Divider Lines
        csqp.updateCellRangeFormatting(1, 2, 6, 2, match, self.box_format)
        csqp.updateCellRangeFormatting(1, 8, 6, 8, match, self.box_format)
        ## Main Category Line
        # Match Info
        csqp.updateCellRangeFormatting(1, 2, 1, 7, match, self.box_left_format)
        csqp.updateCellRangeFormatting(1, 2, 1, 7, match, self.box_right_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 7, match, self.box_left_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 7, match, self.box_right_format)
        # Predictions
        csqp.updateCellRangeFormatting(4, 2, 4, 7, match, self.box_left_format)
        csqp.updateCellRangeFormatting(4, 2, 4, 7, match, self.box_right_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 7, match, self.box_left_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 7, match, self.box_right_format)
        # Auto Score
        csqp.updateCellRangeFormatting(1, 8, 1, 12, match, self.box_left_format)
        csqp.updateCellRangeFormatting(1, 8, 1, 12, match, self.box_right_format)
        csqp.updateCellRangeFormatting(3, 8, 3, 12, match, self.box_left_format)
        csqp.updateCellRangeFormatting(3, 8, 3, 12, match, self.box_right_format)
        # Teleop Score
        csqp.updateCellRangeFormatting(4, 8, 4, 12, match, self.box_left_format)
        csqp.updateCellRangeFormatting(4, 8, 4, 12, match, self.box_right_format)
        csqp.updateCellRangeFormatting(6, 8, 6, 12, match, self.box_left_format)
        csqp.updateCellRangeFormatting(6, 8, 6, 12, match, self.box_right_format)

        ### Thicc Outline Line ###
        csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.box_top_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(6, 1, 6, 16, match, self.box_right_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 1, 16, match, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(1, 12, 6, 12, match, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(1, 16, 6, 16, match, self.box_bottom_thick_format)

        ### Scouting Notes Line Setup ###
        csqp.updateCellValue(1, 13, match, "Scouting Notes")
        csqp.updateCellFormatting(1, 13, match, self.category_title_format)
        csqp.updateCustomCellFormatting(1, 13, 6, 13, match, "merge")
        csqp.updateCustomCellFormatting(1, 14, 6, 16, match, "merge")

        ### Entry Background Color ###
        ## Sets background color to green if the match has been played, grey otherwise
        if (match.schedule["postResultTime"] == "null"):
            csqp.updateCellRangeFormatting(1, 2, 1, 12, match, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(4, 2, 4, 12, match, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(1, 13, 6, 16, match, self.match_scheduled_format)
        else:
            csqp.updateCellRangeFormatting(1, 2, 1, 12, match, self.match_played_format)
            csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.match_played_format)
            csqp.updateCellRangeFormatting(4, 2, 4, 12, match, self.match_played_format)
            csqp.updateCellRangeFormatting(1, 13, 6, 16, match, self.match_played_format)
        
        # Pushes all of the cells to google sheets
        csqp.pushCellUpdate()

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
                    tempMatch.o_y = int(locatorcell_list.pop().row) + 9 # 5 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
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

class CSQP:
    ### Compressed Sheets Query Protocol ###
    ## This was created because of the query limit on Google Sheets API v4
    ## Since it's limited to 100 requests per 100 seconds, we need to use as little requests as possible
    #   This system grabs the cells we need to edit, edits them locally in python,
    #   and then pushes all of the cells back to google sheets in 1 request.
    #   This is much more efficient than using 1 request to edit one cell.
    def __init__(self, ws, match):
        # Class Imports
        self.ws = ws
        self.match = match

        # Local Query List Setup
        self.cell_list = self.ws.ws.range('%s:%s' % (self.ws.convertLocal(1, 1, self.match), self.ws.convertLocal(6, 16, self.match))) # Grabs the whole section of cells
        self.cell_formatting = []
        self.custom_requests = {"requests": []}

    def updateCellValue(self, x, y, match, cellvalue):
        tempcell = self.ws.convertLocal(x, y, match)
        # Finds the cell in the cell list and updates the value
        for cell in self.cell_list:
            if (cell.row == a1_to_rowcol(tempcell)[0] and cell.col == a1_to_rowcol(tempcell)[1]):
                cell.value = cellvalue
                return cell

    def updateCellFormatting(self, x, y, match, formatsettings):
        tempcell = self.ws.convertLocal(x, y, match) # Converts the x, y coordinates to A1 format and local
        tempcellrange = "%s:%s" % (tempcell, tempcell) # GSpread_formatting requires a range of cells for some reason
        # Puts the formatting options into a tuple and stores it in a list
        self.cell_formatting.insert(len(self.cell_formatting), (tempcellrange, formatsettings))

    def updateCellRangeFormatting(self, x1, y1, x2, y2, match, formatsettings):
        # Converts the 2 cells into A1 format and then puts it into the range format for gspread_formatting
        tempcell1 = self.ws.convertLocal(x1, y1, match) 
        tempcell2 = self.ws.convertLocal(x2, y2, match) 
        tempcellrange = "%s:%s" % (tempcell1, tempcell2)
        # Puts the formatting options into a tuple and stores it in a list
        self.cell_formatting.insert(len(self.cell_formatting), (tempcellrange, formatsettings))

    def updateCustomCellFormatting(self, x1, y1, x2, y2, match, requesttype):
        # Formulates a custom google sheets request for merging cells
        if (requesttype == "merge"):
            temp_x1 = match.o_x + x1 - 1
            temp_y1 = match.o_y + y1 - 1
            temp_x2 = match.o_x + x2
            temp_y2 = match.o_y + y2 
            temprequest = {
                            "mergeCells": {
                                "mergeType": "MERGE_ALL",
                                "range": {
                                    "sheetId": self.ws.ws._properties['sheetId'],
                                    "startRowIndex": temp_y1,
                                    "endRowIndex": temp_y2,
                                    "startColumnIndex": temp_x1,
                                    "endColumnIndex": temp_x2
                                } 
                            }
                        }
            self.custom_requests["requests"].insert(len(self.custom_requests), temprequest)
        else:
            print("Custom Request Type Not Found!")

    def pushCellUpdate(self):
        # Pushes the array created by updateCellFormatting to the google sheet
        format_cell_ranges(self.ws.ws, self.cell_formatting)
        # Pushes cell merge requests
        self.ws.sh.batch_update(self.custom_requests)
        # Pushes the cell value list to the google sheet
        self.ws.ws.update_cells(self.cell_list)

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
