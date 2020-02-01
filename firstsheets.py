############################################################################################################################
# FirstSheets                                                                                                              #
# Used to store and retrieve data from the Google Sheets                                                                   #
# Implements CSQP to use less queries                                                                                      #
############################################################################################################################

# pip gspread, gspread_formatting, pyopenssl and oauth2client for google sheets functionality
import time
import gspread
import math
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *

class Sheets:
    def __init__(self, config, data):
        # Class Imports
        self.config = config
        self.data = data

        # Google Drive OAuth2 Setup
        # gspread.readthedocs.io/en/latest/
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('FIRST Python Stats-c64a29c90ec3.json', scope)
        self.gc = gspread.authorize(credentials)
        self.sh = self.gc.open_by_key(self.config.sheetid)

        # Formatting Setup
        self.font_format = CellFormat(textFormat = TextFormat(fontFamily="Montserrat"))
        self.title_format = CellFormat(textFormat = TextFormat(fontSize = 30, bold=True))
        self.venue_format = CellFormat(textFormat = TextFormat(fontSize = 20, bold=True))
        self.bold_format = CellFormat(textFormat = TextFormat(bold=True))
        self.matchtitle_format = CellFormat(textFormat = TextFormat(fontSize = 24, bold=True))
        self.category_title_format = CellFormat(textFormat = TextFormat(fontSize = 14, bold=True))
        self.centered_format = CellFormat(horizontalAlignment = 'CENTER', verticalAlignment = 'MIDDLE')
        self.red_color_format = CellFormat(backgroundColor = Color(0.94, 0.5, 0.5))
        self.blue_color_format = CellFormat(backgroundColor = Color(0.68, 0.85, 0.9))
        self.match_played_format = CellFormat(backgroundColor = Color(0.59, 0.98, 0.59, 0.5))
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
        self.purple_title_format = CellFormat(backgroundColor = Color(0.6, 0.2, 0.4))
        self.purple1_format = CellFormat(backgroundColor = Color(0.78, 0.33, 0.55))
        self.purple2_format = CellFormat(backgroundColor = Color(0.85, 0.55, 0.71))

        # Checks to see if the sheet exists, if not it gets created
        try:
            self.ws = self.sh.worksheet(str(self.config.eventid))
        except Exception as e:
            self.ws = self.sh.add_worksheet(title=str(self.config.eventid), rows="3000", cols="30")
            # Formatting the Cells
            format_cell_range(self.ws, 'A1:A1', self.title_format)
            format_cell_range(self.ws, 'A2:A2', self.venue_format)
            # Adding the event info to the first 2 rows
            self.ws.update_cell(1, 1, self.data.eventData["name"])
            self.ws.update_cell(2, 1, self.data.eventData["venue"])
            # Setting the cell size to 135 for all cells and changing the font
            temprange = "A1:AD3000"
            format_cell_range(self.ws, temprange, self.font_format)
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
    

    def createMatchEntry(self, match):
        ### Compressed Sheets Query Protocol Setup ###
        csqp = UCSQP(self, match.o_x, match.o_y, 6, 17)

        ### Match Title ###
        csqp.updateCellValue(1, 1, match.matchtitle)
        csqp.updateCellFormatting(1, 1, self.matchtitle_format)

        ### Section Names ###
        csqp.updateCellFormatting(1, 2, self.category_title_format)
        csqp.updateCellFormatting(4, 2, self.category_title_format)
        csqp.updateCellFormatting(1, 8, self.category_title_format)
        csqp.updateCellFormatting(4, 8, self.category_title_format)
        csqp.updateCellValue(1, 2, "Match Info")
        csqp.updateCellValue(4, 2, "Predictions")
        csqp.updateCellValue(1, 8, "Auto Score")
        csqp.updateCellValue(4, 8, "Teleop Score")

        ### Sub-Sections ###
        ## Centered Values Setup
        csqp.updateCellRangeFormatting(2, 2, 3, 13, self.centered_format)
        csqp.updateCellRangeFormatting(5, 2, 6, 13, self.centered_format)
        
        ## Match Info
        csqp.updateCellValue(1, 3, "Team 1")
        csqp.updateCellValue(2, 3, match.schedule["teams"][0]["teamNumber"]) # Red
        csqp.updateCellValue(3, 3, match.schedule["teams"][3]["teamNumber"]) # Blue

        csqp.updateCellValue(1, 4, "Team 2")
        csqp.updateCellValue(2, 4, match.schedule["teams"][1]["teamNumber"]) # Red
        csqp.updateCellValue(3, 4, match.schedule["teams"][4]["teamNumber"]) # Blue

        csqp.updateCellValue(1, 5, "Team 3")
        csqp.updateCellValue(2, 5, match.schedule["teams"][2]["teamNumber"]) # Red
        csqp.updateCellValue(3, 5, match.schedule["teams"][5]["teamNumber"]) # Blue

        csqp.updateCellValue(1, 6, "Switch Level")
        if not(match.score["alliances"][1]["endgameRungIsLevel"] == "null"):
            csqp.updateCellValue(2, 6, (match.score["alliances"][1]["endgameRungIsLevel"] == "IsLevel")) # Red
            csqp.updateCellValue(3, 6, (match.score["alliances"][0]["endgameRungIsLevel"] == "IsLevel")) # Blue
        else:
            csqp.updateCellValue(2, 6, "null") # Red
            csqp.updateCellValue(3, 6, "null") # Blue

        csqp.updateCellValue(1, 7, "Ranking Points")
        csqp.updateCellValue(2, 7, match.score["alliances"][1]["rp"]) # Red
        csqp.updateCellValue(3, 7, match.score["alliances"][0]["rp"]) # Blue

        ## Predictions
        # WIP
        csqp.updateCellValue(4, 3, "Team 1 MR")
        csqp.updateCellValue(4, 4, "Team 2 MR")
        csqp.updateCellValue(4, 5, "Team 3 MR")
        csqp.updateCellValue(4, 6, "Alliance Av. MR")
        csqp.updateCellValue(4, 7, "Score Prediction")

        ## Auto Score
        csqp.updateCellValue(1, 9, ("Inner " + u"\u25CF"))
        csqp.updateCellValue(2, 9, match.score["alliances"][1]["autoCellsInner"]) # Red
        csqp.updateCellValue(3, 9, match.score["alliances"][0]["autoCellsInner"]) # Blue

        csqp.updateCellValue(1, 10, ("Outer " + u"\u2B23"))
        csqp.updateCellValue(2, 10, match.score["alliances"][1]["autoCellsOuter"]) # Red
        csqp.updateCellValue(3, 10, match.score["alliances"][0]["autoCellsOuter"]) # Blue

        csqp.updateCellValue(1, 11, ("Bottom " + u"\u25A2"))
        csqp.updateCellValue(2, 11, match.score["alliances"][1]["autoCellsBottom"]) # Red
        csqp.updateCellValue(3, 11, match.score["alliances"][0]["autoCellsBottom"]) # Blue

        csqp.updateCellValue(1, 12, "Auto Total")
        if not(match.score["alliances"][1]["autoPoints"] == "null"):
            csqp.updateCellValue(2, 12, (int(match.score["alliances"][1]["autoCellsInner"]) + int(match.score["alliances"][1]["autoCellsOuter"]) + int(match.score["alliances"][1]["autoCellsBottom"]))) # Red
            csqp.updateCellValue(3, 12, (int(match.score["alliances"][0]["autoCellsInner"]) + int(match.score["alliances"][0]["autoCellsOuter"]) + int(match.score["alliances"][0]["autoCellsBottom"]))) # Blue
        else:
            csqp.updateCellValue(2, 12, "null") # Red
            csqp.updateCellValue(3, 12, "null") # Blue

        csqp.updateCellValue(1, 13, "Other Points")
        csqp.updateCellFormatting(1, 13, self.bold_format)
        if not(match.score["alliances"][1]["foulPoints"] == "null"):
            csqp.updateCellValue(2, 13, (int(match.score["alliances"][1]["foulPoints"]) + int(match.score["alliances"][1]["adjustPoints"]) + int(match.score["alliances"][1]["controlPanelPoints"]) + int(match.score["alliances"][1]["endgamePoints"]))) # Red
            csqp.updateCellValue(3, 13, (int(match.score["alliances"][0]["foulPoints"]) + int(match.score["alliances"][0]["adjustPoints"]) + int(match.score["alliances"][0]["controlPanelPoints"]) + int(match.score["alliances"][0]["endgamePoints"]))) # Blue
        else:
            csqp.updateCellValue(2, 13, "null") # Red
            csqp.updateCellValue(3, 13, "null") # Blue

        ## Teleop Score
        csqp.updateCellValue(4, 9, ("Inner " + u"\u25CF"))
        csqp.updateCellValue(5, 9, match.score["alliances"][1]["teleopCellsInner"]) # Red
        csqp.updateCellValue(6, 9, match.score["alliances"][0]["teleopCellsInner"]) # Blue

        csqp.updateCellValue(4, 10, ("Outer " + u"\u2B23"))
        csqp.updateCellValue(5, 10, match.score["alliances"][1]["teleopCellsOuter"]) # Red
        csqp.updateCellValue(6, 10, match.score["alliances"][0]["teleopCellsOuter"]) # Blue

        csqp.updateCellValue(4, 11, ("Bottom " + u"\u25A2"))
        csqp.updateCellValue(5, 11, match.score["alliances"][1]["teleopCellsBottom"]) # Red
        csqp.updateCellValue(6, 11, match.score["alliances"][0]["teleopCellsBottom"]) # Blue
        
        csqp.updateCellValue(4, 12, "Teleop Total")
        if not(match.score["alliances"][1]["autoPoints"] == "null"):
            csqp.updateCellValue(5, 12, (int(match.score["alliances"][1]["teleopCellsInner"]) + int(match.score["alliances"][1]["teleopCellsOuter"]) + int(match.score["alliances"][1]["teleopCellsBottom"]))) # Red
            csqp.updateCellValue(6, 12, (int(match.score["alliances"][0]["teleopCellsInner"]) + int(match.score["alliances"][0]["teleopCellsOuter"]) + int(match.score["alliances"][0]["teleopCellsBottom"]))) # Blue
        else:
            csqp.updateCellValue(5, 12, "null") # Red
            csqp.updateCellValue(6, 12, "null") # Blue

        csqp.updateCellValue(4, 13, "Score Total")
        csqp.updateCellRangeFormatting(4, 13, 6, 13, self.bold_format)
        csqp.updateCellValue(5, 13, match.score["alliances"][1]["totalPoints"]) # Red
        csqp.updateCellValue(6, 13, match.score["alliances"][0]["totalPoints"]) # Blue

        ### Team Setup ###
        ## Red and Blue Colors
        csqp.updateCellRangeFormatting(2, 2, 2, 13, self.red_color_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 13, self.blue_color_format)
        csqp.updateCellRangeFormatting(5, 2, 5, 13, self.red_color_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 13, self.blue_color_format)
        ## Team Color Labels
        csqp.updateCellValue(2, 2, "Red")
        csqp.updateCellValue(2, 8, "Red")
        csqp.updateCellValue(5, 2, "Red")
        csqp.updateCellValue(5, 8, "Red")
        csqp.updateCellValue(3, 2, "Blue")
        csqp.updateCellValue(3, 8, "Blue")
        csqp.updateCellValue(6, 2, "Blue")
        csqp.updateCellValue(6, 8, "Blue")

        ### Lines & Boxes ###
        ## Top Category Divider Lines
        csqp.updateCellRangeFormatting(1, 2, 6, 2, self.box_format)
        csqp.updateCellRangeFormatting(1, 8, 6, 8, self.box_format)
        csqp.updateCellRangeFormatting(1, 13, 6, 13, self.box_format)
        ## Main Category Line
        # Match Info
        csqp.updateCellRangeFormatting(1, 2, 1, 7, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(1, 2, 1, 7, self.box_right_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 7, self.box_left_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 7, self.box_right_thick_format)
        # Predictions
        csqp.updateCellRangeFormatting(4, 2, 4, 7, self.box_left_format)
        csqp.updateCellRangeFormatting(4, 2, 4, 7, self.box_right_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 7, self.box_left_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 7, self.box_right_thick_format)
        # Auto Score
        csqp.updateCellRangeFormatting(1, 8, 1, 13, self.box_left_format)
        csqp.updateCellRangeFormatting(1, 8, 1, 13, self.box_right_format)
        csqp.updateCellRangeFormatting(3, 8, 3, 13, self.box_left_format)
        csqp.updateCellRangeFormatting(3, 8, 3, 13, self.box_right_thick_format)
        # Teleop Score
        csqp.updateCellRangeFormatting(4, 8, 4, 13, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(4, 8, 4, 13, self.box_right_format)
        csqp.updateCellRangeFormatting(6, 8, 6, 13, self.box_left_format)
        csqp.updateCellRangeFormatting(6, 8, 6, 13, self.box_right_format)

        ### Scouting Notes Line Setup ###
        csqp.updateCellValue(1, 14, "Scouting Notes")
        csqp.updateCellFormatting(1, 14, self.category_title_format)
        csqp.updateCustomCellFormatting(1, 14, 6, 14, "merge")
        csqp.updateCustomCellFormatting(1, 15, 6, 17, "merge")
        csqp.updateCellRangeFormatting(1, 14, 6, 14, self.box_bottom_format)
        
        ### Thicc Outline Setup ###
        csqp.updateCellRangeFormatting(1, 1, 6, 1, self.box_top_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 6, 1, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(6, 1, 6, 17, self.box_right_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 1, 17, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(1, 13, 6, 13, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(1, 17, 6, 17, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(1, 8, 6, 8, self.box_top_thick_format)
        csqp.updateCellRangeFormatting(1, 13, 6, 13, self.box_top_thick_format)

        ### Entry Background Color ###
        ## Sets background color to green if the match has been played, grey otherwise
        if (match.schedule["postResultTime"] == "null"):
            csqp.updateCellRangeFormatting(1, 2, 1, 13, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(1, 1, 6, 1, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(4, 2, 4, 13, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(1, 14, 6, 17, self.match_scheduled_format)
        else:
            csqp.updateCellRangeFormatting(1, 2, 1, 13, self.match_played_format)
            csqp.updateCellRangeFormatting(1, 1, 6, 1, self.match_played_format)
            csqp.updateCellRangeFormatting(4, 2, 4, 13, self.match_played_format)
            csqp.updateCellRangeFormatting(1, 14, 6, 17, self.match_played_format)
        
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
        for x in range(0, len(self.data.qualScheduleData)):
            tempMatch = Match(self.data.qualScheduleData[x], self.data.qualScoreData[x])
            try:
                # Try to find the cell
                cell = self.ws.find(tempMatch.matchtitle)
                tempMatch.o_y = cell.row - 1
                tempMatch.o_x = cell.col - 1
                print ("Match Entry Found: %s" % tempMatch.matchtitle)
            except Exception as e:
                if (x == 0):
                    tempMatch.o_y = 3 # If it's the very first match, set the origin to the third row
                else:
                    # Search fails, get the location of the last entry, set the origin below it, and create the entry
                    locatorcell_list = self.ws.findall("Auto Score")
                    tempMatch.o_y = int(locatorcell_list.pop().row) + 10 # 10 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
                # Creates the entry in sheets
                self.createMatchEntry(tempMatch)
                print ("Match Entry Created: %s" % e)
            self.qualMatchList.insert(len(self.qualMatchList), tempMatch)

        # Playoff Match Object Creation
        for x in range(0, len(self.data.playoffScheduleData)):
            tempMatch = Match(self.data.playoffScheduleData[x], self.data.playoffScoreData[x])
            try:
                # Try to find the cell
                cell = self.ws.find (tempMatch.matchtitle)
                tempMatch.o_y = cell.row - 1
                tempMatch.o_x = cell.col - 1
                print ("Match Entry Found: %s" % tempMatch.matchtitle)
            except Exception as e:
                # Search fails, get the location of the last entry, set the origin below it, and create the entry
                locatorcell_list = self.ws.findall("Auto Score")
                tempMatch.o_y = int(locatorcell_list.pop().row) + 10 # 10 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
            # Creates the entry's in sheets
                self.createMatchEntry(tempMatch)
                print ("Match Entry Created: %s" % e)
            self.playoffMatchList.insert(len(self.playoffMatchList), tempMatch)

    def createTeamObjects(self):
        self.teamDict = {}
        for x in range(0, len(self.data.teamData)):
            tempTeam = Team(self.data.teamData[x])
            try:
                # Try to find the cell
                cell = self.ws.find (tempTeam.name)
                tempTeam.o_y = cell.row - 1
                tempTeam.o_x = cell.col - 1
                print ("Team Entry Found: %s" % tempTeam.name)
            except Exception as e:
                if x == 0:
                    tempTeam.o_y = 3 # If it's the very firtst match, set the origin below the title and column defs
                    tempTeam.o_x = 8 # If it's the very fist match, set the origin to 8 to the right
                else:
                    # Search fails, get the location of the last entry, set the origin below it, and create the entry
                    locatorcell_list = self.ws.findall(self.data.teamData[x - 1].name)
                    lastcell = locatorcell_list.pop()
                    tempTeam.o_y = int(lastcell.row) + 1 # 1 below the last entry
                    tempTeam.o_x = int(lastcell.col)
                print ("Team Entry Created: %s" % e)
            self.teamDict[tempTeam.number] = tempTeam
        self.createTeamEntry()

    def createTeamEntry(self):
        csqp = UCSQP(self, 8, 1, 11, (2 + len(self.teamDict)))
        limit = len(self.teamDict)
        # Title
        csqp.updateCellFormatting(1, 1, self.matchtitle_format)
        csqp.updateCellValue(1, 1, "Team List")
        # Columns
        csqp.updateCellValue(1, 2, "Team Name")
        csqp.updateCellValue(2, 2, "Team Number")
        csqp.updateCellValue(2, 3, ("MitchRating" + u"\u2122"))
        csqp.updateCellValue(2, 4, "Rank Title")
        # Boxes
        csqp.updateCellRangeFormatting(1, 2, 4, 2, self.box_format)
        csqp.updateCellRangeFormatting(1, 3, 1, limit, self.box_right_format)
        csqp.updateCellRangeFormatting(2, 3, 2, limit, self.box_right_format)
        csqp.updateCellRangeFormatting(3, 3, 3, limit, self.box_right_format)
        # Thick Boxes
        csqp.updateCellRangeFormatting(1, 1, 4, 1, self.box_top_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 4, 1, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(4, 1, 4, limit, self.box_right_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 1, limit, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(1, limit, 4, limit, self.box_bottom_thick_format)
        # Colors
        csqp.updateCellRangeFormatting(1, 1, 4, 1, self.purple_title_format)
        csqp.updateCellRangeFormatting(1, 2, 1, limit, self.purple2_format)
        csqp.updateCellRangeFormatting(2, 2, 2, limit, self.purple1_format)
        csqp.updateCellRangeFormatting(3, 2, 3, limit, self.purple2_format)
        csqp.updateCellRangeFormatting(4, 2, 4, limit, self.purple1_format)
        # Inputting Teams
        for team in self.teamDict:
            csqp.updateCellValue(team.o_x, team.o_y, team.name)
            csqp.updateCellValue(team.o_x + 1, team.o_y, team.number)
            csqp.updateCellValue(team.o_x + 2, team.o_y, team.elo)
        # Push the USCQP query to the sheet    
        csqp.pushCellUpdate()


class UCSQP:
    ### Universal Compressed Sheets Query Protocol ###
    ## This was created because of the query limit on Google Sheets API v4
    ## Since it's limited to 100 requests per 100 seconds, we need to use as little requests as possible
    #   This system grabs the cells we need to edit, edits them locally in python,
    #   and then pushes all of the cells back to google sheets in 1 request.
    #   This is much more efficient than using 1 request to edit one cell.
    def __init__(self, ws, r_x1, r_y1, r_x2, r_y2):
        # @param ws - imports the worksheet class to be able to write to the worksheet
        # @param r - range for the section of cells to grab
        # Class Imports
        self.ws = ws
        self.o_x = r_x1
        self.o_y = r_y1

        # Local Query List Setup
        self.cell_list = self.ws.ws.range('%s:%s' % (gspread.utils.rowcol_to_a1(r_y1, r_x1), gspread.utils.rowcol_to_a1(r_y2, r_x2))) # Grabs the whole section of cells
        self.cell_formatting = []
        self.custom_requests = {"requests": []}
        
    def updateCellValue(self, x, y, cellvalue):
        tempcell = self.convertLocal(x, y)
        # Finds the cell in the cell list and updates the value
        for cell in self.cell_list:
            if (cell.row == a1_to_rowcol(tempcell)[0] and cell.col == a1_to_rowcol(tempcell)[1]):
                cell.value = cellvalue
                return cell

    def convertLocal(self, x, y):
        # Calculate pos based on origin and converts to A1
        temp_x = self.o_x + x
        temp_y = self.o_y + y
        return gspread.utils.rowcol_to_a1(temp_y, temp_x)

    def updateCellFormatting(self, x, y, formatsettings):
        tempcell = self.convertLocal(x, y) # Converts the x, y coordinates to A1 format and local
        tempcellrange = "%s:%s" % (tempcell, tempcell) # GSpread_formatting requires a range of cells for some reason
        # Puts the formatting options into a tuple and stores it in a list
        self.cell_formatting.insert(len(self.cell_formatting), (tempcellrange, formatsettings))

    def updateCellRangeFormatting(self, x1, y1, x2, y2, formatsettings):
        # Converts the 2 cells into A1 format and then puts it into the range format for gspread_formatting
        tempcell1 = self.convertLocal(x1, y1) 
        tempcell2 = self.convertLocal(x2, y2) 
        tempcellrange = "%s:%s" % (tempcell1, tempcell2)
        # Puts the formatting options into a tuple and stores it in a list
        self.cell_formatting.insert(len(self.cell_formatting), (tempcellrange, formatsettings))

    def updateCustomCellFormatting(self, x1, y1, x2, y2, requesttype):
        # Formulates a custom google sheets request for merging cells
        if (requesttype == "merge"):
            temp_x1 = self.o_x + x1 - 1
            temp_y1 = self.o_y + y1 - 1
            temp_x2 = self.o_x + x2
            temp_y2 = self.o_y + y2 
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
    def __init__(self, scheduledict, scoredict, fake = False):
        # Class Imports
        self.schedule = scheduledict
        self.score = scoredict

        # Info about the Match Entry
        if fake == False:
            self.matchnum = self.schedule["matchNumber"]
            self.matchtype = self.schedule["tournamentLevel"]
            self.matchtitle = ("%s Match #%s %s" % (self.schedule["description"].split()[0], self.schedule["description"].split()[1], self.formatDate()))
        self.o_x = 0 # X Location of the Entry
        self.o_y = 0 # Y Location of the Entry
    def formatDate(self):
        datevalues = self.schedule["startTime"].split("T")[0].split("-")
        return ("(%s/%s/%s)" % (datevalues[1], datevalues[2], datevalues[0]))

class Team:
    # A class that represents an individual team in a match
    # One is created for each team in an event
    _c = 1
    _q = 0.0057565

    def __init__(self, teamdict):
        # Class Imports
        self.team = teamdict
        # Team Entry Origin
        self.o_x = 0
        self.o_y = 0
        # Team Information
        self.name = self.team["nameShort"]
        self.number = self.team["teamNumber"]
        # Predictions Data
        self.mitchrating = 1500
        self.ratingdeviation = 350

    @property
    def tranformed_rd(self):
        return min([350, math.sqrt(self.ratingdeviation ** 2 + self._c ** 2)])

    @classmethod   
    def _g(cls, x):
        return 1 / (math.sqrt(1 + 3 * cls._q ** 2 * (x ** 2) / math.pi ** 2))

    def expected_score(self, ooamr, ooard, inverse = False):
        #@param oaamr - Opponent Alliance Average Mitch Rating
        #@param oaard - Oponent Alliance Average Rating Deviation
        if inverse == True:
            g_term = self._g(math.sqrt(ooard ** 2 + self.ratingdeviation ** 2) * (ooamr - self.mitchrating) / 400)
            E = 1 / (1 + 10 ** (-1 * g_term))
        else:
            g_term = self._g(math.sqrt(self.ratingdeviation ** 2 + ooard ** 2) * (self.mitchrating - ooamr) / 400)
            E = 1 / (1 + 10 ** (-1 * g_term))
        return E

    def wonAgainst(self, oaamr, oaard):
        #@param oaamr - Opponent Alliance Average Mitch Rating
        #@param oaard - Oponent Alliance Average Rating Deviation
        s = 1
        E_term = self.expected_score(oaamr, oaard)
        d_squared = (self._q ** 2 * (self._g(oaard) ** 2 * E_term * (1 - E_term))) ** -1
        s_new_mitchrating = self.mitchrating + (self._q / (1 / self.ratingdeviation ** 2 + 1 / d_squared)) * self._g(oaard) * (s - E_term)
        s_new_ratingdeviation = math.sqrt((1 / self.ratingdeviation ** 2 + 1 / d_squared) ** -1)
    def lostAgainst(self, oaamr, oaard):
        #@param oaamr - Opponent Alliance Average Mitch Rating
        #@param oaard - Oponent Alliance Average Rating Deviation
        s = 0
        E_term = self.expected_score(oaamr, oaard, inverse = True)
        d_squared = (self._q ** 2 * (self._g(self.ratingdeviation) ** 2 * E_term * (1 - E_term))) ** -1
        s_new_mitchrating = oaamr + (self._q / (1 / ooard ** 2 + 1 / d_squared)) * self._g(self.ratingdeviation) * (s - E_term)




