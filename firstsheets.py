############################################################################################################################
# FirstSheets                                                                                                              #
# Used to store and retrieve data from the Google Sheets                                                                   #
# Implements CSQP to use less queries                                                                                      #
############################################################################################################################

# pip gspread, gspread_formatting, pyopenssl and oauth2client for google sheets functionality
import time
import datetime
import json
import gspread
import operator
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import *
from firstpredictions import *

class Sheets:
    def __init__(self, config, data):
        # Class Imports
        self.config = config
        self.data = data

        # Google Drive OAuth2 Setup
        # gspread.readthedocs.io/en/latest/
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.config.oauthjsonpath, scope)
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
        self.nuke_format = CellFormat(textFormat = TextFormat(bold = False, fontSize = 10, fontFamily = 'Montserrat'))

        # Sheets Config Setup
        try:
            self.config_ws = self.sh.worksheet("Home")
        except Exception as e:
            print (e + " sheet not found! Please create for Sheets Config Functionality")
            pass
        if self.checkIfSheetExists() == False:
            self.createSheet()

    def checkIfSheetExists(self):
        found = False
        worksheetsList = self.sh.worksheets()
        for worksheet in worksheetsList:
            if worksheet._properties['title'] == str(self.config.eventid):
                found = True
                self.ws = self.sh.worksheet(str(self.config.eventid))
        return found
    
    def checkManualEntry(self):
        # Function for checking data if manual data retrieval is on
        csqp = UCSQP(self.config_ws, self.sh, 1, 21, 5, 29)
        update = csqp.readCell(5, 5)
        if update != "":
            matchNum = 1
            matchtitle = csqp.readCell(2, 2)
            matchtime = str(datetime.datetime.now()).replace(" ", "T")
            print ("Match Manually Inputted!")
            if matchtitle != "":
                matchtype = matchtitle.split()[0]
            redTeam1 = csqp.readCell(2, 4)
            redTeam2 = csqp.readCell(2, 5)
            redTeam3 = csqp.readCell(2, 6)
            redSwitchLevel = csqp.readCell(2, 7)
            if redSwitchLevel == "Yes" or bool(redSwitchLevel) == True:
                redSwitchLevel = "IsLevel"
            elif redSwitchLevel == "":
                redSwitchLevel = "null"
            else:
                redSwitchLevel = "NotLevel"
            redRankingPoints = csqp.readCell(2, 8)
            redFinalScore = csqp.readCell(2, 9)
            blueTeam1 = csqp.readCell(3, 4)
            blueTeam2 = csqp.readCell(3, 5)
            blueTeam3 = csqp.readCell(3, 6)
            blueSwitchLevel = csqp.readCell(3, 7)
            if blueSwitchLevel == "Yes" or bool(blueSwitchLevel) == True:
                blueSwitchLevel = "IsLevel"
            elif blueSwitchLevel == "":
                blueSwitchLevel = "null"
            else:
                blueSwitchLevel = "NotLevel"
            blueRankingPoints = csqp.readCell(3, 8)
            blueFinalScore = csqp.readCell(3, 9)
            # For if match happened or not
            if blueFinalScore == "" or redFinalScore == "":
                postResultTime = "null"
            else:
                postResultTime = "1"
            # Event Data
            eventname = csqp.readCell(5, 3)
            eventvenue = csqp.readCell(5, 4)
            # Team Data
            teamname = csqp.readCell(5, 7)
            teamnumber = csqp.readCell(5, 8)
            if redTeam1 != "" and blueTeam1 != "":
                if matchtype == "Qualification" or matchtype == "Qualifier" or matchtype == "qualification" or matchtype == "qualifier":
                    with open('manual/manualqualscheduledata.json') as json_file:
                        qualScheduleData = json.load(json_file)
                    with open('manual/manualqualscoredata.json') as json_file:
                        qualScoreData = json.load(json_file)
                    # Sets the match number
                    if len(qualScoreData) != 0:
                        matchNum = qualScoreData[-1]["matchNumber"] + 1
                    tempMatchEntryScore = {
                                    "alliances": [
                                        {
                                            "alliance": "Blue",
                                            "autoCellsBottom": 0,
                                            "autoCellsOuter": 0,
                                            "autoCellsInner": 0,
                                            "teleopCellsBottom": 0,
                                            "teleopCellsOuter": 0,
                                            "teleopCellsInner": 0,
                                            "endgameRungIsLevel": blueSwitchLevel,
                                            "autoInitLinePoints": 0,
                                            "autoCellPoints": 0,
                                            "autoPoints": 0,
                                            "teleopCellPoints": 0,
                                            "controlPanelPoints": 0,
                                            "endgamePoints": 0,
                                            "teleopPoints": "69",
                                            "shieldOperationalRankingPoint":"False",
                                            "shieldEnergizedRankingPoint":"False",
                                            "foulCount":"0",
                                            "techFoulCount":"0",
                                            "adjustPoints":"0",
                                            "foulPoints":"0",
                                            "rp":blueRankingPoints,
                                            "totalPoints": blueFinalScore
                                        }, 
                                        {
                                            "alliance": "Red",
                                            "autoCellsBottom": 0,
                                            "autoCellsOuter": 0,
                                            "autoCellsInner": 0,
                                            "teleopCellsBottom": 0,
                                            "teleopCellsOuter": 0,
                                            "teleopCellsInner": 0,
                                            "endgameRungIsLevel": redSwitchLevel,
                                            "autoInitLinePoints": 0,
                                            "autoCellPoints": 0,
                                            "autoPoints": 0,
                                            "teleopCellPoints": 0,
                                            "controlPanelPoints": 0,
                                            "endgamePoints": 0,
                                            "teleopPoints": "69",
                                            "shieldOperationalRankingPoint":"False",
                                            "shieldEnergizedRankingPoint":"False",
                                            "foulCount":"0",
                                            "techFoulCount":"0",
                                            "adjustPoints":"0",
                                            "foulPoints":"0",
                                            "rp":redRankingPoints,
                                            "totalPoints": redFinalScore
                                        }
                                    ], 
                                    "matchLevel": "Qualification", 
                                    "matchNumber": matchNum
                                    }
                    tempMatchEntrySchedule = {
                                        "description": matchtitle, 
                                        "matchNumber": matchNum, 
                                        "postResultTime": postResultTime, 
                                        "scoreBlueAuto": 0, 
                                        "scoreBlueFinal": blueFinalScore, 
                                        "scoreBlueFoul": 0, 
                                        "scoreRedAuto": 0, 
                                        "scoreRedFinal": redFinalScore, 
                                        "scoreRedFoul": 0, 
                                        "startTime": matchtime, 
                                        "teams": [
                                            { 
                                                "station": "Red1",  
                                                "teamNumber": redTeam1
                                            }, 
                                            { 
                                                "station": "Red2",  
                                                "teamNumber": redTeam2
                                            }, 
                                            { 
                                                "station": "Red3",  
                                                "teamNumber": redTeam3
                                            }, 
                                            { 
                                                "station": "Blue1",  
                                                "teamNumber": blueTeam1
                                            }, 
                                            { 
                                                "station": "Blue2",  
                                                "teamNumber": blueTeam2
                                            }, 
                                            { 
                                                "station": "Blue3",  
                                                "teamNumber": blueTeam3
                                            }
                                        ], 
                                        "tournamentLevel": "Qualification"
                    }
                    # If the match already exists, simply append it instead of adding it again
                    qualDataMatches = False
                    for match in qualScheduleData:
                        if match["description"] == matchtitle:
                            qualDataMatches = True
                            tempmatchnum = match["matchNumber"]
                            match = tempMatchEntrySchedule
                            for matchScore in qualScoreData:
                                if tempmatchnum == matchScore["matchNumber"]:
                                    tempMatchEntryScore['matchNumber'] = tempmatchnum
                                    match = tempMatchEntryScore
                    if qualDataMatches == False:
                        # If it doesn't find the entry, add them to the end
                        qualScoreData.insert(len(qualScoreData), tempMatchEntryScore)
                        qualScheduleData.insert(len(qualScheduleData), tempMatchEntrySchedule)
                    # Update the JSON files             
                    with open('manual/manualqualscheduledata.json', 'w') as json_file:
                        json.dump(qualScheduleData, json_file)
                    with open('manual/manualqualscoredata.json', 'w') as json_file:
                        json.dump(qualScoreData, json_file)
                else:
                    with open('manual/manualplayoffscheduledata.json') as json_file:
                        playoffScheduleData = json.load(json_file)
                    with open('manual/manualplayoffscoredata.json') as json_file:
                        playoffScoreData = json.load(json_file)
                    # Sets the match number
                    if len(playoffScoreData) == 0:
                        matchNum = qualScoreData[-1]["matchNumber"] + 1
                    else:
                        matchNum = playoffScoreData[-1]["matchNumber"] + 1
                    tempMatchEntryScore = {
                                    "alliances": [
                                        {
                                            "alliance": "Blue",
                                            "autoCellsBottom": 0,
                                            "autoCellsOuter": 0,
                                            "autoCellsInner": 0,
                                            "teleopCellsBottom": 0,
                                            "teleopCellsOuter": 0,
                                            "teleopCellsInner": 0,
                                            "endgameRungIsLevel": blueSwitchLevel,
                                            "autoInitLinePoints": 0,
                                            "autoCellPoints": 0,
                                            "autoPoints": 0,
                                            "teleopCellPoints": 0,
                                            "controlPanelPoints": 0,
                                            "endgamePoints": 0,
                                            "teleopPoints": "69",
                                            "shieldOperationalRankingPoint":"False",
                                            "shieldEnergizedRankingPoint":"False",
                                            "foulCount":"0",
                                            "techFoulCount":"0",
                                            "adjustPoints":"0",
                                            "foulPoints":"0",
                                            "rp":blueRankingPoints,
                                            "totalPoints": blueFinalScore
                                        }, 
                                        {
                                            "alliance": "Red",
                                            "autoCellsBottom": 0,
                                            "autoCellsOuter": 0,
                                            "autoCellsInner": 0,
                                            "teleopCellsBottom": 0,
                                            "teleopCellsOuter": 0,
                                            "teleopCellsInner": 0,
                                            "endgameRungIsLevel": redSwitchLevel,
                                            "autoInitLinePoints": 0,
                                            "autoCellPoints": 0,
                                            "autoPoints": 0,
                                            "teleopCellPoints": 0,
                                            "controlPanelPoints": 0,
                                            "endgamePoints": 0,
                                            "teleopPoints": "69",
                                            "shieldOperationalRankingPoint":"False",
                                            "shieldEnergizedRankingPoint":"False",
                                            "foulCount":"0",
                                            "techFoulCount":"0",
                                            "adjustPoints":"0",
                                            "foulPoints":"0",
                                            "rp":redRankingPoints,
                                            "totalPoints": redFinalScore
                                        }
                                    ], 
                                    "matchLevel": "Playoff", 
                                    "matchNumber": matchNum
                                    }
                    tempMatchEntrySchedule = {
                                        "description": matchtitle, 
                                        "matchNumber": matchNum, 
                                        "postResultTime": postResultTime, 
                                        "scoreBlueAuto": 0, 
                                        "scoreBlueFinal": blueFinalScore, 
                                        "scoreBlueFoul": 0, 
                                        "scoreRedAuto": 0, 
                                        "scoreRedFinal": redFinalScore, 
                                        "scoreRedFoul": 0, 
                                        "startTime": matchtime, 
                                        "teams": [
                                            { 
                                                "station": "Red1",  
                                                "teamNumber": redTeam1
                                            }, 
                                            { 
                                                "station": "Red2",  
                                                "teamNumber": redTeam2
                                            }, 
                                            { 
                                                "station": "Red3",  
                                                "teamNumber": redTeam3
                                            }, 
                                            { 
                                                "station": "Blue1",  
                                                "teamNumber": blueTeam1
                                            }, 
                                            { 
                                                "station": "Blue2",  
                                                "teamNumber": blueTeam2
                                            }, 
                                            { 
                                                "station": "Blue3",  
                                                "teamNumber": blueTeam3
                                            }
                                        ], 
                                        "tournamentLevel": "Playoff"
                    }
                    # If the match already exists, simply append it instead of adding it again
                    playoffDataMatches = False
                    for match in playoffScheduleData:
                        if match["description"] == matchtitle:
                            playoffDataMatches = True
                            tempmatchnum = match["matchNumber"]
                            match = tempMatchEntrySchedule
                            for matchScore in playoffScoreData:
                                if tempmatchnum == matchScore["matchNumber"]:
                                    tempMatchEntryScore['matchNumber'] = tempmatchnum
                                    match = tempMatchEntryScore
                    if playoffDataMatches == False:
                        # If it doesn't find the entry, add them to the end
                        playoffScoreData.insert(len(qualScoreData), tempMatchEntryScore)
                        playoffScheduleData.insert(len(qualScheduleData), tempMatchEntrySchedule)
                    # Update the JSON files
                    with open('manual/manualplayoffscheduledata.json', 'w') as json_file:
                        json.dump(playoffScheduleData, json_file)
                    with open('manual/manualplayoffscoredata.json', 'w') as json_file:
                        json.dump(playoffScoreData, json_file)
            # Add the event data if its not blank
            if eventvenue != "" and eventname != "":
                tempEvent = {
                    "name": eventname, 
                    "venue": eventvenue, 
                }
                with open('manual/manualeventdata.json', 'w') as json_file:
                    json.dump(tempEvent, json_file)
            # Add the team data if its not blank
            if teamname != "" and teamnumber != "":
                with open('manual/manualteamdata.json') as json_file:
                    teamData = json.load(json_file)
                tempTeam = {
                    "nameShort": teamname, 
                    "teamNumber": int(teamnumber), 
                }
                teamData.insert(len(teamData), tempTeam)
                with open('manual/manualteamdata.json', 'w') as json_file:
                    json.dump(teamData, json_file)
            # Wipes the manual data screen
            csqp.updateCellValue(5, 5, "")
            csqp.updateCellValue(2, 2, "")
            csqp.updateCellValue(2, 4, "")
            csqp.updateCellValue(2, 5, "")
            csqp.updateCellValue(2, 6, "")
            csqp.updateCellValue(2, 7, "")
            csqp.updateCellValue(2, 8, "")
            csqp.updateCellValue(2, 9, "")
            csqp.updateCellValue(3, 4, "")
            csqp.updateCellValue(3, 5, "")
            csqp.updateCellValue(3, 6, "")
            csqp.updateCellValue(3, 7, "")
            csqp.updateCellValue(3, 8, "")
            csqp.updateCellValue(3, 9, "")
            csqp.updateCellValue(5, 3, "")
            csqp.updateCellValue(5, 4, "")
            csqp.updateCellValue(5, 7, "")
            csqp.updateCellValue(5, 8, "")
            csqp.pushCellUpdate()
    
    def createSheet(self):
        # Creates the worksheet
        self.ws = self.sh.add_worksheet(title=str(self.config.eventid), rows="3000", cols="30")
        # Creates a CSQP object to setup the worksheet
        csqp = UCSQP(self.ws, self.sh, 1, 1, 2, 30)
        # Creates formatting for the title and competition location
        csqp.updateCellFormatting(1, 1, self.title_format)
        csqp.updateCellFormatting(1, 2, self.venue_format)
        # Adds the event info into the first 2 rows
        csqp.updateCellValue(1, 1, self.data.eventData["name"])
        csqp.updateCellValue(1, 2, self.data.eventData["venue"])
        # Setting the cell size to 136 for all cells
        csqp.updateCustomCellFormatting(1, 0 , 30, 0, "resizecolumn")
        # Changing the font for all cells
        csqp.updateCellRangeFormatting(1, 1, 30, 3000, self.font_format)
        csqp.pushCellUpdate()
        print("Worksheet Created: " + str(self.config.eventid))

    def convertLocal(self, x, y, match):
        # Calculate pos based on origin and converts to A1
        temp_x = match.o_x + x - 1
        temp_y = match.o_y + y - 1
        return gspread.utils.rowcol_to_a1(temp_y, temp_x)
    
    def checkIfMatchExists(self, matchToCheck):
        for match in self.matchList:
            if match.matchtitle == matchToCheck.matchtitle:
                return True
        return False

    def updateMatchIfChanged(self, matchToCheck):
        for match in self.matchList:
            # Finds the match
            if match.matchtitle == matchToCheck.matchtitle:
                # Checks if the data has been changed
                if matchToCheck.schedule["postResultTime"] != match.schedule["postResultTime"]:
                    # Stores the notes so they don't get deleted, swaps the match with the new updated one, and puts the notes back
                    tempNotes = match.notes
                    match = matchToCheck
                    match.notes = tempNotes
                    return True
        return False

    def createMatchObjects(self):
        # Creates lists to store the qualifier and playoff matches
        self.matchList = []

        # Qualifier Match Object Creation
        for x in range(len(self.data.qualScheduleData)):
            tempMatch = Match(self.data.qualScheduleData[x], self.data.qualScoreData[x])
            if self.checkIfMatchExists(tempMatch) == False:
                if (x == 0):
                    tempMatch.o_y = 4 # If it's the very first match, set the origin to the fourth row
                else:
                    tempMatch.o_y = self.matchList[-1].o_y + 18
                # Creates the entry in sheets
                # Updates the Mitch Score of the teams involved in the match if it's happened
                if tempMatch.matchhappened:
                    tempMatch.updateTeamScores(self.teamDict)
                print ("Match Entry Created: ", tempMatch.matchtitle)
                self.matchList.insert(len(self.matchList), tempMatch)
            else:
                # If the match does exist, make sure it hasn't changed
                # If it has, update it to the new data
                self.updateMatchIfChanged(tempMatch)

        # Playoff Match Object Creation
        for x in range(len(self.data.playoffScheduleData)):
            tempMatch = Match(self.data.playoffScheduleData[x], self.data.playoffScoreData[x])
            if self.checkIfMatchExists(tempMatch) == False:
                # In case for some reason theres no qualifier matches, set the origin to 4
                if len(self.data.qualScheduleData) == 0:
                    tempMatch.o_y = 4
                else:
                    tempMatch.o_y = self.matchList[-1].o_y + 18
                # Creates the entry in sheets
                # Updates the Mitch Score of the teams involved in the match if it's happened
                if tempMatch.matchhappened:
                    tempMatch.updateTeamScores(self.teamDict)
                print ("Match Entry Created: ", tempMatch.matchtitle)
                self.matchList.insert(len(self.matchList), tempMatch)
            else:
                # If the match does exist, make sure it hasn't changed
                # If it has, update it to the new data
                self.updateMatchIfChanged(tempMatch)
        # Creates a seperate list of matches that are filtered to the config settings
        self.filterMatches()
        
    def filterMatches(self):
        # Lists for teams to filter and a list of filtered matches
        self.filteredMatchList = []
        teamFilter = []
        if (self.config.matchteamfilter != "ALL" and self.config.matchteamfilter != ""):
            teamFilter = self.config.matchteamfilter.split(',')
            for team in teamFilter: 
                try:
                    intnum = int(team)
                    team = intnum
                except:
                    # Sets the team filter to blank so it doesnt crash the program
                    teamFilter = []
                    print("Invalid text in teamfilter: %s" % team)
                    break
        for match in self.matchList:
            # Conditionals on whether the match should be added to the list or not
            teamMatchFilter = False
            unplayedMatchFilter = False
            qualifierMatchFilter = False
            # Team Filter
            if len(teamFilter) == 0:
                # If the teamfilter is disabled, always add the match
                teamMatchFilter = True
            else:
                # Otherwise, make sure the team is in this match
                for team in match.redTeamList:
                    for teamnum in teamFilter:
                        if int(team) == int(teamnum):
                            teamMatchFilter = True
                for team in match.blueTeamList:
                    for teamnum in teamFilter:
                        if int(team) == int(teamnum):
                            teamMatchFilter = True
            # Unplayed Matches Filter
            if (self.config.displayunplayedmatches == "TRUE"):
                # If the filter is enabled, display all matches
                unplayedMatchFilter = True
            else:
                # If it's off, only display played matches
                if match.matchhappened:
                    unplayedMatchFilter = True
            # Qualifier Match Filter
            if (self.config.displayqualifiermatches == "TRUE"):
                # If the filter is on, display all matches
                qualifierMatchFilter = True
            else:
                # If the filter is off, only display Playoff Matches
                if match.score['matchLevel'] == "Playoff":
                    qualifierMatchFilter = True
            # Add the match if all 3 qualifiers are met
            if teamMatchFilter and unplayedMatchFilter and qualifierMatchFilter:
                self.filteredMatchList.insert(len(self.filteredMatchList), match)
        # Resets the origins so that they are underneath each other
        posCounter = 4
        for match in self.filteredMatchList:
            match.o_y = posCounter # If it's the very first match, set the origin to the fourth row
            posCounter += 18

    def grabNotes(self, csqp):
        # Goes through the currently displayed matches and grabs their notes so they don't get wiped
        for x in range(4, (len(self.matchList) + 5), 18):
            csqp.updateOrigin(1, x)
            for match in self.matchList:
                if match.matchtitle == csqp.readCell(1, 1):
                    match.notes = csqp.readCell(1, 15)

    def nukeMatchEntries(self):
        # Deletes all of the match entries, used for filters
        grabYLimit = (len(self.matchList) * 18) + 2
        csqp = UCSQP(self.ws, self.sh, self.matchList[0].o_x, self.matchList[0].o_y, self.matchList[0].o_x + 5, grabYLimit)

        # Makes sure the notes don't get deleted
        self.grabNotes(csqp)

        # Deletes all of the entries
        csqp.updateCellRangeFormatting(1, 1, 6, grabYLimit - 2, self.nuke_format)
        csqp.updateCustomCellFormatting(1, 1, 6, grabYLimit - 2, "unmerge")
        csqp.nukeValues()

        # Pushes the nuke to the sheet
        csqp.pushCellUpdate()

    def createMatchEntries(self):
        ### Creates all of the match entries
        
        ## Compressed Sheets Query Protocol Setup
        grabXLimit = (len(self.matchList) * 18) + 2
        csqp = UCSQP(self.ws, self.sh, self.matchList[0].o_x, self.matchList[0].o_y, self.matchList[0].o_x + 5, grabXLimit)

        # Makes sure the notes don't get deleted
        self.grabNotes(csqp)

        # Sets the filters
        self.filterMatches()

        # Actual match entry creation, uses one CSQP
        for match in self.filteredMatchList:
            csqp.updateOrigin(match.o_x, match.o_y)
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
            ## Predictions Data
            if match.matchhappened:
                # Updates all the MitchRating and Score Predictions with the stored ones
                # Red Team
                csqp.updateCellValue(5, 3, match.redteam_mr[str(match.schedule["teams"][0]["teamNumber"])])
                csqp.updateCellValue(5, 4, match.redteam_mr[str(match.schedule["teams"][1]["teamNumber"])])
                csqp.updateCellValue(5, 5, match.redteam_mr[str(match.schedule["teams"][2]["teamNumber"])])
                csqp.updateCellValue(5, 6, match.redteam_avmr)
                csqp.updateCellValue(5, 7, match.redteam_avscr)
                # Blue Team
                csqp.updateCellValue(6, 3, match.blueteam_mr[str(match.schedule["teams"][3]["teamNumber"])])
                csqp.updateCellValue(6, 4, match.blueteam_mr[str(match.schedule["teams"][4]["teamNumber"])])
                csqp.updateCellValue(6, 5, match.blueteam_mr[str(match.schedule["teams"][5]["teamNumber"])])
                csqp.updateCellValue(6, 6, match.blueteam_avmr)
                csqp.updateCellValue(6, 7, match.blueteam_avscr)
            else:
                # Makes predictions based off previous match data
                csqp.updateCellValue(5, 3, self.teamDict[str(match.schedule["teams"][0]["teamNumber"])].mitchrating)
                csqp.updateCellValue(5, 4, self.teamDict[str(match.schedule["teams"][1]["teamNumber"])].mitchrating)
                csqp.updateCellValue(5, 5, self.teamDict[str(match.schedule["teams"][2]["teamNumber"])].mitchrating)
                tempredteam_avmr = (self.teamDict[str(match.schedule["teams"][0]["teamNumber"])].mitchrating + self.teamDict[str(match.schedule["teams"][1]["teamNumber"])].mitchrating + self.teamDict[str(match.schedule["teams"][2]["teamNumber"])].mitchrating) / 3
                csqp.updateCellValue(5, 6, tempredteam_avmr)
                if self.teamDict[str(match.schedule["teams"][0]["teamNumber"])].matchesPlayed != 0 and self.teamDict[str(match.schedule["teams"][1]["teamNumber"])].matchesPlayed != 0 and self.teamDict[str(match.schedule["teams"][2]["teamNumber"])].matchesPlayed != 0:
                    tempredteam_avscr = (((self.teamDict[str(match.schedule["teams"][0]["teamNumber"])].totalScore // self.teamDict[str(match.schedule["teams"][0]["teamNumber"])].matchesPlayed) + (self.teamDict[str(match.schedule["teams"][1]["teamNumber"])].totalScore // self.teamDict[str(match.schedule["teams"][1]["teamNumber"])].matchesPlayed) + (self.teamDict[str(match.schedule["teams"][2]["teamNumber"])].totalScore // self.teamDict[str(match.schedule["teams"][2]["teamNumber"])].matchesPlayed)) // 3)
                else:
                    tempredteam_avscr = 0
                csqp.updateCellValue(5, 7, tempredteam_avscr)
                # Blue Team
                csqp.updateCellValue(6, 3, self.teamDict[str(match.schedule["teams"][3]["teamNumber"])].mitchrating)
                csqp.updateCellValue(6, 4, self.teamDict[str(match.schedule["teams"][4]["teamNumber"])].mitchrating)
                csqp.updateCellValue(6, 5, self.teamDict[str(match.schedule["teams"][5]["teamNumber"])].mitchrating)
                tempblueteam_avmr = ((self.teamDict[str(match.schedule["teams"][3]["teamNumber"])].mitchrating + self.teamDict[str(match.schedule["teams"][4]["teamNumber"])].mitchrating + self.teamDict[str(match.schedule["teams"][5]["teamNumber"])].mitchrating) // 3)
                csqp.updateCellValue(6, 6, tempblueteam_avmr)
                if self.teamDict[str(match.schedule["teams"][3]["teamNumber"])].matchesPlayed != 0 and self.teamDict[str(match.schedule["teams"][4]["teamNumber"])].matchesPlayed != 0 and self.teamDict[str(match.schedule["teams"][5]["teamNumber"])].matchesPlayed != 0:
                    tempblueteam_avscr = (((self.teamDict[str(match.schedule["teams"][3]["teamNumber"])].totalScore // self.teamDict[str(match.schedule["teams"][3]["teamNumber"])].matchesPlayed) + (self.teamDict[str(match.schedule["teams"][4]["teamNumber"])].totalScore // self.teamDict[str(match.schedule["teams"][4]["teamNumber"])].matchesPlayed) + (self.teamDict[str(match.schedule["teams"][5]["teamNumber"])].totalScore // self.teamDict[str(match.schedule["teams"][5]["teamNumber"])].matchesPlayed)) // 3)
                else:
                    tempblueteam_avscr = 0
                csqp.updateCellValue(6, 7, tempblueteam_avscr)
            
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
            csqp.updateCellValue(1, 15, match.notes)
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
            if match.matchhappened:
                csqp.updateCellRangeFormatting(1, 2, 1, 13, self.match_played_format)
                csqp.updateCellRangeFormatting(1, 1, 6, 1, self.match_played_format)
                csqp.updateCellRangeFormatting(4, 2, 4, 13, self.match_played_format)
                csqp.updateCellRangeFormatting(1, 14, 6, 17, self.match_played_format)
            else:
                csqp.updateCellRangeFormatting(1, 2, 1, 13, self.match_scheduled_format)
                csqp.updateCellRangeFormatting(1, 1, 6, 1, self.match_scheduled_format)
                csqp.updateCellRangeFormatting(4, 2, 4, 13, self.match_scheduled_format)
                csqp.updateCellRangeFormatting(1, 14, 6, 17, self.match_scheduled_format)

        # Pushes all of the cells to google sheets
        csqp.pushCellUpdate()

    def createTeamObjects(self):
        self.teamDict = {}
        # Temporary location storage for origin creation
        templocationholdy = 0
        for x in range(len(self.data.teamData)):
            tempTeam = Team(self.data.teamData[x])
            if x == 0:
                tempTeam.o_y = 3 # If it's the very first match, set the origin below the title and column defs
                templocationholdy = 3
            else:
                templocationholdy += 1
                tempTeam.o_y = templocationholdy
            print ("Team Entry Created: %s" % tempTeam.name)
            self.teamDict[str(tempTeam.number)] = tempTeam

    def createTeamEntry(self):
        csqp = UCSQP(self.ws, self.sh, 8, 4, 14, (5 + len(self.teamDict)))
        # Checks to see if the match entry is already there
        alreadyExists = csqp.readCell(1, 1) == "Team List"
        # Takes all the team entries and stores their notes, robot weight and robot height so they don't get wiped
        # Only happens if the list already exists
        if alreadyExists:
            for x in range(3, (3 + len(self.teamDict))):
                teamNum = str(csqp.readCell(2, x))
                self.teamDict[teamNum].robotType = csqp.readCell(5, x)
                self.teamDict[teamNum].robotWeight = csqp.readCell(6, x)
                self.teamDict[teamNum].notes = csqp.readCell(7, x)

        # Team Entry Creation
        limit = len(self.teamDict) + 2 # +2 adds space for the title and column names
        # Only creates the table and formatting if the entry isn't there
        if not alreadyExists:
            # Title
            csqp.updateCellFormatting(1, 1, self.matchtitle_format)
            csqp.updateCellValue(1, 1, "Team List")
            # Centered Values
            csqp.updateCellRangeFormatting(1, 2, 7, 2, self.centered_format)
            csqp.updateCellRangeFormatting(3, 3, 3, limit, self.centered_format)
            # Columns
            csqp.updateCellValue(1, 2, "Team Name")
            csqp.updateCellValue(2, 2, "Team Number")
            csqp.updateCellValue(3, 2, ("MitchRating" + u"\u2122"))
            csqp.updateCellValue(4, 2, "Rank Title")
            csqp.updateCellValue(5, 2, "Robot Type")
            csqp.updateCellValue(6, 2, "Robot Weight")
            csqp.updateCellValue(7, 2, "Notes")
            # Boxes
            csqp.updateCellRangeFormatting(1, 2, 7, 2, self.box_format)
            csqp.updateCellRangeFormatting(1, 3, 1, limit, self.box_right_format)
            csqp.updateCellRangeFormatting(2, 3, 2, limit, self.box_right_format)
            csqp.updateCellRangeFormatting(3, 3, 3, limit, self.box_right_format)
            csqp.updateCellRangeFormatting(4, 3, 4, limit, self.box_right_format)
            csqp.updateCellRangeFormatting(5, 3, 5, limit, self.box_right_format)
            csqp.updateCellRangeFormatting(6, 3, 6, limit, self.box_right_format)
            # Thick Boxes
            csqp.updateCellRangeFormatting(1, 1, 7, 1, self.box_top_thick_format)
            csqp.updateCellRangeFormatting(1, 1, 7, 1, self.box_bottom_thick_format)
            csqp.updateCellRangeFormatting(7, 1, 7, limit, self.box_right_thick_format)
            csqp.updateCellRangeFormatting(1, 1, 1, limit, self.box_left_thick_format)
            csqp.updateCellRangeFormatting(1, limit, 7, limit, self.box_bottom_thick_format)
            # Colors
            csqp.updateCellRangeFormatting(1, 1, 7, 1, self.purple_title_format)
            csqp.updateCellRangeFormatting(1, 2, 1, limit, self.purple2_format)
            csqp.updateCellRangeFormatting(2, 2, 2, limit, self.purple1_format)
            csqp.updateCellRangeFormatting(3, 2, 3, limit, self.purple2_format)
            csqp.updateCellRangeFormatting(4, 2, 4, limit, self.purple1_format)
            csqp.updateCellRangeFormatting(5, 2, 5, limit, self.purple2_format)
            csqp.updateCellRangeFormatting(6, 2, 6, limit, self.purple1_format)
            csqp.updateCellRangeFormatting(7, 2, 7, limit, self.purple2_format)
            # Name Column Resize
            csqp.updateCustomCellFormatting(1, 0, 2, 0, "resizecolumn", columnsize = 300)
            csqp.updateCustomCellFormatting(7, 0, 8, 0, "resizecolumn", columnsize = 500)

        # Sorts the dictionary by MitchRating
        sortedTeamList = []
        # Resets the origins to the order in the list
        posCounter = 3
        if (self.config.teamsort == "TRUE"):
            for team in (sorted(self.teamDict.values(), key=operator.attrgetter('mitchrating'), reverse = True)):
                team.o_y = posCounter
                posCounter += 1
                sortedTeamList.insert(len(sortedTeamList), team)
        else:
            # If it's not enabled in config, sort the teams by teamnumber
            for team in (sorted(self.teamDict.values(), key=operator.attrgetter('number'))):
                team.o_y = posCounter
                posCounter += 1
                sortedTeamList.insert(len(sortedTeamList), team)
        # Inputting Team Data
        for value in sortedTeamList:
            csqp.updateCellValue(value.o_x, value.o_y, value.name)
            csqp.updateCellValue(value.o_x + 1, value.o_y, value.number)
            csqp.updateCellValue(value.o_x + 2, value.o_y, value.mitchrating)
            # csqp.updateCellValue(value.o_x + 3, value.o_y, value.getRankTitle())
            csqp.updateCellValue(value.o_x + 4, value.o_y, value.robotType)
            csqp.updateCellValue(value.o_x + 5, value.o_y, value.robotWeight)
            csqp.updateCellValue(value.o_x + 6, value.o_y, value.notes)
        # Push the USCQP query to the sheet    
        csqp.pushCellUpdate()


class UCSQP:
    ### Universal Compressed Sheets Query Protocol ###
    ## This was created because of the query limit on Google Sheets API v4
    ## Since it's limited to 100 requests per 100 seconds, we need to use as little requests as possible
    #   This system grabs the cells we need to edit, edits them locally in python,
    #   and then pushes all of the cells back to google sheets in 1 request.
    #   This is much more efficient than using 1 request to edit one cell.
    def __init__(self, ws, sh, r_x1, r_y1, r_x2, r_y2):
        # @param ws - imports the worksheet class to be able to write to the worksheet
        # @param r - range for the section of cells to grab
        # Class Imports
        self.ws = ws
        self.sh = sh
        self.o_x = r_x1
        self.o_y = r_y1
        self.o2_x = r_x2
        self.o2_y = r_y2

        # Local Query List Setup
        self.cell_list = self.ws.range('%s:%s' % (gspread.utils.rowcol_to_a1(r_y1, r_x1), gspread.utils.rowcol_to_a1(r_y2, r_x2))) # Grabs the whole section of cells
        self.cell_formatting = []
        self.custom_requests = {"requests": []}
    
    def findCell(self, tempcell): 
        for cell in self.cell_list:
            if (cell.row == a1_to_rowcol(tempcell)[0] and cell.col == a1_to_rowcol(tempcell)[1]):
                return cell
        
    def updateCellValue(self, x, y, cellvalue):
        tempcell = self.convertLocal(x, y)
        # Finds the cell in the cell list and updates the value
        try:
            self.findCell(tempcell).value = cellvalue
        except Exception as e:
            print("Could Not Find Cell: " + str(e))
            
    def convertLocal(self, x, y):
        # Calculate pos based on origin and converts to A1
        temp_x = self.o_x + x - 1
        temp_y = self.o_y + y - 1
        return gspread.utils.rowcol_to_a1(temp_y, temp_x)

    def convertX(self, x):
        # Calculate x coordinate based on origin
        temp_x = self.o_x + x - 1
        return temp_x

    def convertY(self, y):
        temp_y = self.o_y + y - 1
        return temp_y

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

    def updateCustomCellFormatting(self, x1, y1, x2, y2, requesttype, columnsize = 136):
        # Formulates a custom google sheets request for merging cells
        if (requesttype == "merge"):
            # Does a merge all on the range of cells, used for notes section on the sheet
            temp_x1 = self.convertX(x1) - 1
            temp_y1 = self.convertY(y1) - 1
            temp_x2 = self.convertX(x2)
            temp_y2 = self.convertY(y2)
            temprequest = {
                            "mergeCells": {
                                "mergeType": "MERGE_ALL",
                                "range": {
                                    "sheetId": self.ws._properties['sheetId'],
                                    "startRowIndex": temp_y1,
                                    "endRowIndex": temp_y2,
                                    "startColumnIndex": temp_x1,
                                    "endColumnIndex": temp_x2
                                } 
                            }
                        }
            self.custom_requests["requests"].insert(len(self.custom_requests), temprequest)
        elif (requesttype == "unmerge"):
            # Does a merge all on the range of cells, used for notes section on the sheet
            temp_x1 = self.convertX(x1) - 1
            temp_y1 = self.convertY(y1) - 1
            temp_x2 = self.convertX(x2)
            temp_y2 = self.convertY(y2)
            temprequest = {
                            "unmergeCells": {
                                "range": {
                                    "sheetId": self.ws._properties['sheetId'],
                                    "startRowIndex": temp_y1,
                                    "endRowIndex": temp_y2,
                                    "startColumnIndex": temp_x1,
                                    "endColumnIndex": temp_x2
                                } 
                            }
                        }
            self.custom_requests["requests"].insert(len(self.custom_requests), temprequest)
        elif (requesttype == "resizecolumn"):
            # Resizes the column to the given size, must use @param columnsize
            columnsizeparam = columnsize
            temprequest = {
                            "updateDimensionProperties": {
                                "range": {
                                    "sheetId": self.ws._properties['sheetId'],
                                    "dimension": "COLUMNS",
                                    "startIndex": self.convertX(x1) - 1,
                                    "endIndex": self.convertX(x2) - 1
                                },
                                "properties": {
                                    "pixelSize": columnsizeparam
                                },
                                "fields": "pixelSize"
                            }
                        }
            self.custom_requests["requests"].insert(len(self.custom_requests), temprequest)
        else:
            print("Custom Request Type Not Found!")

    def readCell(self, x, y):
        tempcell = self.convertLocal(x, y)
        # Finds the cell in the cell list and returns the value
        try:
            return self.findCell(tempcell).value
        except Exception as e:
            print("Could Not Find Cell: " + str(e))

    def nukeValues(self):
        # Sets the values and formatting of all the cells to blank
        temprequest = {
                        "updateCells": {
                            "range": {
                                "sheetId": self.ws._properties['sheetId'],
                                "startRowIndex": self.o_y - 1,
                                "endRowIndex": self.o2_y,
                                "startColumnIndex": self.o_x - 1,
                                "endColumnIndex": self.o2_x
                            },
                            "fields": "userEnteredFormat"
                        }
                    }
        self.custom_requests["requests"].insert(len(self.custom_requests), temprequest)
        # Value Clearing
        for cell in self.cell_list:
            cell.value = ""

    def updateOrigin(self, x, y):
        # Use this to update the origin of where the UCSQP places cells
        self.o_x = x
        self.o_y = y
    
    def updateList(self):
        # Use this when you want to re-grab the data from the sheet
        self.cell_list = self.ws.range('%s:%s' % (gspread.utils.rowcol_to_a1(self.o_y, self.o_x), gspread.utils.rowcol_to_a1(self.o2_y, self.o2_x)))

    def pushCellUpdate(self):
        # Pushes the array created by updateCellFormatting to the google sheet
        if not (len(self.cell_formatting) == 0):
            format_cell_ranges(self.ws, self.cell_formatting)
        # Pushes cell merge requests
        if not (len(self.custom_requests["requests"]) == 0):
            self.sh.batch_update(self.custom_requests)
        # Pushes the cell value list to the google sheet
        if not (len(self.cell_list) == 0):
            self.ws.update_cells(self.cell_list)

