############################################################################################################################
# MMMMMMMMMMMMMMMMWKOOKNMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMMMMMMMMWXOkdlxXWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMMMMMMMWXOkdc;:oKWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMMMMMMWXOkdl;;;;oKWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMMMMMWXOkdc;;;;;;l0WMMMMMMMMMMMMMMMMMMMMWNWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMMMMWXOkxl;;;:c:;;lOWMMMMMMMMMMMMMMMMMMXkloONMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMMMMXOkxl;;;ckOo:;;ckNMMMMMMMMMMMMMMWXkl;,,;lONMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMMMXOkxl;;;cOWW0o:;;:kNMMMMMMMMMMMMNOl;,,,,,,;lONMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMWXOkxl;;;cOWMMWKd:;;:xXWWWWWWMMMNOl;,,,:ll:,,,;lkXWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMMXOkxl;;;cOWMMMMWKdc;;:okOOOO000kl;,,,:okOkxo:;,,;ckXWWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMMXOkxl;;;cONMMMMWX0kdc;;;cx000ko:;,,,:lkXWNXOkxo:;,,,coxXWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMMXOkxl;;;cOWMMMWX0kO0Kkl;;;l0N0o;,,,;ld0WMMMMWXOkxoc;,,,,cxKWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMMXOkxl;;;cONMMMMXO0XNNX0dc;;;cl:,,,;lk000NMMMMMMWX0kxoc;,,,,:xKWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMMXOkxl;;;cONMMMMWKkOkxdlcc:;;;;;,,,:dKWMWXNMMMMMMMMWX0kdc;,,,,,lKMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMMXOkxl:;;:xKKXWMMW0l::;;;;;;;;:::;,,,ckXWMMMMMMWWMMMMMXxc;,,,,;lONMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MMXOkxl:;;;cllcxNMMMNx:;;:codxkkxxdl:,,,;ckNMMMMWXXWMMN0o;,,,,;lkKWMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM #
# MNOkxl:;;;;;;;;lONMMMN0O0KNNWMMWNKOkdl:,;lkNMMMWK0XMN0o;,,,;:cx0OkkkkkkkkOXXOkkk0NXOkkkkkkkkOKWMWN0kxdddxOXWKOOOOOOOOOOO #
# NOkxl:;;;;;:cloxOOKWMMMMMMMMMMMMMMNK00OOKNMMMWX0OKN0o:,,,;cxOKNk,........,xl....c0o..........'o0o,........;o;..........' #
# 0kxl::lodkO0XNWWWKOOKXWMMMMMMMMMMMMMWWMMMMWNK0kddkd:,,,;cxKNWMNl....'loookk;....xO;....lx,....;;....:xoccllolc,.....;ccx #
# NX0O0KNNWMMMMMMMMWN0kkO0KXNNWWWWWWWWWNNXK0Okxdl:,,,,,;cxKNMMMM0,.....;;;lOd....;0d.....cc...,dk:.....';:oONWMNo....;KMWM #
# MMMMMMMMMMMMMMMMMMMWNX0OkkkkOO00000OOOkkkkkkxdc;,,,;cd0NMMMMMWd.....',''lk:....o0:..........:KWKxl:,......dWM0;....oWMMM #
# MMMMMMMMMMMMMMMMMMMMMMWWNXK000OOOOOO00KKXNNXOkxoc;cd0NWMMMMMMK:....lKXXXNx....,Ox....,kd'....collldOx,....dWWd....,0MMMM #
# MMMMMMMMMMMMMMMMMMMMMMMMMMMMMWWWWWWWWMMMMMMMWX0kxx0NWMMMMMMWWx....,OMMMMXc....l0c....oNk'...,c,....''...,dNMX:....lKNWMM #
# MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMNK0XWMMMMMMMMWNx;,,,oXMMMMKl,,,;kKl,,,:OM0:,,,l0Oc'....';dKWMM0c''';kKXWMM #
############################################################################################################################
 #                                                   FIRST STATS v1.0 2020                                                #
 #                                     A statistical scouting application for FIRST Robotics                              #
 #                                                 Written By: Mitch Zakocs                                               #
############################################################################################################################



import requests
import json
import base64
import configparser
import os
import collections
import time

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
    # TODO: Make config changable from sheets file

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
            with open('testing/qualscheduledata.json') as json_file:
                self.qualScheduleData = json.load(json_file)
            with open('testing/playoffscheduledata.json') as json_file:
                self.playoffScheduleData = json.load(json_file)
        else:    
            response = requests.get('%s/v2.0/2019/schedule/%s/qual/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.qualScheduleData = response.json()["Schedule"]
            response = requests.get('%s/v2.0/2019/schedule/%s/playoff/hybrid' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.playoffScheduleData = response.json()["Schedule"]
        if dprint == True:
            print(json.dumps(self.qualScheduleData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScheduleData, sort_keys=True, indent=4))

    def getScoreData (self, dprint = False):
        if (self.testing == True):
            with open('testing/qualscoredata.json') as json_file:
                self.qualScoreData = json.load(json_file)
            with open('testing/playoffscoredata.json') as json_file:
                self.playoffScoreData = json.load(json_file)
        else:
            response = requests.get('%s/v2.0/2019/scores/%s/qual' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.qualScoreData = response.json()["MatchScores"]
            response = requests.get('%s/v2.0/2019/scores/%s/playoff' % (self.config.host, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.playoffScoreData = response.json()["MatchScores"]
        if dprint == True:
            print(json.dumps(self.qualScoreData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScoreData, sort_keys=True, indent=4))

    def getEventData (self, dprint = False):
        if (self.testing == True):
            with open('testing/eventdata.json') as json_file:
                self.eventData = json.load(json_file)
        else:    
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
        ## Centered Values Setup
        csqp.updateCellRangeFormatting(2, 2, 3, 13, match, self.centered_format)
        csqp.updateCellRangeFormatting(5, 2, 6, 13, match, self.centered_format)
        
        ## Match Info
        csqp.updateCellValue(1, 3, match, "Team 1")
        csqp.updateCellValue(2, 3, match, match.schedule["teams"][0]["teamNumber"]) # Red
        csqp.updateCellValue(3, 3, match, match.schedule["teams"][3]["teamNumber"]) # Blue

        csqp.updateCellValue(1, 4, match, "Team 2")
        csqp.updateCellValue(2, 4, match, match.schedule["teams"][1]["teamNumber"]) # Red
        csqp.updateCellValue(3, 4, match, match.schedule["teams"][4]["teamNumber"]) # Blue

        csqp.updateCellValue(1, 5, match, "Team 3")
        csqp.updateCellValue(2, 5, match, match.schedule["teams"][2]["teamNumber"]) # Red
        csqp.updateCellValue(3, 5, match, match.schedule["teams"][5]["teamNumber"]) # Blue

        csqp.updateCellValue(1, 6, match, "Switch Level")
        csqp.updateCellValue(2, 6, match, (match.score["alliances"][1]["endgameRungIsLevel"] == "IsLevel")) # Red
        csqp.updateCellValue(3, 6, match, (match.score["alliances"][0]["endgameRungIsLevel"] == "IsLevel")) # Blue

        csqp.updateCellValue(1, 7, match, "Ranking Points")
        csqp.updateCellValue(2, 7, match, match.score["alliances"][1]["rp"]) # Red
        csqp.updateCellValue(3, 7, match, match.score["alliances"][0]["rp"]) # Blue

        ## Predictions
        # WIP
        csqp.updateCellValue(4, 3, match, "Team 1 MR")
        csqp.updateCellValue(4, 4, match, "Team 2 MR")
        csqp.updateCellValue(4, 5, match, "Team 3 MR")
        csqp.updateCellValue(4, 6, match, "Average MR")
        csqp.updateCellValue(4, 7, match, "Score Predict")

        ## Auto Score
        csqp.updateCellValue(1, 9, match, ("Inner " + u"\u25CF"))
        csqp.updateCellValue(2, 9, match, match.score["alliances"][1]["autoCellsInner"]) # Red
        csqp.updateCellValue(3, 9, match, match.score["alliances"][0]["autoCellsInner"]) # Blue

        csqp.updateCellValue(1, 10, match, ("Outer " + u"\u2B23"))
        csqp.updateCellValue(2, 10, match, match.score["alliances"][1]["autoCellsOuter"]) # Red
        csqp.updateCellValue(3, 10, match, match.score["alliances"][0]["autoCellsOuter"]) # Blue

        csqp.updateCellValue(1, 11, match, ("Bottom " + u"\u25A2"))
        csqp.updateCellValue(2, 11, match, match.score["alliances"][1]["autoCellsBottom"]) # Red
        csqp.updateCellValue(3, 11, match, match.score["alliances"][0]["autoCellsBottom"]) # Blue

        csqp.updateCellValue(1, 12, match, "Auto Total")
        csqp.updateCellValue(2, 12, match, match.score["alliances"][1]["autoPoints"]) # Red
        csqp.updateCellValue(3, 12, match, match.score["alliances"][0]["autoPoints"]) # Blue

        csqp.updateCellValue(1, 13, match, "Other Points")
        csqp.updateCellFormatting(1, 13, match, self.bold_format)
        csqp.updateCellValue(2, 13, match, (int(match.score["alliances"][1]["foulPoints"]) + int(match.score["alliances"][1]["adjustPoints"]) + int(match.score["alliances"][1]["controlPanelPoints"]) + int(match.score["alliances"][1]["endgamePoints"]))) # Red
        csqp.updateCellValue(3, 13, match, (int(match.score["alliances"][0]["foulPoints"]) + int(match.score["alliances"][0]["adjustPoints"]) + int(match.score["alliances"][0]["controlPanelPoints"]) + int(match.score["alliances"][0]["endgamePoints"]))) # Blue


        ## Teleop Score
        csqp.updateCellValue(4, 9, match, ("Inner " + u"\u25CF"))
        csqp.updateCellValue(5, 9, match, match.score["alliances"][1]["teleopCellsInner"]) # Red
        csqp.updateCellValue(6, 9, match, match.score["alliances"][0]["teleopCellsInner"]) # Blue

        csqp.updateCellValue(4, 10, match, ("Outer " + u"\u2B23"))
        csqp.updateCellValue(5, 10, match, match.score["alliances"][1]["teleopCellsOuter"]) # Red
        csqp.updateCellValue(6, 10, match, match.score["alliances"][0]["teleopCellsOuter"]) # Blue

        csqp.updateCellValue(4, 11, match, ("Bottom " + u"\u25A2"))
        csqp.updateCellValue(5, 11, match, match.score["alliances"][1]["teleopCellsBottom"]) # Red
        csqp.updateCellValue(6, 11, match, match.score["alliances"][0]["teleopCellsBottom"]) # Blue
        
        csqp.updateCellValue(4, 12, match, "Teleop Total")
        csqp.updateCellValue(5, 12, match, match.score["alliances"][1]["teleopPoints"]) # Red
        csqp.updateCellValue(6, 12, match, match.score["alliances"][0]["teleopPoints"]) # Blue

        csqp.updateCellValue(4, 13, match, "Score Total")
        csqp.updateCellRangeFormatting(4, 13, 6, 13, match, self.bold_format)
        csqp.updateCellValue(5, 13, match, match.score["alliances"][1]["totalPoints"]) # Red
        csqp.updateCellValue(6, 13, match, match.score["alliances"][0]["totalPoints"]) # Blue

        ### Team Setup ###
        ## Red and Blue Colors
        csqp.updateCellRangeFormatting(2, 2, 2, 13, match, self.red_color_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 13, match, self.blue_color_format)
        csqp.updateCellRangeFormatting(5, 2, 5, 13, match, self.red_color_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 13, match, self.blue_color_format)
        ## Team Color Labels
        csqp.updateCellValue(2, 2, match, "Red")
        csqp.updateCellValue(2, 8, match, "Red")
        csqp.updateCellValue(5, 2, match, "Red")
        csqp.updateCellValue(5, 8, match, "Red")
        csqp.updateCellValue(3, 2, match, "Blue")
        csqp.updateCellValue(3, 8, match, "Blue")
        csqp.updateCellValue(6, 2, match, "Blue")
        csqp.updateCellValue(6, 8, match, "Blue")

        ### Lines & Boxes ###
        ## Top Category Divider Lines
        csqp.updateCellRangeFormatting(1, 2, 6, 2, match, self.box_format)
        csqp.updateCellRangeFormatting(1, 8, 6, 8, match, self.box_format)
        csqp.updateCellRangeFormatting(1, 13, 6, 13, match, self.box_format)
        ## Main Category Line
        # Match Info
        csqp.updateCellRangeFormatting(1, 2, 1, 7, match, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(1, 2, 1, 7, match, self.box_right_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 7, match, self.box_left_format)
        csqp.updateCellRangeFormatting(3, 2, 3, 7, match, self.box_right_thick_format)
        # Predictions
        csqp.updateCellRangeFormatting(4, 2, 4, 7, match, self.box_left_format)
        csqp.updateCellRangeFormatting(4, 2, 4, 7, match, self.box_right_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 7, match, self.box_left_format)
        csqp.updateCellRangeFormatting(6, 2, 6, 7, match, self.box_right_thick_format)
        # Auto Score
        csqp.updateCellRangeFormatting(1, 8, 1, 13, match, self.box_left_format)
        csqp.updateCellRangeFormatting(1, 8, 1, 13, match, self.box_right_format)
        csqp.updateCellRangeFormatting(3, 8, 3, 13, match, self.box_left_format)
        csqp.updateCellRangeFormatting(3, 8, 3, 13, match, self.box_right_thick_format)
        # Teleop Score
        csqp.updateCellRangeFormatting(4, 8, 4, 13, match, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(4, 8, 4, 13, match, self.box_right_format)
        csqp.updateCellRangeFormatting(6, 8, 6, 13, match, self.box_left_format)
        csqp.updateCellRangeFormatting(6, 8, 6, 13, match, self.box_right_format)

        ### Scouting Notes Line Setup ###
        csqp.updateCellValue(1, 14, match, "Scouting Notes")
        csqp.updateCellFormatting(1, 14, match, self.category_title_format)
        csqp.updateCustomCellFormatting(1, 14, 6, 14, match, "merge")
        csqp.updateCustomCellFormatting(1, 15, 6, 17, match, "merge")
        csqp.updateCellRangeFormatting(1, 14, 6, 14, match, self.box_bottom_format)
        
        ### Thicc Outline Setup ###
        csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.box_top_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(6, 1, 6, 17, match, self.box_right_thick_format)
        csqp.updateCellRangeFormatting(1, 1, 1, 17, match, self.box_left_thick_format)
        csqp.updateCellRangeFormatting(1, 13, 6, 13, match, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(1, 17, 6, 17, match, self.box_bottom_thick_format)
        csqp.updateCellRangeFormatting(1, 8, 6, 8, match, self.box_top_thick_format)
        csqp.updateCellRangeFormatting(1, 13, 6, 13, match, self.box_top_thick_format)

        ### Entry Background Color ###
        ## Sets background color to green if the match has been played, grey otherwise
        if (match.schedule["postResultTime"] == "null"):
            csqp.updateCellRangeFormatting(1, 2, 1, 13, match, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(4, 2, 4, 13, match, self.match_scheduled_format)
            csqp.updateCellRangeFormatting(1, 14, 6, 17, match, self.match_scheduled_format)
        else:
            csqp.updateCellRangeFormatting(1, 2, 1, 13, match, self.match_played_format)
            csqp.updateCellRangeFormatting(1, 1, 6, 1, match, self.match_played_format)
            csqp.updateCellRangeFormatting(4, 2, 4, 13, match, self.match_played_format)
            csqp.updateCellRangeFormatting(1, 14, 6, 17, match, self.match_played_format)
        
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
                print ("Entry Found: %s" % tempMatch.matchtitle)
            except Exception as e:
                if (x == 0):
                    tempMatch.o_y = 3 # If it's the very first match, set the origin to the third row
                else:
                    # Search fails, get the location of the last entry, set the origin below it, and create the entry
                    locatorcell_list = self.ws.findall("Auto Score")
                    tempMatch.o_y = int(locatorcell_list.pop().row) + 10 # 10 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
                # Creates the entry's template
                self.createEntry(tempMatch)
                print ("Entry Created: %s" % e)
            self.qualMatchList.insert(len(self.qualMatchList), tempMatch)

        # Playoff Match Object Creation
        for x in range(0, len(self.data.playoffScheduleData)):
            tempMatch = Match(self.data.playoffScheduleData[x], self.data.playoffScoreData[x])
            try:
                # Try to find the cell
                cell = self.ws.find (tempMatch.matchtitle)
                tempMatch.o_y = cell.row - 1
                tempMatch.o_x = cell.col - 1
                print ("Entry Found: %s" % tempMatch.matchtitle)
            except Exception as e:
                # Search fails, get the location of the last entry, set the origin below it, and create the entry
                locatorcell_list = self.ws.findall("Auto Score")
                tempMatch.o_y = int(locatorcell_list.pop().row) + 10 # 10 is how many rows needs to be added to the bottom of Auto score to have a spot for a new entry
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
        self.cell_list = self.ws.ws.range('%s:%s' % (self.ws.convertLocal(1, 1, self.match), self.ws.convertLocal(6, 17, self.match))) # Grabs the whole section of cells
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
        self.matchtitle = ("%s Match #%s" % (self.schedule["description"].split()[0], self.schedule["description"].split()[1]))
        self.o_x = 0 # X Location of the Entry
        self.o_y = 0 # Y Location of the Entry

def main():
    # Config and Data Retrieval Setup
    config = Config()
    data = MatchData(config, testing=True)
    data.getScheduleData()
    data.getScoreData()
    data.getEventData(dprint=True)

    # # Sheets object creation
    sheets = Sheets(config, data)

    # # Create objects for each match and create a template in the sheets
    sheets.createMatchObjects()

if __name__ == "__main__":
    main()
