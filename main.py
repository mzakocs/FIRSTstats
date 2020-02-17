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
 #                                               Written for Python 3.0 and up                                            #
 #                                                 Written By: Mitch Zakocs                                               #
############################################################################################################################

# TODO: Eventually update rating system to GLICKO 2
# TODO: Create checks to make sure config input is valid and doesn't crash the service, especially match codes and seasons
# TODO: Fix display of points

import requests
import json
import firstconfig
import firstsheets
import time

class MatchData:
    def __init__(self, config):
        # Class Imports
        self.config = config

    ### Data Retrieval Functions ###
    ## Gets Data from the FIRST API
    # @param dprint - formats the json dictionary into a readable format and prints it
    def getScheduleData(self, dprint = False):
        if (self.config.dataretrieval == "Manual"):
            with open('manual/manualqualscheduledata.json') as json_file:
                self.qualScheduleData = json.load(json_file)
            with open('manual/manualplayoffscheduledata.json') as json_file:
                self.playoffScheduleData = json.load(json_file)
        elif (self.config.dataretrieval == "Testing"):
            with open('testing/qualscheduledata.json') as json_file:
                self.qualScheduleData = json.load(json_file)
            with open('testing/playoffscheduledata.json') as json_file:
                self.playoffScheduleData = json.load(json_file)
        else:
            try:    
                response = requests.get('%s/v2.0/%s/schedule/%s/qual/hybrid' % (self.config.host, str(self.config.season), self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
                self.qualScheduleData = response.json()["Schedule"]
                response = requests.get('%s/v2.0/%s/schedule/%s/playoff/hybrid' % (self.config.host, str(self.config.season), self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
                self.playoffScheduleData = response.json()["Schedule"]
            except:
                print("Cannot retrieve Match Data from FIRST API!")
        if dprint == True:
            print(json.dumps(self.qualScheduleData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScheduleData, sort_keys=True, indent=4))

    def getScoreData(self, dprint = False):
        if (self.config.dataretrieval == "Manual"):
            with open('manual/manualqualscoredata.json') as json_file:
                self.qualScoreData = json.load(json_file)
            with open('manual/manualplayoffscoredata.json') as json_file:
                self.playoffScoreData = json.load(json_file)
        elif (self.config.dataretrieval == "Testing"):
            with open('testing/qualscoredata.json') as json_file:
                self.qualScoreData = json.load(json_file)
            with open('testing/playoffscoredata.json') as json_file:
                self.playoffScoreData = json.load(json_file)
        else:
            try:
                response = requests.get('%s/v2.0/%s/scores/%s/qual' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
                self.qualScoreData = response.json()["MatchScores"]
                response = requests.get('%s/v2.0/%s/scores/%s/playoff' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
                self.playoffScoreData = response.json()["MatchScores"]
            except:
                print("Cannot retrieve Score Data from FIRST API!")
        if dprint == True:
            print(json.dumps(self.qualScoreData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScoreData, sort_keys=True, indent=4))

    def getEventData(self, dprint = False):
        if (self.config.dataretrieval == "Manual"):
            with open('manual/manualeventdata.json') as json_file:
                self.eventData = json.load(json_file)
        elif (self.config.dataretrieval == "Testing"):
            with open('testing/eventdata.json') as json_file:
                self.eventData = json.load(json_file)
        else:
            try:
                response = requests.get('%s/v2.0/%s/events?eventCode=%s' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
                self.eventData = response.json()["Events"][0]
            except:
                print("Cannot retrieve Event Data from FIRST API!")
        if dprint == True:
            print(json.dumps(self.eventData, sort_keys=True, indent=4))

    def getTeamData(self, dprint = False):
        if (self.config.dataretrieval == "Manual"):
            with open('manual/manualteamdata.json') as json_file:
                self.teamData = json.load(json_file)
        elif (self.config.dataretrieval == "Testing"):
            with open('testing/teamdata.json') as json_file:
                self.teamData = json.load(json_file)
        else:
            try: 
                response = requests.get('%s/v2.0/%s/teams?eventCode=%s' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
                if (response.json()["pageTotal"] > 1):
                    pageList = []
                    self.teamData = {"teams" : []}
                    for x in range(0, response.json()["pageTotal"]):
                        tempresponse = requests.get('%s/v2.0/%s/teams?eventCode=%s?page=%s' % (self.config.host, self.config.season, self.config.eventid, x), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
                        pageList.insert(len(pageList), tempresponse.json()['teams'])
                    for x in pageList:
                        self.teamData["teams"] += x
                else:
                    self.teamData = response.json()["teams"]
            except:
                print("Cannot retrieve Team Data from FIRST API!")
        if dprint == True:
            print(json.dumps(self.teamData, sort_keys=True, indent=4))
            
    def dataChanged(self):
        # Stores the old schedules to check
        oldQualSchedule = self.qualScheduleData
        oldQualScore = self.qualScoreData
        oldPlayoffSchedule = self.playoffScheduleData
        oldPlayoffScore = self.playoffScoreData
        # Gets the new data
        self.getScoreData()
        self.getScheduleData()
        # Checks to see if any of them have changed
        if oldQualSchedule != self.qualScheduleData or oldQualScore != self.qualScoreData or oldPlayoffSchedule != self.playoffScheduleData or oldPlayoffScore != self.playoffScoreData:
            print("Detected Change in Data from FIRST API!")
            return True
        else:
            return False
    
    def checkIfMatchValid(self):
        # Checks to see if the match actually exists
        # Used to protect against program crashing if an invalid
        # match ID or season is entered
        try:
            # Attempts to get the event data
            response = requests.get('%s/v2.0/%s/events?eventCode=%s' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            tempEventData = response.json()["Events"][0]
            return True
        except:
            # If it fails then return false
            print("Invalid Event Entered: %s" % self.config.eventid)
            return False


def main():
    # Config and Data Retrieval Setup
    config = firstconfig.FirstConfig()
    data = MatchData(config)
    data.getScheduleData()
    data.getScoreData()
    data.getEventData()
    data.getTeamData()

    # Sheets object creation
    sheets = firstsheets.Sheets(config, data)

    # Create objects for each match and create a template in the sheets
    sheets.createTeamObjects()
    sheets.createMatchObjects()
    sheets.createMatchEntries()
    sheets.createTeamEntry()

    # Main Execution Loop
    starttime = time.ctime()
    print("Service started at:", starttime)
    csqp = firstsheets.UCSQP(sheets.config_ws, sheets.sh, 4, 6, 5, 14)
    while True:
        # Grabs the new config from the sheet
        csqp.updateList()
        # Checks to see if the MatchID is valid
        validMatch = data.checkIfMatchValid()
        # Checks to see if the config has changed
        configChanged = config.checkSheetsConfig(sheets.config_ws, csqp)
        if config.powerswitch == "On" and validMatch == True:
            # Setup for manual data entry, currently not working
            if config.dataretrieval == "Manual":
                sheets.checkManualEntry()
            # Checks to see if the data has changed at all from the stored value
            dataChanged = data.dataChanged()
            # This is in case somebody deletes the sheet without changing config
            if sheets.checkIfSheetExists() == False and configChanged != "Match":
                sheets.createSheet()
                sheets.createMatchEntries(noNotes = True) # Makes sure to not grab the notes so that it uses the stored ones
                sheets.createTeamEntry(noNotes = True)
            # Update data, push data to sheet and create entries and new objects if match schedule or score data is different from server
            if dataChanged == True and sheets.checkIfSheetExists() == True:
                data.getScheduleData()
                data.getScoreData()
                data.getEventData()
                data.getTeamData()
                sheets.data = data
                sheets.createMatchObjects()
                sheets.createMatchEntries()
                sheets.createTeamEntry()
        # Push new config and update info if config has changed
        if configChanged != False:
            print("Config Change Detected!")
            # Pushes the new config to the sheets and data objects
            sheets.config = config
            data.config = config
            validMatch = data.checkIfMatchValid()
            if validMatch == True:
                if configChanged == "Match":
                    # If the sheet doesn't exist, that means the match was changed
                    # This grabs all of the new match data and sets up a new sheet for it
                    data.getScheduleData()
                    data.getScoreData()
                    data.getEventData()
                    data.getTeamData()
                    sheets.data = data
                    if sheets.checkIfSheetExists() == False:
                        sheets.createSheet()
                    sheets.createTeamObjects()
                    sheets.createMatchObjects(wipeList = True)
                if configChanged == "MatchFilter":
                    # Deletes all entries to make way for the new filtered entries
                    # This is because there may be less entries now and we can't
                    # Simply write over them, there will be extra on the bottom
                    sheets.nukeMatchEntries()
                # Creates the new Match Entries for the new filters, match or fresh data
                if configChanged != "Power":
                    sheets.createMatchEntries()
                    sheets.createTeamEntry()
                else:
                    print("Turning Service %s!" % config.powerswitch)
        # Sleeps for 8 seconds before checking again
        time.sleep(8)

if __name__ == "__main__":
    main()
