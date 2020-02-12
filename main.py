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

# TODO: Test out the match making filters
    # SUB TODO: Figure out a way to delete or edit all entries to show only the filtered ones 
# TODO: Test out the sheets config
# TODO: Make a match checker and updater
# TODO: Finish prediction for unplayed matches
# TODO: Update config class to include the new playoff match customization

import requests
import json
import firstconfig
import firstsheets
import time

class MatchData:
    def __init__(self, config, testing = False):
        # Class Imports
        self.config = config
        self.testing = testing

    ### Data Retrieval Functions ###
    ## Gets Data from the FIRST API
    # @param testing - allows usage of a local JSON file for testing instead of the FIRST API
    # @param dprint - formats the json dictionary into a readable format and prints it
    def getScheduleData (self, dprint = False):
        if (self.testing == True):
            with open('testing/qualscheduledata.json') as json_file:
                self.qualScheduleData = json.load(json_file)
            with open('testing/playoffscheduledata.json') as json_file:
                self.playoffScheduleData = json.load(json_file)
        else:    
            response = requests.get('%s/v2.0/%s/schedule/%s/qual/hybrid' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.qualScheduleData = response.json()["Schedule"]
            response = requests.get('%s/v2.0/%s/schedule/%s/playoff/hybrid' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
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
            response = requests.get('%s/v2.0/%s/scores/%s/qual' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.qualScoreData = response.json()["MatchScores"]
            response = requests.get('%s/v2.0/%s/scores/%s/playoff' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.playoffScoreData = response.json()["MatchScores"]
        if dprint == True:
            print(json.dumps(self.qualScoreData, sort_keys=True, indent=4))
            print(json.dumps(self.playoffScoreData, sort_keys=True, indent=4))

    def getEventData (self, dprint = False):
        if (self.testing == True):
            with open('testing/eventdata.json') as json_file:
                self.eventData = json.load(json_file)
        else:    
            response = requests.get('%s/v2.0/%s/events?eventCode=%s' % (self.config.host, self.config.season, self.config.eventid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.config.authString})
            self.eventData = response.json()["Events"][0]
        if dprint == True:
            print(json.dumps(self.eventData, sort_keys=True, indent=4))

    def getTeamData (self, dprint = False):
        if (self.testing == True):
            with open('testing/teamdata.json') as json_file:
                self.teamData = json.load(json_file)
        else:    
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
        if dprint == True:
            print(json.dumps(self.teamData, sort_keys=True, indent=4))

def main():
    # Config and Data Retrieval Setup
    config = firstconfig.FirstConfig()
    data = MatchData(config, testing = True)
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
    starttime = time.time()
    print("Service started at: ", starttime)
    while True:
        csqp = firstsheets.UCSQP(sheets.config_ws, sheets.sh, 4, 6, 5, 13)
        configChanged = config.checkSheetsConfig(sheets.config_ws, csqp)
        print ("configChanged: " + str(configChanged))
        # Add another condition to the if statement that checks
        # if the data from the API has been changed and if it has
        # update or add the entries
        # Either check the sheet data or store the last used JSON and check it
        # against that
        if configChanged != False:
            # Pushes the new config to the sheets and data objects
            sheets.config = config
            data.config = config
            if sheets.checkIfSheetExists() == False:
                # If the sheet doesn't exist, that means the match was changed
                # This grabs all of the new match data and sets up a new sheet for it
                data.getScheduleData()
                data.getScoreData()
                data.getEventData()
                data.getTeamData()
                sheets.data = data
                sheets.createTeamObjects()
                sheets.createMatchObjects()
            if configChanged == "Filter":
                # Deletes all entries to make way for the new filtered entries
                # This is because there may be less entries now and we can't
                # Simply write over them, there will be extra on the bottom
                sheets.nukeMatchEntries()
            # if data.dataMatches == False:
                # if it's a new match, creates a new match object and pushes it to the list
                # if it's simply a match that's been played, update the data
                # then creates a new match entry and puts it at the bottom 
            sheets.createTeamEntry()
            sheets.createMatchEntries()
        # Sleeps for 60 seconds on the clock before checking again
        time.sleep(10)

    

if __name__ == "__main__":
    main()
