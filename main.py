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
import firstconfig
import firstsheets

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
                self.eventData = json.load(json_file)
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
    data = MatchData(config)
    data.getScheduleData()
    data.getScoreData()
    data.getEventData()
    data.getTeamData()

    # Sheets object creation
    sheets = firstsheets.Sheets(config, data)

    # Create objects for each match and create a template in the sheets
    sheets.createMatchObjects()
    sheets.createTeamObjects()

if __name__ == "__main__":
    main()
