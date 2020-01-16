import requests
import json
import base64
import configparser
import os
import mysql.connector
import collections

# Angels Asks
# Have a scale of how bad they are Absolutely Useless - Amazing

# pip gspread, gspread_formatting, pyopenssl and oauth2client for google sheets functionality
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import gspread_formatting

class Data:
    def __init__(self):
        # Config File Creation and Retrieval
        self.config = configparser.ConfigParser()
        if not os.path.exists('config.ini'):
            self.config ['Match Config'] = {'matchid':'CMPMO'}
            self.config ['FIRST API'] = {'Host':'', 'Username':'', 'Token':'frc-api.firstinspires.org'}
            self.config ['MySQL'] = {'Host':'localhost', 'User':'', 'Password':'', 'Database':''}
            with open ('config.ini', 'w') as configfile:
                self.config.write(configfile)
        self.config.read('config.ini')
        self.matchid = self.config['Match Config']['matchid']
        self.host = self.config['FIRST API']['Host']
        self.username = self.config['FIRST API']['Username']
        self.password = self.config['FIRST API']['Token']
        self.authString = base64.b64encode('%s:%s' % (self.username, self.password))
        # Connects to the SQL Database
        # self.db = mysql.connector.connect (
        #     host = self.config['MySQL']['Host'],
        #     user = self.config['MySQL']['User'],
        #     passwd = self.config['MySQL']['Password'],
        #     database = self.config['MySQL']['Database']
        # )
        # self.cursor = self.db.cursor()
    def GetMatchData (self):
        response = requests.get('%s/v2.0/2019/matches/%s' % (self.host, self.matchid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.authString})
        self.matchData = response.json()
    def GetScoreData (self):
        response = requests.get('%s/v2.0/2019/scores/%s/playoff' % (self.host, self.matchid), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.authString})
        self.scoreData = response.json()

class Sheets:
    def __init__(self):
        # Google Drive OAuth2 Setup
        # gspread.readthedocs.io/en/latest/
        scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('FIRST Python Stats-c64a29c90ec3.json', scope)
        self.gc = gspread.authorize(credentials)
        self.sh = self.gc.open_by_key('1F5In14PEAsNHCy-moPQ_rd85HUcSzjElGpW0IsOTBKY')
    def WorksheetSetup(self, matchid):
        # Creates the worksheet if it doesn't exist
        try:
            self.ws = self.sh.worksheet(str(matchid))
        except Exception as e:
            self.ws = self.sh.add_worksheet(title=str(matchid), rows="200", cols="30")
            print("Worksheet Created: " + str(e))
def main():
    data = Data()
    sheets = Sheets()

    data.GetMatchData()
    data.GetScoreData()
    print(json.dumps(data.matchData, sort_keys=True, indent=4, default=str))
    print(json.dumps(data.scoreData, sort_keys=True, indent=4, default=str))

    sheets.WorksheetSetup(data.matchid)

if __name__ == "__main__":
    main()