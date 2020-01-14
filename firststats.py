import requests
import json
import base64
import configparser
import os
import mysql.connector
import collections

# pip gspread, pyopenssl and oauth2client for google sheets functionality
import gspread
from oauth2client.service_account import ServiceAccountCredentials

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
        self.matchnum = self.config['Match Config']['matchid']
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

        # # Sets up a requests session
        # req = requests.Session()
        # req.auth = (self.username, self.password)
        # req.headers.update({'Accept': 'application/json'})
    def GetMatchData (self):
        response = requests.get('%s/v2.0/2017/matches/%s' % (self.host, self.matchnum), headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % self.authString})
        self.matchData = response.json()

class Sheets:
    def __init__(self):
        # Google Drive OAuth2 Setup
        # gspread.readthedocs.io/en/latest/
        scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('FIRST Python Stats-c64a29c90ec3.json', scope)
        self.gc = gspread.authorize(credentials)
def main():
    data = Data()
    data.GetMatchData()
    print(json.dumps(data.matchData, sort_keys=True, indent=4, default=str))

if __name__ == "__main__":
    main()