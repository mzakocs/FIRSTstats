import urllib2
import base64
import json
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
    def GetMatchData (self):
        request = urllib2.Request('%s/v2.0/2017/matches/%s' % (self.host, self.matchnum))
        request.add_header('Accept', 'application/json')
        request.add_header('Authorization', 'Basic %s' % self.authString)
        matchJSON = json.load(urllib2.urlopen(request))
        self.matchData = matchJSON
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
    print(data.matchData)

if __name__ == "__main__":
    main()