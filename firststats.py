import urllib2
import openpyxl
import base64
import json
import configparser
import os
import mysql.connector

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
        self.authString = base64.encodestring('%s:%s' % (self.username, self.password))
        # Connects to the SQL Database
        # self.db = mysql.connector.connect (
        #     host = self.config['MySQL']['Host'],
        #     user = self.config['MySQL']['User'],
        #     passwd = self.config['MySQL']['Password'],
        #     database = self.config['MySQL']['Database']
        # )
        # self.cursor = self.db.cursor()
    def GetMatchData ():
        test = {}
def main():
    data = Data()


if __name__ == "__main__":
    main()