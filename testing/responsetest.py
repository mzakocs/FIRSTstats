import requests
import json
import configparser
import os
import base64

config = configparser.ConfigParser()
if not os.path.exists('config.ini'):
    config ['Match Config'] = {'matchid':'CMPMO'}
    config ['FIRST API'] = {'Host':'', 'Username':'', 'Token':'frc-api.firstinspires.org'}
    config ['Google Sheets'] = {'sheetid': ''}
    with open ('config.ini', 'w') as configfile:
        config.write(configfile)
config.read('config.ini')
matchid = config['Match Config']['matchid']
sheetid = config['Google Sheets']['sheetid']
host = config['FIRST API']['Host']
username = config['FIRST API']['Username']
password = config['FIRST API']['Token']
authString = base64.b64encode('%s:%s' % (username, password))
response = requests.get('https://frc-staging-api.firstinspires.org/v2.0/2020/events', headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % authString})
print (response)