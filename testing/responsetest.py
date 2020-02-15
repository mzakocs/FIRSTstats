import requests
import json
import configparser
import os
import base64

config = configparser.ConfigParser()
if not os.path.exists('config.ini'):
    config ['FIRST API'] = {'Host':'', 'Username':'', 'Token':'frc-api.firstinspires.org'}
    config ['Google Sheets'] = {'sheetid': ''}
    with open ('config.ini', 'w') as configfile:
        config.write(configfile)
config.read('config.ini')
sheetid = config['Google Sheets']['sheetid']
host = config['FIRST API']['Host']
username = config['FIRST API']['Username']
password = config['FIRST API']['Token']
authString = base64.b64encode(('%s:%s' % (username, password)).encode('utf-8')).decode()
response = requests.get('https://frc-api.firstinspires.org/v2.0/2019/schedule/CALN/qual/hybrid', headers={'Accept': 'application/json', 'Authorization': 'Basic %s' % authString})
print (response)