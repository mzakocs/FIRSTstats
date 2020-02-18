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
# TODO: Code unit tests eventually
# TODO: Add a table on the home screen or another page that has a global list of teams with global MitchRatings for the 2020 season,
# would be either only matches that we have analyzed orrrr maybe all matches for 2020?
# TODO: Add predictions for whether or not the robot has a functional auto phase and whether
# or not it has a functional hanging thing

import firstconfig
import firstsheets
import firstdata
import datetime
import time

def main():
    # Config and Data Retrieval Setup
    config = firstconfig.FirstConfig()
    data = firstdata.MatchData(config)
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
    print("Service started at:", datetime.datetime.fromtimestamp(starttime).strftime('%Y-%m-%d %H:%M:%S'))
    csqp = firstsheets.UCSQP(sheets.config_ws, sheets.sh, 4, 6, 5, 14)
    while True:
        try:
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
                        if configChanged != "TeamFilter":
                            sheets.createMatchEntries()
                        sheets.createTeamEntry()
                    else:
                        print("Turning Service %s!" % config.powerswitch)
        except:
            # If any of the functions fail, try to relogin to GSpread so it doesn't drop connection
            print("Attempting to reconnect to Google Sheets API...")
            sheets.gc.login()                
            print("Done!")
        # Sleep for 8 seconds before checking again
        time.sleep(8)

if __name__ == "__main__":
    main()
