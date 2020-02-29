############################################################################################################################
# FirstPredictions                                                                                                         #
# Used to rank teams and predict matches                                                                                   #
# Also used to track matches and teams                                                                                     #
# If the ranking system isn't working, make sure you're using Python 3.8.1                                                 #                                                                                     #
############################################################################################################################

import math

class Match:
    # A class that represents a match entry in the spreadsheet
    # One is created for each match in an event
    def __init__(self, scheduledict, scoredict):
        # Class Imports
        self.schedule = scheduledict
        self.score = scoredict

        # Grabs the teams depending on color
        self.redTeamList = []
        self.blueTeamList = []
        for team in self.schedule["teams"]:
            # Adds the team to the list if it's Red and Winning
            if team["station"][0] == 'R':
                self.redTeamList.insert(len(self.redTeamList), str(team["teamNumber"]))
            # Adds the team to the list if it's Blue and Winning
            elif team["station"][0] == 'B':
                self.blueTeamList.insert(len(self.blueTeamList), str(team["teamNumber"]))

        ## Info about the Match Entry
        self.matchnum = self.schedule["matchNumber"]
        self.matchtype = self.schedule["tournamentLevel"]
        self.matchhappened = not (self.schedule["actualStartTime"] == "null")
        self.matchtitle = ("%s Match #%s %s" % (self.schedule["description"].split()[0], self.schedule["description"].split()[1], self.formatDate()))
        self.o_x = 1 # X Location of the Entry
        self.o_y = 1 # Y Location of the Entry
        self.notes = ""

        # Prediction info for the Match Entry
        self.redteam_mr = {}
        self.blueteam_mr = {}
        self.redteam_avmr = 0
        self.redteam_avrd = 0
        self.blueteam_avmr = 0
        self.blueteam_avrd = 0
        self.redteam_avscr = 0
        self.blueteam_avscr = 0

    def formatDate(self):
        datevalues = self.schedule["startTime"].split("T")[0].split("-")
        return ("(%s/%s/%s)" % (datevalues[1], datevalues[2], datevalues[0]))

    def calculateAverages(self, teamDict, scoreav = False):
        self.redteam_avmr = 0
        self.redteam_avrd = 0
        # Calculates the average MR and RD of the Red Team
        for teamNumber in self.redTeamList:
            self.redteam_avmr += teamDict[teamNumber].mitchrating
            self.redteam_avrd += teamDict[teamNumber].ratingdeviation
        self.redteam_avmr //= len(self.redTeamList)
        self.redteam_avrd //= len(self.redTeamList)

        self.blueteam_avmr = 0
        self.blueteam_avrd = 0
        # Calculates the average MR and RD of the Blue Team
        for teamNumber in self.blueTeamList:
            self.blueteam_avmr += teamDict[teamNumber].mitchrating
            self.blueteam_avrd += teamDict[teamNumber].ratingdeviation
        self.blueteam_avmr //= len(self.blueTeamList)
        self.blueteam_avrd //= len(self.blueTeamList)

        if scoreav == True:
            self.redteam_avscr = 0
            self.blueteam_avscr = 0
            # Updating score averages
            for team in self.redTeamList:
                teamDict[team].totalScore += int(self.schedule["scoreRedFinal"])
                teamDict[team].matchesPlayed += 1
                self.redteam_avscr += (teamDict[team].totalScore // teamDict[team].matchesPlayed)
            for team in self.blueTeamList:
                teamDict[team].totalScore += int(self.schedule["scoreBlueFinal"])
                teamDict[team].matchesPlayed += 1
                self.blueteam_avscr += (teamDict[team].totalScore // teamDict[team].matchesPlayed)
            self.redteam_avscr //= len(self.redTeamList)
            self.blueteam_avscr //= len(self.blueTeamList)

    def updateTeamScores(self, teamDict):
        # Figures out which team won or if it was a tie
        self.teamWinner = ""
        if (int(self.schedule["scoreRedFinal"]) > int(self.schedule["scoreBlueFinal"])):
            self.teamWinner = "R"
        elif (int(self.schedule["scoreRedFinal"]) < int(self.schedule["scoreBlueFinal"])):
            self.teamWinner = "B"
        else:
            self.teamWinner = "T" # If neither of those conditions are met, the match is a tie or something else went wrong
        
        # Calculates the averages of all the teams to use in the calculations
        self.calculateAverages(teamDict)

        # Updates each individual team and the match entry
        if self.teamWinner == "R":
            for teamNumber in self.redTeamList:
                teamDict[teamNumber].wonAgainst(self.blueteam_avmr, self.blueteam_avrd)
                self.redteam_mr[teamNumber] = teamDict[teamNumber].mitchrating
            for teamNumber in self.blueTeamList:
                teamDict[teamNumber].lostAgainst(self.redteam_avmr, self.redteam_avrd)
                self.blueteam_mr[teamNumber] = teamDict[teamNumber].mitchrating
        elif self.teamWinner == "B":
            for teamNumber in self.redTeamList:
                teamDict[teamNumber].lostAgainst(self.blueteam_avmr, self.blueteam_avrd)
                self.redteam_mr[teamNumber] = teamDict[teamNumber].mitchrating
            for teamNumber in self.blueTeamList:
                teamDict[teamNumber].wonAgainst(self.redteam_avmr, self.redteam_avrd)
                self.blueteam_mr[teamNumber] = teamDict[teamNumber].mitchrating
        else:
            for teamNumber in self.redTeamList:
                teamDict[teamNumber].tiedAgainst(self.blueteam_avmr, self.blueteam_avrd)
                self.redteam_mr[teamNumber] = teamDict[teamNumber].mitchrating
            for teamNumber in self.blueTeamList:
                teamDict[teamNumber].tiedAgainst(self.redteam_avmr, self.redteam_avrd)
                self.blueteam_mr[teamNumber] = teamDict[teamNumber].mitchrating
        
        # Recalculates the averages to use on the Google Sheet
        self.calculateAverages(teamDict, scoreav = True)    


class Team:
    # A class that represents an individual team in a match
    # One is created for each team in an event
    ## Uses an implementation of the GLICKO Rating system to rank teams
    # http://www.glicko.net/glicko.html
    # Rating deviation is almost like standard deviation, it shows how consistent the team is or if they are just getting carried
    _c = 7.95
    _q = 0.00975646273 # ln(10)/400

    def __init__(self, teamdict):
        # Class Imports
        self.team = teamdict
        # Team Entry Origin
        self.o_x = 1
        self.o_y = 1
        # Team Information
        self.name = self.team["nameShort"]
        self.number = self.team["teamNumber"]
        self.notes = ""
        self.robotType = ""
        self.robotWeight = ""
        # Predictions Data
        self.mitchrating = 2000
        self.ratingdeviation = 350
        self.totalScore = 0
        self.matchesPlayed = 0

    @property
    def tranformed_rd(self):
        return min([350, math.sqrt(self.ratingdeviation ** 2 + self._c ** 2)])

    @classmethod   
    def _g(cls, x):
        return 1 / (math.sqrt(1 + 3 * cls._q ** 2 * (x ** 2) / math.pi ** 2))

    def expected_score(self, ooamr, ooard, inverse = False):
        #@param oaamr - Opponent Alliance Average Mitch Rating
        #@param oaard - Oponent Alliance Average Rating Deviation
        if inverse == True:
            g_term = self._g(math.sqrt(ooard ** 2 + self.ratingdeviation ** 2) * (ooamr - self.mitchrating) / 400)
        else:
            g_term = self._g(math.sqrt(self.ratingdeviation ** 2 + ooard ** 2) * (self.mitchrating - ooamr) / 400)
        return (1 / (1 + 10 ** (-1 * g_term)))

    def wonAgainst(self, oaamr, oaard):
        #@param oaamr - Opponent Alliance Average Mitch Rating
        #@param oaard - Oponent Alliance Average Rating Deviation
        s = 1
        E_term = self.expected_score(oaamr, oaard)
        d_squared = (self._q ** 2 * (self._g(oaard) ** 2 * E_term * (1 - E_term))) ** -1
        s_new_mitchrating = self.mitchrating + (self._q / (1 / self.ratingdeviation ** 2 + 1 / d_squared)) * self._g(oaard) * (s - E_term)
        s_new_ratingdeviation = math.sqrt((1 / self.ratingdeviation ** 2 + 1 / d_squared) ** -1)
        self.mitchrating = round(s_new_mitchrating)
        self.ratingdeviation = round(s_new_ratingdeviation)

    def lostAgainst(self, oaamr, oaard):
        #@param oaamr - Opponent Alliance Average Mitch Rating
        #@param oaard - Opponent Alliance Average Rating Deviation
        s = 0
        E_term = self.expected_score(oaamr, oaard, inverse = True)
        d_squared = (self._q ** 2 * (self._g(self.ratingdeviation) ** 2 * E_term * (1 - E_term))) ** -1
        s_new_mitchrating = oaamr + (self._q / (1 / oaard ** 2 + 1 / d_squared)) * self._g(self.ratingdeviation) * (s - E_term)
        s_new_ratingdeviation = math.sqrt((1 / oaard ** 2 + 1 / d_squared) ** -1)
        self.mitchrating = round(s_new_mitchrating)
        self.ratingdeviation = round(s_new_ratingdeviation)
    
    def tiedAgainst(self, oaamr, oaard):
        #@param oaamr - Opponent Alliance Average Mitch Rating
        #@param oaard - Opponent Alliance Average Rating Deviation
        s = 0.5
        E_term = self.expected_score(oaamr, oaard)
        d_squared = (self._q ** 2 * (self._g(oaard) ** 2 * E_term * (1 - E_term))) ** -1
        s_new_mitchrating = self.mitchrating + (self._q / (1 / self.ratingdeviation ** 2 + 1 / d_squared)) * self._g(oaard) * (s - E_term)
        s_new_ratingdeviation = math.sqrt((1 / self.ratingdeviation ** 2 + 1 / d_squared) ** -1)
        self.mitchrating = round(s_new_mitchrating)
        self.ratingdeviation = round(s_new_ratingdeviation)
    
    def getRankTitle(self, max, min):
        # Returns a function to display the corresponding CSGO rank for the MitchRating
        # Implemented on the request of Angel Heredia
        # Can be removed, has literally no purpose other than looking nice
        i = range(min, max, round((max-min)/17))
        # Silver 1
        if self.mitchrating <= i[0]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/1.png", 4, 20, 50)'
        # Silver 2
        elif i[0] < self.mitchrating <= i[1]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/2.png", 4, 20, 50)'
        # Silver 3
        elif i[1] < self.mitchrating <= i[2]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/3.png", 4, 20, 50)'
        # Silver 4
        elif i[2] < self.mitchrating <= i[3]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/4.png", 4, 20, 50)'
        # Silver Elite
        elif i[3] < self.mitchrating <= i[4]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/5.png", 4, 20, 50)'
        # Silver Elite Master
        elif i[4] < self.mitchrating <= i[5]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/6.png", 4, 20, 50)'
        # Gold Nova 1
        elif i[5] < self.mitchrating <= i[6]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/7.png", 4, 20, 50)'
        # Gold Nova 2
        elif i[6] < self.mitchrating <= i[7]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/8.png", 4, 20, 50)'
        # Gold Nova 3
        elif i[7] < self.mitchrating <= i[8]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/9.png", 4, 20, 50)'
        # Gold Nova Master
        elif i[8] < self.mitchrating <= i[9]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/10.png", 4, 20, 50)'
        # Master Guard
        elif i[9] < self.mitchrating <= i[10]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/11.png", 4, 20, 50)'
        # Master Guard 2
        elif i[10] < self.mitchrating <= i[11]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/12.png", 4, 20, 50)'
        # Master Guard Elite
        elif i[11] < self.mitchrating <= i[12]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/13.png", 4, 20, 50)'
        # Distinguished Master Guard
        elif i[12] < self.mitchrating <= i[13]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/14.png", 4, 20, 50)'
        # Legendary Eagle
        elif i[13] < self.mitchrating <= i[14]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/15.png", 4, 20, 50)'
        # Legendary Eagle Master
        elif i[14] < self.mitchrating <= i[15]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/16.png", 4, 20, 50)'
        # Supreme Master First Class
        elif i[15] < self.mitchrating <= i[16]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/17.png", 4, 20, 50)'
        # Global Elite
        elif self.mitchrating > i[16]:
            return '=IMAGE("https://csgo-stats.com/custom/img/ranks/18.png", 4, 20, 50)'
        # Theoretically should never happen but who tf knows
        else:
            return '=IMAGE("https://ih1.redbubble.net/image.792313560.3852/flat,550x550,075,f.u2.jpg")'
            print("Rank Title Not Found!")

