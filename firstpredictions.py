############################################################################################################################
# FirstPredictions                                                                                                         #
# Used to rank teams and predict matches                                                                                   #
# Also used to track matches and teams                                                                                     #                                                                                     #
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

        # Figures out which team won or if it was a tie
        self.teamWinner = ""
        if (self.schedule["scoreRedFinal"] > self.schedule["scoreBlueFinal"]):
            self.teamWinner = "R"
        elif (self.schedule["scoreRedFinal"] < self.schedule["scoreBlueFinal"]):
            self.teamWinner = "B"
        else:
            self.teamWinner = "T" # If neither of those conditions are met, the match is a tie or something else went wrong

        ## Info about the Match Entry
        self.matchnum = self.schedule["matchNumber"]
        self.matchtype = self.schedule["tournamentLevel"]
        self.matchhappened = not (self.schedule["postResultTime"] == "null")
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
    _c = 1
    _q = 0.0057565

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
        self.mitchrating = 1500
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
    
    def getRankTitle(self):
        pass
