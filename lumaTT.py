import math
import random
import getpass
import datetime
import collections
import itertools

import enum

import challonge

import gdata.service
import gdata.spreadsheet.service

class ScheduleMakerError(Exception): pass
class WorksheetNotFoundError(ScheduleMakerError): pass
class TimeSlotNotFoundError(ScheduleMakerError): pass
class DistributionError(ScheduleMakerError): pass
class SpreadsheetUpdateError(ScheduleMakerError): pass

#==============================================================================
# Utility functions
#==============================================================================
def _parseDate(dateString):
    return datetime.datetime.strptime(dateString, '%m/%d/%Y').date()

def _weekdayNum(dayString):
    return ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday',
            'sunday'].index(dayString.lower())

def boolstr(inputStr):
    inputStr = inputStr.lower()
    if inputStr in ('false', 'no', 'nil', 'null', 'none', '-'):
        return False
    elif inputStr in ('true', 'yes', '+'):
        return True
    return bool(int(inputStr))

Availability = enum.Enum('hitMax', 'hitMin', 'underMin')
HIT_MAX = Availability.hitMax
HIT_MIN = Availability.hitMin
UNDER_MIN = Availability.underMin

Services = enum.Enum('google', 'challonge')
GOOGLE = Services.google
CHALLONGE = Services.challonge

class DayData(object):
    def __init__(self, date, group=None, matches=None, minGames=None,
                 maxGames=None, allottedGames=0):
        if matches is None:
            matches = []
        self.date = date
        self.group = group
        self.matches = matches
        self.minGames = minGames
        self.maxGames = maxGames
        self.allottedGames = allottedGames

    def availability(self):
        if self.maxGames is not None and len(self.matches) >= self.maxGames:
            return HIT_MAX
        elif self.minGames is None or len(self.matches) >= self.minGames:
            return HIT_MIN
        else:
            return UNDER_MIN

    def _id(self):
        return (self.date, self.group)

    def __cmp__(self, other):
        if not isinstance(other, DayData):
            return NotImplemented
        return cmp(self._id(), other._id())

    def __hash__(self):
        return hash(self._id())

    def __repr__(self):
        return '%s(%r, group=%r, matches=%r, minGames=%r, maxGames=%r, allottedGames=%r)' % (type(self).__name__,
                                             self.date,
                                             self.group,
                                             self.matches,
                                             self.minGames,
                                             self.maxGames,
                                             self.allottedGames)


    def dateInfoStr(self):
        infoStr = str(self.date)
        if self.group:
            infoStr = '%s (%s)' % (infoStr, self.group)
        return infoStr

class Match(object):
    def __init__(self, slot, player1, player2, round=None):
        self.slot = slot
        self.player1 = player1
        self.player2 = player2
        self.round = round

    @property
    def date(self):
        if self.slot is None:
            return None
        return self.slot.date

    def dateInfoStr(self):
        if self.date is None:
            return 'None'
        else:
            return self.slot.dateInfoStr()

    def _id(self):
        return (self.round, self.slot, self.player1, self.player2)

    def __cmp__(self, other):
        if not isinstance(other, Match):
            return NotImplemented
        return cmp(self._id(), other._id())

    def __hash__(self):
        return hash(self._id())

    def __contains__(self, player):
        return player == self.player1 or player == self.player2

    def __repr__(self):
        return '%s(%r, %r, %r, round=%r)' % (type(self).__name__,
                                             self.slot,
                                             self.player1,
                                             self.player2,
                                             self.round)

    def __str__(self):
        return '%s vs %s - Round %s - %s' % (self.player1, self.player2,
                                             self.round, self.dateInfoStr())

    def toDict(self):
        data = {}
        data['Date'] = self.date.strftime('%m/%d/%Y')
        data['Round'] = str(self.round)
        data['Player1'] = self.player1
        data['Player2'] = self.player2
        group = self.slot.group
        data['Type'] = '' if group is None else group
        return data

class ScheduleMaker(object):
    DEFAULT_SEED = 1
    SPREADSHEET_KEY = r'0Aj7bwN1ZgqdVdEFWc2Z3YWZHblQ4dnNtNjZIbU1zUmc'
    CHALLONGE_TOURNAMENT_ID = 175472
    MAX_LOGIN_TRIES = 3

    GAMES_PER_LEAGUE_NIGHT = 1
    GAMES_PER_LUNCH = 1

    # titles of worksheets
    ROSTER_WS_TITLE='Roster'
    SEASON_DATA_WS_TITLE = 'SeasonData'
    SCHEDULE_WS_TITLE = 'Schedule'

    ROSTER_DATA_TYPES = {
                         'challongeid':int,
                         'leaguenight':boolstr,
                         'lunch':boolstr,
                        }

    SEASON_DATA_TYPES = {
                         'startdate':_parseDate,
                         'enddate':_parseDate,
                         'gamespermatchup':int,
                         'leaguenights':_weekdayNum,
                        }
    SCHEDULE_DATA_TYPES = {
                           'date':_parseDate,
                           'round':int
                          }
    ERROR_ON_POOR_DISTRIBUTION = False

    SCHEDULE_COLUMNS = ('Date', 'Round', 'Player1', 'Player2', 'Type')

    def __init__(self, seed=DEFAULT_SEED, spreadsheetKey=SPREADSHEET_KEY):
        self.rand = random.Random()
        self.seed = seed
        self.rand.seed(self.seed)
        self.spreadsheetKey = spreadsheetKey
        self._client = None
        self._wsTitleKeyMap = {}

    @property
    def client(self):
        if self._client is None:
            self._client = gdata.spreadsheet.service.SpreadsheetsService()

            triesLeft = self.MAX_LOGIN_TRIES
            success = False
            while not success:
                login, pw = self.getGoogleLogPw()
                self._client.email = login
                self._client.password = pw
                try:
                    self._client.ProgrammaticLogin()
                    success = True
                except gdata.service.BadAuthentication:
                    triesLeft -= 1
                    if triesLeft <= 0:
                        raise
        return self._client

    def getGoogleLogPw(self):
        if hasattr(self, '_google_login'):
            login = self._google_login
        else:
            login = raw_input('Google Login: ').replace('\r', '')
        if hasattr(self, '_google_pw'):
            pw = self._google_pw
        else:
            pw = getpass.getpass('Google Password: ')
        return login, pw

    def setChallongeLogPw(self):
        oldUser, oldKey = challonge.get_credentials()
        if not oldUser or not oldKey:
            login = raw_input('Challonge Login: ').replace('\r', '')
            pw = getpass.getpass('Challonge API Key: ')
            challonge.set_credentials(login, pw)

    def getWorksheets(self):
        '''Returns a list of tuples, (title, url, key) for each worksheet
        '''
        worksheets = []
        for worksheet in self.client.GetWorksheetsFeed(self.spreadsheetKey).entry:
            url = worksheet.id.text
            worksheets.append((worksheet.title.text,
                               url, url.rsplit('/', 1)[-1]))
        return worksheets

    def renameWorksheet(self, oldName, newName):
        for worksheet in self.client.GetWorksheetsFeed(self.spreadsheetKey).entry:
            if worksheet.title.text == oldName:
                worksheet.title.text = newName
                self.client.UpdateWorksheet(worksheet)
                break
        else:
            return
        self._wsTitleKeyMap = {}

    def getWorksheetKey(self, title):
        '''Returns the key for the worksheet with the given title
        '''
        if title not in self._wsTitleKeyMap:
            for curTitle, _, key in self.getWorksheets():
                if title == curTitle:
                    break
            else:
                raise WorksheetNotFoundError('Unable to find worksheet titled %r' % title)
            self._wsTitleKeyMap[title] = key
            return key
        else:
            return self._wsTitleKeyMap[title]

    def _getWorksheetListFeed(self, title=None, key=None):
        if key is None:
            if title is None:
                raise ValueError("must give either worksheet title or key")
            key = self.getWorksheetKey(title)
        return self.client.GetListFeed(self.spreadsheetKey, key).entry

    def getWorksheetRows(self, title=None, key=None, types=None):
        rows = []
        for entry in self._getWorksheetListFeed(title=title, key=key):
            row = dict( zip( entry.custom.keys(), [ value.text for value in entry.custom.values() ] ) )
            if types:
                for key, val in row.iteritems():
                    if key in types:
                        row[key] = types[key](val)
            rows.append(row)
        return rows

    def getRoster(self):
        if getattr(self, '_roster', None) is None:
            roster = {}
            # Index by player name

            for row in self.getWorksheetRows(self.ROSTER_WS_TITLE,
                                             types=self.ROSTER_DATA_TYPES):
                roster[row['name']] = row
            self._roster = roster
            return roster
        else:
            return self._roster

    def getSeasonData(self):
        if getattr(self, '_seasonData', None) is None:
            seasonData = self.getWorksheetRows(self.SEASON_DATA_WS_TITLE,
                                               types=self.SEASON_DATA_TYPES)
            self._seasonData = seasonData[0]
        return self._seasonData

    def getSchedule(self):
        scheduleRows = self.getWorksheetRows(title=self.SCHEDULE_WS_TITLE)


    def idToPlayer(self, challongeId):
        if not hasattr(self, '_idToPlayer'):
            self._idToPlayer = dict( (player['challongeid'], player) for player in self.getRoster().itervalues())
        return self._idToPlayer[challongeId]

    def getChallongePlayers(self):
        self.setChallongeLogPw()
        return challonge.participants.index(self.CHALLONGE_TOURNAMENT_ID)

    def setGooglePlayerChallongeIds(self):
        cPlayers = self.getChallongePlayers()
        gPlayersRaw = self._getWorksheetListFeed(title=self.ROSTER_WS_TITLE)
        gPlayersByName = dict((row.custom['name'].text, row) for row in gPlayersRaw)
        gPlayersByEmail = dict((row.custom['lumaemail'].text, row) for row in gPlayersRaw)
        for cPlayer in cPlayers:
            # first match by name
            gPlayer = gPlayersByName.get(cPlayer['name'])
            # then match by email
            if not gPlayer:
                gPlayer = gPlayersByEmail.get(cPlayer['new-user-email'])

            if gPlayer:
                gPlayerRow = dict( (key, val.text) for key, val in gPlayer.custom.iteritems())
                gPlayerRow['challongeid'] = str(cPlayer['id'])
                print gPlayerRow
                self.client.UpdateRow(gPlayer, gPlayerRow)

    def getTimeSlots(self):
        if getattr(self, '_slots', None) is None:
            seasonData = self.getSeasonData()

            currentDate = seasonData['startdate']
            endDate = seasonData['enddate']
            leagueNight = seasonData['leaguenights']

            slots = []
            while currentDate <= endDate:
                weekday = currentDate.weekday()
                # weekdays start with monday=0, so friday=4
                if weekday <= 4:
                    slots.append(DayData(currentDate))
                    slots.append(DayData(currentDate,
                                             minGames=self.GAMES_PER_LUNCH,
                                             maxGames=self.GAMES_PER_LUNCH,
                                             group='lunch'))

                    if weekday == leagueNight:
                        slots.append(DayData(currentDate,
                                             minGames=self.GAMES_PER_LEAGUE_NIGHT,
                                             maxGames=self.GAMES_PER_LEAGUE_NIGHT,
                                             group='leagueNight'))
                currentDate += + datetime.timedelta(days=1)
            slots = [x for x in slots if x.maxGames is None or x.maxGames > 0]
            self._slots = slots
        return self._slots

    def getRoundMatches(self, service=CHALLONGE, cached=True):
        if not (isinstance(service, enum.EnumValue)
                and service.enumtype == Services):
            raise TypeError(service)
        if not cached or getattr(self, '_rounds', None) is None:
            if service == CHALLONGE:
                rounds = self.getChallongeMatches()
            elif service == GOOGLE:
                rounds = self.getGoogleMatches()
            else:
                raise ValueError(service)
            self._rounds = rounds
        return self._rounds

    def getChallongeMatches(self):
        self.setChallongeLogPw()
        matches = challonge.matches.index(self.CHALLONGE_TOURNAMENT_ID)
        rounds = {}
        for match in matches:
            roundNum = match['round']
            rounds.setdefault(roundNum, []).append( Match(None,
                                                          self.idToPlayer(match['player1-id'])['name'],
                                                          self.idToPlayer(match['player2-id'])['name']))
        return rounds

    def getGoogleMatches(self):
        rounds = {}
        allSlots = self.getTimeSlots()
        for row in self.getWorksheetRows(title=self.SCHEDULE_WS_TITLE,
                                         types=self.SCHEDULE_DATA_TYPES):
            roundNum = row['round']
            for slot in allSlots:
                if slot.group == row['type'] and slot.date == row['date']:
                    break
            else:
                raise ScheduleMakerError("could not find a time slot to match entry %r in google schedule" % row)
            rounds.setdefault(roundNum, []).append( Match(slot,
                                                          row['player1'],
                                                          row['player2'],
                                                          round=roundNum))
        return rounds

    def makeSchedule(self):
        timeSlots = {UNDER_MIN:[],
                     HIT_MIN:[],
                     HIT_MAX:[],
                    }

        allSlots = self.getTimeSlots()
        for timeSlot in allSlots:
            timeSlots[timeSlot.availability()].append(timeSlot)

        # First, we go through the process of allotting games - figuring out
        # how many games each day will have

        allottedGames = 0

        for slot in timeSlots[UNDER_MIN]:
            allottedGames += slot.minGames
            slot.allottedGames = slot.minGames

        # now that we've reserved the minimum number of games for days that
        # have minimums, find out how much more we can allot to each day

        # This holds lists of days, by the number of additional games that
        # can be allotted to them; 0 means they've hit their max, -1 means
        # they have no max, and n means they can add n more games before hitting
        # thier max
        remainingAllotment = {
                              0:list(timeSlots[HIT_MAX]),
                              -1:[],
                             }

        for slot in timeSlots[UNDER_MIN] + timeSlots[HIT_MIN]:
            if slot.maxGames is None:
                remainingAllotment[-1].append(slot)
            elif slot.minGames == slot.maxGames:
                remainingAllotment[0].append(slot)
            else:
                minGames = 0 if slot.minGames is None else slot.minGames
                gamesLeft = slot.maxGames - minGames
                remainingAllotment.setdefault(gamesLeft, []).append(slot)

        # Now, we need to allot games to days where we don't have an exact
        # requirement (ie, min != max).  For the most even distribution
        # possible, we want to end up with all remaining days having either
        # n or n + 1 games alloted to them (or have hit their max, if max < n)

        # To do this, we can imagine that we go through in "allotment rounds";
        # in each round, we look at how many days can take at least one more
        # game (D) and the number of games left to allot (N). If N >= D, we
        # allot one more game to all these days, then remove that number of
        # games from the number left to allot (N = N - D), and remove from D
        # any games which have now hit their maximum. If N < D, then we have
        # finished our allotment rounds - we know that N of the remaining days
        # will have one more game, and the rest will stand pat. Since we don't
        # want to actually portion out games just yet, we simply set the
        # maximum number of games for all these remaining days to their current
        # allotment, + 1

        # To go through the allotment rounds, it's actually easier to go
        # backwards: we know that if we kept doing allotment rounds forever,
        # the final round(s) would consist of only those dates with no max;
        # then, assuming the max key value in remainingAllotment is M,
        # we know that on the Mth round, we would also have the days in
        # remainingAllotment[M]; on the M-1 round, we would additionally
        # have any days in remainingAllotment[M-1]... and so on...
        currentAvailable = list(remainingAllotment[-1])
        maxRound = max(remainingAllotment) + 1
        availableByRound = {maxRound:list(currentAvailable)}
        currentRound = maxRound - 1
        while currentRound > 0:
            currentAvailable += remainingAllotment[currentRound]
            availableByRound[currentRound] = list(currentAvailable)
            currentRound -= 1

        # now, go through forwards, as described above
        currentRound = 1
        roundMatches = self.getRoundMatches()
        totalGames = sum(len(x) for x in roundMatches.itervalues())
        gamesLeft = totalGames - allottedGames
        if gamesLeft < 0:
            raise ValueError("not enough games to fill the minimum number of games")
        numPlusOnes = 0
        while gamesLeft > 0:
            remainingSlots = availableByRound[min(currentRound, maxRound)]
            numRemainingSlots = len(remainingSlots)
            if gamesLeft >= numRemainingSlots:
                for slot in remainingSlots:
                    slot.allottedGames += 1
                    allottedGames += 1
            else:
                numPlusOnes = gamesLeft
                for slot in remainingSlots:
                    slot.maxGames = slot.allottedGames + 1
            gamesLeft -= numRemainingSlots
            currentRound += 1

        if numPlusOnes == 0:
            possiblePlusOnes = set()
        else:
            possiblePlusOnes = set(remainingSlots)

        # Now set the min games to the alloted
        for slot in allSlots:
            if slot.minGames < slot.allottedGames:
                slot.minGames = slot.allottedGames
            if slot.maxGames is None:
                slot.maxGames = slot.allottedGames
            else:
                # Max games should either be allottedGames, or allottedGames + 1
                assert slot.maxGames in (slot.allottedGames, slot.allottedGames + 1)

        # print out some info - mostly for debugging
        print "total games:", totalGames
        print "numPlusOnes:", numPlusOnes

        # Now that we know we have a certain number of "+1" slots, distribute
        # those +1 days evenly between the various rounds
        plusOnesPerRound = numPlusOnes // len(roundMatches)
        extraPlusOnes = numPlusOnes % len(roundMatches)
        roundPlusOneAllotments = dict( (roundNum, plusOnesPerRound)
                                       for roundNum in roundMatches )

        # Now that we know the number of games to be played each day (to within
        # one game), go through each round, and figure out how many games from
        # that round will be played in each time slot

        # reverse remainingSlots, so we can efficiently pop from the end
        remainingSlots = list(reversed(allSlots))
        currentDayMatchesUsed = 0
        currentDayRoundMatchesUsed = 0
        # for the initial currentDay, just set it to a dummy value, such that
        # the initial check to get a new day will be triggred
        currentDay = DayData(None, minGames=0, maxGames=0)
        roundSlots = dict( (roundNum, []) for roundNum in roundMatches )

        for roundNum in sorted(roundMatches):
            matches = roundMatches[roundNum]
            currentDayRoundMatchesUsed = 0
            matchesLeft = len(matches)
            while matchesLeft > 0:
                # we don't automatically always pop a new day for the current
                # day, because it's possible that we have a "partially used"
                # day left over from the last round...
                if currentDayMatchesUsed >= currentDay.maxGames:
                    currentDay = remainingSlots.pop()
                    currentDayMatchesUsed = 0
                    currentDayRoundMatchesUsed = 0

                minLeft = currentDay.minGames - currentDayMatchesUsed
                if matchesLeft >= minLeft:
                    currentDayRoundMatchesUsed += minLeft
                    matchesLeft -= minLeft
                    currentDayMatchesUsed += minLeft

                    if currentDay.maxGames > currentDay.minGames:
                        # if this slot is a possible +1, decide if we want to
                        # use it - first, check that we have matches left (if
                        # not, we obviously can't use the extra day); then
                        # check if we've used up our allotted mandatory +1s (if
                        # not, we must use a +1); then, see if the number of
                        # slots left that can take a +1 is equal to the number
                        # of extraPlusOnes left (if so, we must use a +1);
                        # finally, if the other checks failed, this +1 is
                        # optional; use it only if doing so would mean we
                        # exactly finish up this round + day

                        usePlusOne = False
                        if roundPlusOneAllotments[roundNum]:
                            # we have remaining mandatory allotments
                            usePlusOne = True
                            roundPlusOneAllotments[roundNum] -= 1
                        else:
                            # We don't have mandatory allotments - check if
                            # the number of possible games left that can take
                            # +1s is equal to the number of +1s left
                            possiblePlusOnes.intersection_update(remainingSlots)
                            remainingMandatoryPlusOnes = 0
                            nextRound = roundNum + 1
                            while nextRound in roundPlusOneAllotments:
                                remainingMandatoryPlusOnes += roundPlusOneAllotments[nextRound]
                                nextRound += 1

                            # The + 1 is to account for the currentDay
                            totalRemainingPlusOnes = len(possiblePlusOnes) + 1
                            remainingOptionalPlusOnes = totalRemainingPlusOnes - remainingMandatoryPlusOnes

                            if not extraPlusOnes:
                                usePlusOne = False
                            elif extraPlusOnes == remainingOptionalPlusOnes:
                                # we don't have enough slots left NOT to use
                                # our +1s!
                                usePlusOne = True
                            elif not matchesLeft:
                                # if there are no matchesLeft, then not using
                                # a +1 means we will end the round "exactly"
                                # on a day - a good thing! don't use the +1
                                usePlusOne = False
                            elif matchesLeft == 1:
                                # conversely, if there is exactly one match
                                # left, then using the +1 would end the
                                # round exactly on the day
                                usePlusOne = True
                            else:
                                # Whether to use the +1 at this point is truly
                                # optional - so decide randomly, with the
                                # probability of using it being equal to the
                                # ratio of remaining plus ones to remaining
                                # possible plus ones
                                if self.rand.random() <= float(extraPlusOnes) / remainingOptionalPlusOnes:
                                    usePlusOne = True
                                else:
                                    usePlusOne = False

                            if usePlusOne:
                                extraPlusOnes -= 1

                        if usePlusOne:
                            currentDay.minGames += 1
                            currentDayMatchesUsed += 1
                            currentDayRoundMatchesUsed += 1
                            matchesLeft -= 1
                        else:
                            currentDay.maxGames -= 1
                else:
                    # The matches left for this round is less than the minimum
                    # for the day - the day will have to be split between two
                    # rounds...
                    currentDayMatchesUsed += matchesLeft
                    currentDayRoundMatchesUsed += matchesLeft
                    matchesLeft = 0

                # Add in the day info for this round
                roundSlots[roundNum].append( [currentDay, currentDayRoundMatchesUsed] )
                currentDayRoundMatchesUsed = 0

        # Double check that our allotment is correct
        assert not remainingSlots
        slottedGames = 0
        for slot in allSlots:
            assert(slot.minGames == slot.maxGames)
            slottedGames += slot.minGames
        assert slottedGames == totalGames

        # We should now have exactly figured out how many games wil be played
        # on each day, for each round; we can now go about assigning actual
        # games to those days
        # (...we didn't just go ahead and do this at the same time we were
        # figuring out the numbers because we may have "special rules" for
        # allotting games to special nights - ie, leagueNights, lunchGames)

        # want to make a copy of roundMatches
        # Note that since I'm going to be removing items from the lists, need
        # to make a copy of the lists too
        roundMatchesRemain = dict( (roundNum, list(matches))
                                   for roundNum, matches
                                   in roundMatches.iteritems() )

        # code for special rules for leagueNights, lunchGames, etc, goes here
        roster = self.getRoster()
        leagueNighters = [x['name'] for x in roster.itervalues() if x['leaguenight']]

        self.distributeMatches(roundSlots, roundMatchesRemain, 'leagueNight',
                               leagueNighters)
        lunchers = [x['name'] for x in roster.itervalues() if x['lunch']]
        self.distributeMatches(roundSlots, roundMatchesRemain, 'lunch',
                               lunchers)

        # We now just have "normal" matches left to assign - do it!
        for roundNum, slots in roundSlots.iteritems():
            matchesLeft = roundMatchesRemain[roundNum]
            for slot, numMatches in slots:
                for _ in xrange(numMatches):
                    # pick a random match remaining, add it to
                    # roundMatchesFinal
                    chosenIndex = self.rand.randrange(len(matchesLeft))
                    chosenMatch = matchesLeft.pop(chosenIndex)
                    self._updateMatch(chosenMatch, slot, roundNum)

        # do sanity check to ensure that everyone has played the same number
        # of games
        gamesPlayed = dict( (x,0) for x in roster )
        for slot in allSlots:
            for match in slot.matches:
                gamesPlayed[match.player1] += 1
                gamesPlayed[match.player2] += 1
        for player, games in gamesPlayed.iteritems():
            if games != (len(roster) - 1) * self.getSeasonData()['gamespermatchup']:
                raise DistributionError("player %s did not play correct number of games" % player)

        for matches in roundMatches.itervalues():
            # Now that we have dates, sort by them
            matches.sort()

        # print it out
        print
        for roundNum in sorted(roundMatches):
            for match in roundMatches[roundNum]:
                print match

        self.putScheduleInGoogle()
        self.putPlayerSchedulesInGoogle()

    def _updateMatch(self, match, slot, roundNum):
        match.slot = slot
        match.round = roundNum
        slot.matches.append(match)

    # FIXME: despite all the changes, this still isn't guaranteed to give
    # a properly-distributed result - as an example of a failure, I was using
    # LumaTT_Test,  with a player ordering / player seeding of:
    #    Some Guy
    #    Jason Fittipaldi
    #    Elaine Wu
    #    Ryan Sivley
    #    Brent Hensarling
    #    Alex Khan
    #    Kevin Curtin
    #    Thana Siripopungul
    #    Ruy Delgado
    #    Marcos Romero
    #    Nathan Rusch
    #    Chad Dombrova
    #    Richard Sutherland
    #    Sam Bourne
    #    Jared Simeth
    #    Lenny Gordon
    #    Raphael Pimentel
    #    Brandon Barney
    # and the following players opting OUT of night matches (in the google doc):
    #    Thana Siripopungul
    #    Sam Bourne
    # ...and the random number seed set to 5.
    # For now, simply going to choose new random seeds until I get a correct
    # result

    def distributeMatches(self, roundSlots, roundMatchesRemain, gameType, playerPool):
        '''Try to distribute the matches on slots of the given type evenly
        between the given players
        '''
        playerPool = set(playerPool)
        maxedPlayers = set()

        # count the total number of players that must be picked for all
        # games of the desired type - we can use this to determine the mininum
        # number of times a player must be picked
        totalPicks = 0
        for slots in roundSlots.itervalues():
            for slot, numMatches in slots:
                if slot.group != gameType:
                    continue
                # mult matches by 2, because 2 players per match
                totalPicks += (numMatches * 2)
        minPicks = totalPicks // len(playerPool)
        if totalPicks % len(playerPool):
            maxPicks = minPicks + 1
        else:
            maxPicks = minPicks


        playerPicks = dict( (player, 0) for player in playerPool )
        for roundNum in sorted(roundSlots):
            slots = roundSlots[roundNum]
            matchesLeft = roundMatchesRemain[roundNum]
            for slotI, (slot, numMatches) in enumerate(slots):
                if slot.group != gameType:
                    continue
                for _ in xrange(numMatches):
                    # Due to the fact that the playerPool is a subset of
                    # possible players, if we make picks purely randomly,
                    # it's possible to end up in a situation where we
                    # cannot make further picks to ensure an even
                    # distribution (ie, say in our next to last round,
                    # players A, B, and C all have 2 games, and everybody
                    # else has 3. We then pick a game between A and B;
                    # then, in the final round, to ensure even
                    # distribution, we would need to pick C's game - but in
                    # the final round, C plays a game against somebody not
                    # in the playerPool for this game type!

                    # to help get around this, we will keep track of, for each
                    # round and player, how many remaining games that
                    # player might play, then use that value as a
                    # "tiebreaker" (and also to determine if a player MUST be
                    # picked)

                    gamesRemaining = dict( (player, 0) for player in playerPool)
                    for futureRound in xrange(roundNum + 1, max(roundSlots) + 1):
                        futureMatchesLeft = roundMatchesRemain[futureRound]
                        for match in futureMatchesLeft:
                            if (match.player1 not in playerPool
                                    or match.player2 not in playerPool
                                    or match.player1 in maxedPlayers
                                    or match.player2 in maxedPlayers):
                                continue
                            gamesRemaining[match.player1] += 1
                            gamesRemaining[match.player2] += 1

                    # Find the list of "required" players - that is, players
                    # for whom minPicks - picks >= gamesRemaining
                    requiredPlayers = set()
                    for player in playerPool:
                        if (minPicks - playerPicks[player] >= gamesRemaining[player]
                                and player not in maxedPlayers):
                            requiredPlayers.add(player)

                    matchesByRequired = {}
                    for match in matchesLeft:
                        requiredCount = 0
                        if match.player1 in requiredPlayers:
                            requiredCount += 1
                        if match.player2 in requiredPlayers:
                            requiredCount += 1
                        matchesByRequired.setdefault(requiredCount, []).append(match)
                    potentialMatches = sorted(matchesByRequired.items())[-1][1]

                    # Now, find all the remaining matches between eligible players,
                    # and sort them into lists based on the total number of picks
                    # between the two players
                    def filterByPlayersInPool(potentialMatches):
                        matchesByPicks = {}
                        for match in potentialMatches:
                            p1Picks = playerPicks.get(match.player1, None)
                            p2Picks = playerPicks.get(match.player2, None)
                            if p1Picks is None or p2Picks is None:
                                continue
                            matchesByPicks.setdefault(p1Picks + p2Picks, []).append(match)

                        # If we didn't find ANY matches between eligible players, panic
                        if not matchesByPicks:
                            raise DistributionError("no matches left for round %d that contain players for %s matches" % (roundNum, gameType))

                        # otherwise, check out the minimum group
                        potentialMatches = sorted(matchesByPicks.items())[0][1]
                        return potentialMatches

                    try:
                        potentialMatches = filterByPlayersInPool(potentialMatches)
                    except DistributionError:
                        if self.ERROR_ON_POOR_DISTRIBUTION:
                            raise
                        potentialMatches = filterByPlayersInPool(matchesLeft)

                    # Now do similarly to how we sorted matches by lowest
                    # number of picks earlier - except sort by future
                    # games remaining now
                    matchesByFuture = {}
                    for match in potentialMatches:
                        p1Future = gamesRemaining[match.player1]
                        p2Future = gamesRemaining[match.player2]
                        matchesByFuture.setdefault(p1Future + p2Future, []).append(match)
                    potentialMatches = sorted(matchesByFuture.items())[0][1]

                    chosenMatch = self.rand.choice(potentialMatches)
                    matchesLeft.remove(chosenMatch)
                    self._updateMatch(chosenMatch, slot, roundNum)

                    # see if either player has hit the max - if so, remove
                    # them from the player pool
                    for player in (chosenMatch.player1, chosenMatch.player2):
                        picks = playerPicks[player] + 1
                        if picks == maxPicks:
                            maxedPlayers.add(player)
                        else:
                            playerPicks[player] = picks

                slots[slotI] = [slot, 0]

        # Now check that we got an even distribution
        # regenerate playerPicks "just to be certain"
        playerPicks = dict( (x,0) for x in playerPool )
        for slots in roundSlots.itervalues():
            for slot, _ in slots:
                if slot.group != gameType:
                    continue
                for match in slot.matches:
                    playerPicks[match.player1] += 1
                    playerPicks[match.player2] += 1
        if max(playerPicks.itervalues()) - min(playerPicks.itervalues()) > 1:
            from pprint import pprint
            pprint(playerPicks)
            msg = "%s games not distributed properly" % gameType
            if self.ERROR_ON_POOR_DISTRIBUTION:
                raise DistributionError(msg)
            else:
                print 'WARNING!:',
                print msg

    def fillTable(self, wsKey, headers, dataRows):
        numCols = len(headers)
        numRows = len(dataRows) + 1

        # For speed, do this using a batch request, instead of a bunch of
        # UpdateCell or InsertRow calls
        batchRequest = gdata.spreadsheet.SpreadsheetsCellsFeed()

        query = gdata.spreadsheet.service.CellQuery()
        query.min_col = '1'
        query.max_col = str(numCols)
        query.min_row = '1'
        query.max_row = str(numRows)
        query.return_empty = 'true'
        cells = self.client.GetCellsFeed(self.SPREADSHEET_KEY, wksht_id=wsKey,
                                         query=query)

        def setCell(rowI, colI, val):
            entry = cells.entry[rowI * numCols + colI]
            entry.cell.inputValue = val
            batchRequest.AddUpdate(entry)

        for colI, header in enumerate(self.SCHEDULE_COLUMNS):
            setCell(0, colI, header)

        for rowI, row in enumerate(dataRows):
            for colI, header in enumerate(headers):
                setCell(rowI + 1, colI, row[header])

        updated = self.client.ExecuteBatch(batchRequest, cells.GetBatchLink().href)
        if not updated:
            raise SpreadsheetUpdateError("Error updating spreadsheet")

    def putScheduleInGoogle(self):
        # First check if a 'schedule' worksheet exists
        worksheets = self.getWorksheets()
        titles = set(x[0] for x in worksheets)
        if self.SCHEDULE_WS_TITLE in titles:
            # back it up
            basename = '%s - %s backup' % (self.SCHEDULE_WS_TITLE,
                                           datetime.date.today())
            name = basename
            num = 1
            while name in titles:
                num += 1
                name = '%s %d' % (basename, num)
            self.renameWorksheet(self.SCHEDULE_WS_TITLE, name)

        numCols = len(self.SCHEDULE_COLUMNS)

        roundMatches = self.getRoundMatches()
        numMatches = sum(len(matches) for matches in roundMatches.itervalues())
        numRows = numMatches + 1
        self.client.AddWorksheet(self.SCHEDULE_WS_TITLE,
                                 numRows,
                                 numCols,
                                 self.SPREADSHEET_KEY)
        wsKey = self.getWorksheetKey(self.SCHEDULE_WS_TITLE)

        dataRows = []
        for roundNum in sorted(roundMatches):
            for match in roundMatches[roundNum]:
                dataRows.append(match.toDict())
        self.fillTable(wsKey, self.SCHEDULE_COLUMNS, dataRows)

    def getPlayerSchedules(self):
        roster = self.getRoster()
        playerSchedules = dict( (player, []) for player in roster)
        roundMatches = self.getRoundMatches()
        for roundNum in sorted(roundMatches):
            for match in roundMatches[roundNum]:
                playerSchedules[match.player1].append(match)
                playerSchedules[match.player2].append(match)
        return playerSchedules

    def putPlayerSchedulesInGoogle(self):
        for player, matches in self.getPlayerSchedules().iteritems():
            title = '%s - %s' % (self.SCHEDULE_WS_TITLE, player)
            numCols = len(self.SCHEDULE_COLUMNS)
            numRows = len(matches) + 1
            try:
                wsKey = self.getWorksheetKey(title)
            except WorksheetNotFoundError:
                self.client.AddWorksheet(title,
                         numRows,
                         numCols,
                         self.SPREADSHEET_KEY)
                wsKey = self.getWorksheetKey(title)

            self.fillTable(wsKey, self.SCHEDULE_COLUMNS,
                           [x.toDict() for x in matches])

if __name__ == '__main__':
    # run a test
    sm = ScheduleMaker()
    sm.makeSchedule()

