import os
import time
import random
from random import randint
import glob
import subprocess
import win32com.client as comclt
import copy
import re
import win32api
import winsound

wsh= comclt.Dispatch("WScript.Shell") #for .SendKeys()
scriptDirectory = os.getcwd()

#mappings for configs. These correlate to array positions in the configs.
platformName = 0
emulator = 1
gamePath = 2
gameExtensions = 3
games = 4
saveStateKey = 5
loadStateKey = 6
loadStateDelay = 7

#BEGIN SETTINGS

#pause.wav
#resume.wav
#skipgame.wav
#savestate.wav

enableDebug = False
playAllGamesBeforeRepeating = True

#Play time ranges in seconds
minPlayTime = 5
maxPlayTime = 8

skipKey = ord('S') #Key to listen for when skipping game
pauseKey = ord('P')
resumeKey = ord('R')

global configs
configs = [
["NES", "G:\Emulators\emusanity\\nestopia\\nestopia.exe", "G:\\ROMs & ISOs\\NES\\", ["nes"], [], "+1", "1", 1],
["SNES", "G:\Emulators\emusanity\snes9x\snes9x-x64.exe", "G:\ROMs & ISOs\SNES\\", ["smc"], [], "1", "{F1}", 1]
["N64", "G:\Emulators\emusanity\Project64\Project64.exe", "G:\\ROMs & ISOs\\Nintendo 64\\USA\\", ["z64"], [], "{F5}", "{F7}", 2],
["Genesis", "G:\Emulators\emusanity\Kega Fusion\Fusion.exe", "G:\\ROMs & ISOs\\Genesis\\", ["gen"], [], "{F5}", "{F8}",1],
#["GameCube", "Z:\Emulators\emusanity\Dolphin-x64\Dolphin.exe", "Z:\ROMs & ISOs\\Gamecube\\", ["gcm"], [], "+{F1}", "{F1}", 3]
["GameBoy", "G:\Emulators\emusanity\VisualBoyAdvanceM1229\\VisualBoyAdvance-M.exe", "G:\\ROMs & ISOs\\Game Boy\\USA\\", ["gb"], [], "+1", "1", 1],
["GameBoy Color", "G:\Emulators\emusanity\VisualBoyAdvanceM1229\\VisualBoyAdvance-M.exe", "G:\\ROMs & ISOs\\Game Boy Color\\USA\\", ["gbc"], [], "+1", "1", 1]
]

#END SETTINGS DO NOT EDIT BELOW HERE

#Begin Functions
def debug(debugMessage):
    if enableDebug:
        print(debugMessage)

def isKeyPressed(key):
    #"if the high-order bit is 1, the key is down; otherwise, it is up."
    return (win32api.GetKeyState(key) & (1 << 7)) != 0

def skipGame():
    global playTime
    print ("Game skip triggered, hold key to confirm")
    time.sleep(1)
    if isKeyPressed(skipKey):
        print ("Game skip confirmed. Skipping current game.")
        playTime = 0
        playSound("skipgame.wav")
    else:
        playTime -= 1

def playSound(filename): #takes the filename of a wav file as a parameter
    try: #Attempt to play sound
        winsound.PlaySound(scriptDirectory + "\\" + filename, winsound.SND_FILENAME) #Play wav file in function parameter
    except: #If sound could not be played
        print(filename + " was not found")

#End Functions
exclude = "" #Create an empty string to store exclusions

exclusions = [line.rstrip('\n') for line in open('exclude.txt')] #Get all exclusion names from text file, put them into a list and trim off the newline character
for element in exclusions:
    exclude = exclude + '|' + element #Append all elements in list into ones tring with a '|' (regex or operator) between them
exclude = exclude[1:] #Trim off the first '|'
excludeRegex = re.compile(r'(?i)(\W|^)(' + exclude + ')(\W|$)') #Create a variable from the compiled regex pattern for excluding keywords

print("Populating rom lists...")
for platform in configs:
    os.chdir(platform[gamePath])
    for extension in platform[gameExtensions]:
        print ("Checking for " + platform[platformName] + " games...")
        platform[games] += glob.glob('*.' + extension)
print("Filtering rom lists...")
for platform in configs:
    gamesToRemove = []
    for game in platform[games]:
        if excludeRegex.search(game) is not None:
            print("REMOVING: " + game)
            gamesToRemove.append(game)
    for game in gamesToRemove:
        platform[games].remove(game)
print("Complete!")

global originalConfigs
if playAllGamesBeforeRepeating:
    originalConfigs = copy.deepcopy(configs) #save a copy of configs

while True:
    #load game
    platform = randint(0, len(configs)-1) #Select a platform (Ex: NES, SNES, Genesis)
    debug("platform = " + configs[platform][platformName])
    processName = configs[platform][emulator].rsplit('\\', 1)[1] #Grab process name from executable in path (Ex: nestopia.exe)
    debug("processName = " + processName)
    subprocess.call(['taskkill', '/im', processName, '/f'], shell=True) #Force kill process of currently selected platform in case it may already be open to avoid conflicts. Suppress console output
    gamePos = randint(0, len(configs[platform][games])-1) #Generate random number to select game from platform list with
    debug (gamePos)
    curGame = configs[platform][games][gamePos] #Use randomly generated number to select a game from the list of games and get it's filename
    debug("curGame = " + curGame)
    subprocess.Popen([configs[platform][emulator], configs[platform][gamePath] + configs[platform][games][gamePos]]) #Open game in the proper emulator
    time.sleep(configs[platform][loadStateDelay]) #Allow game time to open in emulator before loading state
    wsh.SendKeys(configs[platform][loadStateKey]) #Send key(s) to load state
    playTime = randint(minPlayTime, maxPlayTime) #Determine how long the next loaded game will be played for using a random number between defined ranges in settings
    debug("playTime = " + str(playTime))
    while playTime > 0: #stay in loop until playTime reaches zero
        if playTime == 3: #Save state when only three seconds remain
            wsh.SendKeys(configs[platform][saveStateKey]) #Send key(s) to save state
            playSound("savestate.wav")
            #time.sleep(0)
        time.sleep(1) #Sleep for 1 second
        playTime -= 1 #decrement playTime by 1 until it reaches zero

        #CHECK FOR KEYS HERE
        if isKeyPressed(pauseKey):
            stayOnCurrentGame = True
            playSound("pause.wav")
            print ("Pause triggered")
            while stayOnCurrentGame:
                if isKeyPressed(resumeKey):
                    stayOnCurrentGame = False
                    print ("Resuming")
                    playSound("resume.wav")
                elif isKeyPressed(skipKey):
                    stayOnCurrentGame = False
                    skipGame()
        
        if isKeyPressed(skipKey):
            skipGame()
        #END CHECK FOR KEYS

    subprocess.call(['taskkill', '/im', processName, '/f'], shell=True) #Force kill process of currently selected platform. Suppress console output

    if playAllGamesBeforeRepeating: #Go in here if we are not reusing games until they have all been played
        configs[platform][games].pop(gamePos) #Remove the game we just played from the game list for the current platform being played
        debug("Games Remaining = " + str(len(configs[platform][games])))
        if len(configs[platform][games]) < 1: #If game list has been depleted from a particular platform
            debug(configs)
            configs.pop(platform) #Remove the platform from the playlist
            debug(configs)
            if len(configs) < 1: #If all games from all platforms have been played and there are no platforms remaining
                debug(configs)
                configs = copy.deepcopy(originalConfigs) #copy the original config back, thus restoring all games
                debug(configs)
                print("Configs deplted. Restoring original configs")
