import os
import time
import random
from random import randint
import glob
import subprocess
import win32com.client as comclt
import copy

wsh= comclt.Dispatch("WScript.Shell") #for .SendKeys()

#BEGIN SETTINGS

enableDebug = True
playAllGamesBeforeRepeating = True

#Play time ranges in seconds
minPlayTime = 5
maxPlayTime = 10

#mappings for configs. These correlate to array positions in the configs.
platformName = 0
emulator = 1
gamePath = 2
gameExtensions = 3
games = 4
saveStateKey = 5
loadStateKey = 6
loadStateDelay = 7

global configs
configs = [
["NES", "Z:\Emulators\emusanity\\nestopia\\nestopia.exe", "Z:\ROMs & ISOs\\NES1\\", ["nes"], [], "+1", "1", 1],
#["N64", "Z:\Emulators\emusanity\Project64\Project64.exe", "Z:\ROMs & ISOs\\N64\\", ["z64"], [], "{F5}", "{F7}", 2],
##["Z:\ROMs & ISOs\SNES", "Z:\ROMs & ISOs\SNES", ["smc"], []]
#["Genesis", "Z:\Emulators\emusanity\Kega Fusion\Fusion.exe", "Z:\ROMs & ISOs\Genesis\\", ["gen"], [], "{F5}", "{F8}",1],
["GameCube", "Z:\Emulators\emusanity\Dolphin-x64\Dolphin.exe", "Z:\ROMs & ISOs\Gamecube\\", ["gcm"], [], "+{F1}", "{F1}", 3]
]

#END SETTINGS DO NOT EDIT BELOW HERE

#Begin Functions
def debug(debugMessage):
    if enableDebug:
        print(debugMessage)

#End Functions

print("Populating rom lists...")
for platform in configs:
    os.chdir(platform[gamePath])
    for extension in platform[gameExtensions]:
        print ("Checking for " + platform[platformName] + " games...")
        platform[games] += glob.glob('*.' + extension)
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
        time.sleep(1) #Sleep for 1 second
        playTime -= 1 #decrement playTime by 1 until it reaches zero
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
