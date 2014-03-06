import time
import glob
import random
import os
from random import randint
import win32com.client
import win32api

debug = False
recyclePlayedGames = False

#Name of Settings File
settings = "settings.txt"
emulators = "emulators.txt"

#Variable to store current position in settins file so parser can go back to it
filePos = 0
#Variable to save position of tag in settings file so parser can go back to it
savedTagPos = 0

def getFromFile(filename, fieldType, lookFor):
    file = open(filename, 'r')

    file.seek(filePos)

    #Encase lookFor in angle brackets if looking for a tag
    if fieldType == 0 or fieldType == 2:
        lookFor = '<' + lookFor + '>'
        if fieldType == 2:
            endLookFor = lookFor[:1] + '/' + lookFor[1:]
        
    ##Use while loop to read file
    ##for loop caused an OSError with file.tell()
    while True:
        line = " "
        if fieldType != 2:
            line = file.readline()
        if not line:
            break

        #Look for a tag
        if fieldType == 0:
            if lookFor in line:
                curFilePos = file.tell()
                return curFilePos
                
        #Look for a setting
        elif fieldType == 1:
            if lookFor in line:
                #seperate the field from the parameter
                line = line.split(":")[1]
                #Remove whitespace
                line = line.strip()
                #Return the parameter to be used
                if debug:
                    print ("Variable Get: " + line)
                return line
            
        #Get all settings from within a tag
        elif fieldType == 2:
            #TODO: Scan tag from start to end for fields and variables, store them in an array(s) using split. Return this array
            fieldArray = []
            parameterArray = []
            while endLookFor not in line:
                line = file.readline()
                if endLookFor not in line:
                    #if debug:
                    print (line)
                    tempArray = line.split()
                    fieldArray.append(tempArray[0][:-1]) #removes ':' from end of field
                    parameterArray.append(tempArray[1])
            return [fieldArray,parameterArray] #return fields and parameters as a two-dimensional array

def parse2DArray(array, keepTerm, sideToCompare, sideToKeep):
    pos = 0
    parsedArray = []
    while pos < len(array[0]):
        if array[sideToCompare][pos] == keepTerm:
            parsedArray.append(array[sideToKeep][pos])
        pos += 1
    return parsedArray

#Get the mode from settings file. Can be "random" or "constant"
mode = getFromFile(settings, 1, "mode")

#save position the random or constant tags depending on the mode set
filePos = getFromFile(settings, 0, mode)
savedTagPos = filePos

if mode == "random":
    
    filePos = savedTagPos #Go back to <random> tag
    filePos = getFromFile(settings, 0, "starting-play-time")
    playTimeMin = getFromFile(settings, 1, "min")
    playTimeMax = getFromFile(settings, 1, "max")

    filePos = savedTagPos #Go back to <random> tag
    filePos = getFromFile(settings, 0, "number-of-games-until-time-decrease")
    numGamesUntilTimeDecrease = getFromFile(settings, 1, "min")
    numGamesUntilTimeDecrease = getFromFile(settings, 1, "max")

    filePos = savedTagPos #Go back to <random> tag
    filePos = getFromFile(settings, 0, "decrease-time-by")
    decreaseTimeByMin = getFromFile(settings, 1, "min")
    decreaseTimeByMax = getFromFile(settings, 1, "max")

elif mode == "constant":
    filePos = savedTagPos #Go back to <constant> tag
    playTime = getFromFile(settings, 1, "starting-play-time")
    numGameUntilDecrease = getFromFile(settings, 1, "number-of-games-until-decrease")
    decreaseTimeBy = getFromFile(settings, 1, "decrease-time-by")

filePos = 0
filePos = getFromFile(settings, 0, "variables")
absoluteMinPlayTime = decreaseTimeByMin = getFromFile(settings, 1, "absolute-min-play-time")

filePos = getFromFile(settings, 0, "emulators-to-use")
emulatorsToUse = getFromFile(settings, 2, "emulators-to-use")

filePos = getFromFile(settings, 0, "save-slots-to-use")
saveSlotsToUse = getFromFile(settings, 2, "save-slots-to-use")

emulatorsToUse = parse2DArray(emulatorsToUse, "yes", 1, 0)

print (emulatorsToUse)
print (saveSlotsToUse)

pos = 0
emulatorArray = []
romArray = []

while pos < len(emulatorsToUse):
    index = saveSlotsToUse[0].index(emulatorsToUse[pos])
    emulatorArray.append([saveSlotsToUse[0][pos], [saveSlotsToUse[1][pos]]])
    pos += 1

emulatorParameters = []
pos = 0
filePos = 0
while pos < len(emulatorArray):
    tempArray2 = []
    print (str(emulatorArray[pos][1]))
    filePos = getFromFile(emulators, 0 , str(emulatorArray[pos][0]))
    print (filePos)
    filePos = getFromFile(emulators, 0, "directories")
    tempArray2.append( emulatorArray[pos][0] )
    tempArray2.extend( getFromFile(emulators, 2, "directories")[1])
    filePos = 0
    filePos = getFromFile(emulators, 0, emulatorArray[pos][0])
    tempArray2.append(getFromFile(emulators,1,"rom-extension"))
    emulatorName = "slot-" + str(emulatorArray[pos][1][0])
    print (emulatorName)
    filePos = getFromFile(emulators, 0, emulatorName)
    tempArray2.extend( getFromFile(emulators, 2, emulatorName)[1] )
    tempArray2.append([])
    emulatorParameters.append( tempArray2 )
    filePos = 0
    pos += 1
#Result: 2 dimensional array of emulators with their paths, save extension, save/load keys
#emulator name, rom directory, save directory, save extension, save state key, load state key
print (emulatorParameters)

shell = win32com. client.Dispatch("WScript.Shell")

#if mode == random:
    
while True:
    print ("ohai")
    if mode == random:
        playTime = randint(playTimeMin, playTimeMax)
        numGamesUntilDecrease = randint(numGamesUntilDecreaseMin, numGamesUntilDecreaseMax)
        decreaseTimeBy = randint(decreaseTimeByMin, decreaseTimeByMax)
        if (playTimeMin >= absoluteMinPlayTime and numGamesUntilDecrease < 1):
            decreaseTimeBy = randint(decreaseTimeByMin, decreaseTimeByMax)
            playTimeMin -= decreaseTimeBy
            playTimeMax -= decreaseTimeBy
            while (playTimeMax < 1 or playTimeMin < 1):
                playTimeMin += 1
                playTimeMax +=1
            numGamesUntilDecrease = randint(numGamesUntilDecreaseMin, numGamesUntilDecreaseMax)
        else:
            numGamesUntilDecrease -= 1
            
    playTime = randint(int(playTimeMin), int(playTimeMax))

    emulatorToUse = randint(0,len(emulatorParameters)-1)
    print (emulatorToUse)
    emulatorName = emulatorParameters[emulatorToUse][0]
    romDirectory = emulatorParameters[emulatorToUse][1]
    saveStateDirectory = emulatorParameters[emulatorToUse][2]
    romExtension = emulatorParameters[emulatorToUse][3]
    saveStateExtension = emulatorParameters[emulatorToUse][4]
    saveStateKey = emulatorParameters[emulatorToUse][5]
    loadStateKey = emulatorParameters[emulatorToUse][6]
    romArray = emulatorParameters[emulatorToUse][7]
    print (romDirectory)

    shell.AppActivate(emulatorName)

    #Populate Rom Array
    os.chdir(romDirectory)
    if len(romArray) < 1:
        romArray = glob.glob('*' + romExtension)
        print (romArray)
        if len(romArray) < 1:
            print("**Error: No files.")
            exit(1)

    romIndex = random.randrange(0, len(romArray))
    selectedRom = romArray[romIndex]
    if not recyclePlayedGames:
        del romArray[romIndex]

    #Generate file name to be used for save state
    #Truncate the file extension of the rom so  that a save state extension can be appended
    saveStateName = selectedRom[:len(selectedRom)-4]
    #Append file extension to saveStateName
    saveStateName += saveStateExtension
    if debug:
        print('Save name is: ' + saveStateName)
    
    #Open Rom
    os.startfile(selectedRom)
    
    #Check if save state exists
    os.chdir(saveStateDirectory)
    if len(glob.glob(saveStateName)) > 0:   
        saveStateExists = True
        if debug:
            print ("saveStateExists = True")
    else:
        saveStateExists = False
        if debug:
            print("saveStateExists = False")
    
    time.sleep(0.2)
    
    #Load State
    #If a save state already exists
    if saveStateExists:
        if debug:
            print ("Load State Name" + saveStateName)
        #Send keystrokes to emulator to load state
        shell.SendKeys(loadStateKey)
        #If a save state does not already exist
    else:
        #Create a dummy savestate file so the date can still be compared for when waitForSaveDateChange is run
        open(saveStateName, 'a').close()
    saveDate = time.ctime(os.path.getmtime(saveStateName))

    if debug:
        print ("Time Until Change: " + str(timeUntilChange))
    time.sleep(playTime)

    if debug:
        print ("Save State Name" + saveStateName)
        #Send keystrokes to emulator to save state
    shell.SendKeys(saveStateKey)
    os.chdir(saveStateDirectory)
    print (saveStateDirectory)
    #while saveDate == time.ctime(os.path.getmtime(saveStateName)):
    #    pass
    print ("hello after")
    
    emulatorParameters[emulatorToUse][7] = romArray
