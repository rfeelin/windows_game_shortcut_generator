import sys, os, glob, shutil, time, re, stat, subprocess, json, urllib2, pprint
#import sys, os, glob, shutil, time, re, stat, subprocess, winshell, json, urllib2, pprint
#from win32com.client import Dispatch

#Variable stuff
drmFreeGameFolders = []
steamGameFolders = []
shortcutFolder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Game Shortcuts')
#End variable stuff

steamIdApiUrl = 'http://store.steampowered.com/api/appdetails?appids='
steamShortcutBase = 'steam://rungameid/'

shortcutSteamImageFolder =  os.path.join(shortcutFolder, 'SteamIcons')
shortcutsDump = os.path.join(shortcutFolder, 'ShortcutsDump.json')

acceptedFileTypes = ['.exe']
newShortcuts = []
existingShortcuts = []

valid_chars = '-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'

#Create Dir
if not os.path.exists(shortcutFolder):
    os.makedirs(shortcutFolder)

#Open shortcutsDump
if os.path.isfile(shortcutsDump):
    with open(shortcutsDump, 'r') as fp:
        existingShortcuts = json.load(fp)
else:
    shortcutsDumpFile = file(shortcutsDump, 'w')
    shortcutsDumpFile.write('[]')
    shortcutsDumpFile.close()

#Scan for games in game paths. Skip for existing in shortcutsDump
#DRM FREE
#clean up existingShortcuts
print '****Clean Up DRM Free****'
for existingShortcut in existingShortcuts:
    if existingShortcut['workingDirectory'] and not os.path.isfile(existingShortcut['target']):
        print existingShortcut['shortcut']
        os.remove(existingShortcut['shortcut'])
        existingShortcuts.remove(existingShortcut)

#scan for new stuff
print '****New DRM Free****'
newtargets = []
for gameFolder in drmFreeGameFolders:
    print gameFolder
    for root, subFolders, files in os.walk(gameFolder, topdown=False):
        for fileName in files:
            if any(fileName.find(acceptedFileType) != -1 for acceptedFileType in acceptedFileTypes):
                target = os.path.abspath(os.path.join(root, fileName))
                if all(existingShortcut['target'] != target for existingShortcut in existingShortcuts):
                    newtargets.append(target)

#Get names for shortcuts/ignore
for target in newtargets:
    targetFolders = os.path.split(target)
    targetFolders = targetFolders[-7:]
    print '*******'
    print target
    for index, folder in enumerate(targetFolders):
        print str(index+1) + ': ' + folder
    print '8: Custom Name'
    print '9: Ignore'
    print '0: Skip This Time'
    try:
        nameOption = int(raw_input('Enter: '))
    except ValueError:
        print "Not a number"
    newShortcut = {}
    newShortcut['target'] = target
    newShortcut['workingDirectory'] = os.path.split(target)[:-1]
    if( nameOption > 0 and nameOption < 8 ):
        shortcutName = targetFolders[nameOption-1] + '.lnk'
        shortcutName = ''.join(c for c in shortcutName if c in valid_chars).strip()
        newShortcut['shortcut'] = os.path.join(shortcutFolder, shortcutName)
    elif (nameOption == 8):
        shortcutName = raw_input('Enter Custom Name: ') + '.lnk'
        shortcutName = ''.join(c for c in shortcutName if c in valid_chars).strip()
        newShortcut['shortcut'] = os.path.join(shortcutFolder, shortcutName)
    elif (nameOption == 9):
        newShortcut['shortcut'] = ''
    if nameOption > 0:
        newShortcuts.append(newShortcut)


#Steam
steamGameIds = []
newSteamGameIds = []
for gameFolder in steamGameFolders:
    for fileName in os.listdir(gameFolder):
        if fileName.find('.acf') != -1:
            steamGameId = fileName.strip('appmanifest_')
            steamGameId = steamGameId.strip('.acf')
            steamGameIds.append(steamGameId)
            if all(existingShortcut['target'] != (steamShortcutBase + steamGameId) for existingShortcut in existingShortcuts):
                newSteamGameIds.append(steamGameId)

#Clean up existingShortcuts
print '****Clean Up Steam****'
for existingShortcut in existingShortcuts:
    exists = False
    for steamGameId in steamGameIds:
        if( (steamShortcutBase + steamGameId) == existingShortcut['target'] ) :
            exists = True
    if exists == False:
        print existingShortcut['shortcut']
        os.remove(existingShortcut['shortcut'])
        existingShortcuts.remove(existingShortcut)

print '****New Steam****'
for steamGameId in newSteamGameIds:
    print steamGameId
    steamGameInfoDump = urllib2.urlopen(steamIdApiUrl + steamGameId).read() 
    steamGameInfo = json.loads(steamGameInfoDump)
    try:
        shortcutName = steamGameInfo[steamGameId]['data']['name']
        shortcutName = ''.join(c for c in shortcutName if c in valid_chars).strip()
        print shortcutName
        #shortcutImageName = os.path.join(shortcutSteamImageFolder, (steamGameId + '.jpg') )

        #img = urllib2.urlopen( steamGameInfo[steamGameId]['data']['header_image'] )
        #with open(shortcutImageName, 'wb') as fp:
        #    fp.write( img.read() )
        #print shortcutImageName
        #urllib.urlopen( steamGameInfo[steamGameId]['data']['header_image'] ).read()

        newShortcut = {}

        newShortcut['target'] = steamShortcutBase + steamGameId 
        newShortcut['workingDirectory'] = ''
        newShortcut['shortcut'] = os.path.join(shortcutFolder, (shortcutName + '.url'))
        newShortcuts.append(newShortcut) 
    except:
        print 'no data for: ' + str(steamGameInfo)


#Generate shortcuts
print '****Creating****'
#shell = Dispatch('WScript.Shell')
for newShortcut in newShortcuts:
    if newShortcut['shortcut']:
        print newShortcut['shortcut'] + ' : Shortcut'
        print newShortcut['target'] + ' : Target'
        if newShortcut['workingDirectory']:
            shortcut = shell.CreateShortCut(newShortcut['shortcut'])
            shortcut.Targetpath = newShortcut['target']
            shortcut.WorkingDirectory = newShortcut['workingDirectory']
            shortcut.save()
            print 'reg Short Cut'
        else:
            shortcut = file(newShortcut['shortcut'], 'w')
            shortcut.write('[InternetShortcut]\n')
            shortcut.write('URL=%s' % newShortcut['target'])
            shortcut.close()
            print 'url shortcut'


#Update shortcutsDump 
combinedShortcuts = existingShortcuts + newShortcuts
with open(shortcutsDump, 'w') as fp:
    json.dump(combinedShortcuts, fp)