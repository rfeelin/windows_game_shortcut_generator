import sys, os, glob, shutil, time, re, mutagen, stat, subprocess, winshell, json, urllib2, pprint
from win32com.client import Dispatch

#Varible stuffs
drmFreeGameFolders = [os.path.normpath('P:\Games\Non-Steam'), os.path.normpath('P:\Games\Pirate')]
#drmFreeGameFolders = ['']
steamGameFolders = [os.path.normpath('P:\Games\Steam\steamapps'),os.path.normpath('C:\Program Files (x86)\Steam\steamapps')]
#steamGameFolders = [os.path.normpath('P:\Games\Steam\steamapps')]

steamIdApiUrl = 'http://store.steampowered.com/api/appdetails?appids='
steamShortcutBase = 'steam://rungameid/'

shortcutFolder = os.path.normpath('P:\Games\Game Shortcuts')
shortcutSteamImageFolder =  os.path.join(shortcutFolder, 'SteamIcons')
shortcutsDump = os.path.join(shortcutFolder, 'ShortcutsDump.json')

acceptedFileTypes = ['.exe']
newShortcuts = []
existingShortcuts = []


def clean_up(fix_me):
    fix_me = fix_me.replace('<',' ')
    fix_me = fix_me.replace('>',' ')
    fix_me = fix_me.replace(':','')
    fix_me = fix_me.replace('"','')
    fix_me = fix_me.replace('/',' ')
    fix_me = fix_me.replace("\\",' ')
    fix_me = fix_me.replace('|',' ')
    fix_me = fix_me.replace('?',' ')
    fix_me = fix_me.replace('*',' ')
    fix_me = fix_me.strip()
    return fix_me


#Open shortcutsDump
if os.path.isfile(shortcutsDump):
    with open(shortcutsDump, 'r') as fp:
        existingShortcuts = json.load(fp)
else:
    with open(shortcutsDump, 'w') as fp:
        fp.write('[]')

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
    for root, subFolders, files in os.walk(gameFolder, topdown=False):
        for fileName in files:
            for acceptedFileType in acceptedFileTypes: 
                if(fileName.find(acceptedFileType) != -1):
                    target = os.path.abspath(os.path.join(root, fileName))
                    targetExists = False
                    for existingShortcut in existingShortcuts:
                        if os.path.normpath(existingShortcut['target']) == target :
                            targetExists = True
                    if targetExists == False:
                        print target
                        newtargets.append(target)


#Get names for shortcuts/ignore
for target in newtargets:
    targetFolders = target.split('\\')
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
    newShortcut['workingDirectory'] = target.rpartition('\\')[0]
    if( nameOption > 0 and nameOption < 8 ):
        shortcutName = targetFolders[nameOption-1] + '.lnk'
        shortcutName = clean_up(shortcutName)
        newShortcut['shortcut'] = os.path.join(shortcutFolder, shortcutName)
    elif (nameOption == 8):
        shortcutName = raw_input('Enter Custom Name: ') + '.lnk'
        shortcutName = clean_up(shortcutName)
        newShortcut['shortcut'] = os.path.join(shortcutFolder, shortcutName)
    elif (nameOption == 9):
        newShortcut['shortcut'] = ''
    if nameOption > 0:
        newShortcuts.append(newShortcut)


#STEAM
steamGameIds = []
newSteamGameIds = []
for gameFolder in steamGameFolders:
    for fileName in os.listdir(gameFolder):
        if(fileName.find('.acf') != -1):
            targetExists = False
            steamGameId = fileName.strip('appmanifest_')
            steamGameId = steamGameId.strip('.acf')
            steamGameIds.append(steamGameId)
            for existingShortcut in existingShortcuts:
                if existingShortcut['target'] == (steamShortcutBase + steamGameId) :
                    targetExists = True
            if targetExists == False:
                newSteamGameIds.append(steamGameId)

#clean up existingShortcuts
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
        shortcutName = clean_up(shortcutName)
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
        print 'no data for: ' + str(steamGameInfo )


#Generate shortcuts
print '****Creating****'
shell = Dispatch('WScript.Shell')
for newShortcut in newShortcuts:
    if newShortcut['shortcut']:
        print newShortcut['shortcut'] + ' : Shortcut'
        print newShortcut['target'] + ' : Target'
        if newShortcut['workingDirectory']:
            shortcut = shell.CreateShortCut(newShortcut['shortcut'])
            shortcut.Targetpath = newShortcut['target']
            shortcut.WorkingDirectory = newShortcut['workingDirectory']
            shortcut.save()
        else:
            shortcut = file(newShortcut['shortcut'], 'w')
            shortcut.write('[InternetShortcut]\n')
            shortcut.write('URL=%s' % newShortcut['target'])
            shortcut.close()


#Update shortcutsDump 
combinedShortcuts = existingShortcuts + newShortcuts
with open(shortcutsDump, 'w') as fp:
    json.dump(combinedShortcuts, fp)