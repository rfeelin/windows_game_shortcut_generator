import sys, os, glob, shutil, time, re, stat, subprocess, winshell, json, urllib2, pprint
from win32com.client import Dispatch

#Variable stuff
project64Path = ""
n64RomsFolder = ""
#End variable stuff

acceptedFileTypes = ['z64', '.n64']
newShortcuts = []
existingShortcuts = []

valid_chars = '-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'

#Open shortcutsDump
if os.path.isfile(shortcutsDump):
    with open(shortcutsDump, 'r') as fp:
        existingShortcuts = json.load(fp)
else:
    with open(shortcutsDump, 'w') as fp:
        fp.write('[]')


#Scan for new stuff
print '****New Rom****'
newtargets = []
for gameFolder in n64RomsFolder:
    for root, subFolders, files in os.walk(gameFolder, topdown=False):
        for fileName in files:
            if any(fileName.find(acceptedFileType) != -1 for acceptedFileType in acceptedFileTypes):
                target = os.path.abspath(os.path.join(root, fileName))
                if any(existingShortcut['target'] == target for existingShortcut in existingShortcuts):
                    newtargets.append(target)
