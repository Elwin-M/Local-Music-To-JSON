#https://en.wikipedia.org/wiki/Audio_file_format
import win32com.client
import json
import time
import os
from tkinter import *
from tkinter import filedialog
from collections import OrderedDict

#grabs path -> change to take input
def getFilePath():
    labelPath["text"] = filedialog.askdirectory().replace('/', '\\')
    buttonDict.config(state='normal')
    textWidget.config(state='normal')
    textWidget.delete("1.0", END)
    textWidget.config(state='disabled')
    #return labelPath["text"]

#loops through path to get variables first, then grabs values from items
def getDetailsOfFiles(pathToFiles):
    #print(pathToFiles)
    sh=win32com.client.gencache.EnsureDispatch('Shell.Application',0)
    pathToFiles = sh.NameSpace(pathToFiles)
    colnum = 0
    columns = []
    dictOfFiles = {}
    itemCount = 0
    #getting variables
    while True:
        colname=pathToFiles.GetDetailsOf(None, colnum)
        if not colname:
            break
        columns.append(colname)
        colnum += 1

    #getting values
    for item in pathToFiles.Items():
        for colnum in range(len(columns)):
            colval=pathToFiles.GetDetailsOf(item, colnum)
        
            if columns[colnum]=='Name':
                fName = colval
            if  columns[colnum]=='Album':
                fAlbum = colval
            if  columns[colnum]=='Title':
                fTitle = colval.strip()
                if fTitle == '':
                    fTitle = fName.strip()
            if  columns[colnum]=='Authors':
                fAuthor = colval
            if  columns[colnum]=='Item type':
                fType = colval
        itemCount += 1
        #dictOfFiles[fName] ={ 'Title' : fTitle, 'Authors' : fAuthor, 'Album' : fAlbum, 'Type' : fType}
        dictOfFiles[itemCount] ={ 'File Name' : fName,'Title' : fTitle, 'Authors' : fAuthor, 'Album' : fAlbum, 'Type' : fType}
    dictOfFiles = OrderedDict(sorted(dictOfFiles.items(), key=lambda t:t[1]['Title'].lower()))
    return dictOfFiles

#removes keys from dict if type does not exist in array
def delUselessFiles(dictToDeleteFrom, arrayOfFormats):
    #arrayOfFormats = ['mp3', 'm4a']
    filesToDelete = []
    #checking if type value exists in array
    for x in dictToDeleteFrom:
        checkFor = dictToDeleteFrom[x]['Type'].lower()
        if any(ext in checkFor for ext in arrayOfFormats):
            pass
        else:
            filesToDelete.append(x)

    #deleting from dict
    for z in filesToDelete:
        del dictToDeleteFrom[z]

    return dictToDeleteFrom

#reduces dict to title and artist
def getShortDict(dictToReduce):
    shortDictOfFiles = {}
    num = 1
    for a in dictToReduce:
        shortDictOfFiles.update({num : {dictToReduce[a]['Title'] : dictToReduce[a]['Authors']}})
        num +=1

    #songsDict = shortDictOfFiles
    global songsDict
    songsDict = shortDictOfFiles
    return shortDictOfFiles

#for printing any dict
def printDict(dictToPrint):
    prettyResponse = json.dumps(dictToPrint, indent=4, sort_keys=True)
    textWidget.config(state='normal')
    textWidget.delete("1.0", END)
    textWidget.insert(END, 'Saving to file...')
    textWidget.insert(END, '\n')
    textWidget.insert(END, 'Please Be Patient')
    textWidget.insert(END, '\n')

    saveDict(songsDict)

    textWidget.insert(END, prettyResponse)
    textWidget.config(state='disabled')

    #buttonSave = Button(win, text="Save To File", state='normal', width=15, height=2, font=('Aerial 10'), command=saveDict(songsDict))
    #buttonSave.pack(pady=5, side=TOP, fill=X)

#for saving the dict to file
def saveDict(dictToSave):
    createDir()
    currentTime = time.localtime()
    currentTimeString = time.strftime('%Y-%m-%d %I-%M-%S %p', currentTime)

    textWidget.insert(END, 'Saving JSON')
    textWidget.insert(END, '\n')
    #print("Saving JSON")      
    with open("Songs JSON\Songs " + currentTimeString + ".json", 'w') as JSONDump:
         json.dump(dictToSave, JSONDump, indent=4)
    textWidget.insert(END, 'Saving JSON Finished')
    textWidget.insert(END, '\n \n')         

def createDir():
    try:
        os.mkdir('Songs JSON')
        textWidget.insert(END, 'Directory Songs JSON Created')
        textWidget.insert(END, '\n')  
    except FileExistsError:
        textWidget.insert(END, 'Directory Songs JSON Already Exists')
        textWidget.insert(END, '\n') 


def main():
    global win 
    global labelPath
    global scrollBar
    global textWidget
    global buttonDict
    global arrayOfFormats
    arrayOfFormats = ['mp3', 'm4a']
    arrayString = ', '.join(str(x) for x in arrayOfFormats)

    win = Tk()
    #Set the size of the tkinter window
    win.geometry("700x750")
    #create button widgets
    buttonPath = Button(win, text="Select Path", width=15, height=2, font=('Aerial 10'), command=getFilePath)
    buttonDict = Button(win, text="Display & Save Dict", state='disabled', width=15, height=2, font=('Aerial 10'), command=lambda: 
    #####for different use cases#####
    printDict(getShortDict(delUselessFiles(getDetailsOfFiles(labelPath["text"]), arrayOfFormats))))
    #printDict(delUselessFiles(getDetailsOfFiles(labelPath["text"]))))
    #printDict(getDetailsOfFiles(labelPath["text"])))
    #####for different use cases#####
    
    buttonPath.pack(pady=5, side=TOP, fill=X)
    buttonDict.pack(pady=5, side=TOP, fill=X)
    
    # Create a label widgets
    labelPath = Label(win, text="Path Here", font=('Aerial 13'))
    labelPath.pack(pady=5, side=TOP, fill=X)
    #create text widgets
    scrollBar = Scrollbar(win)
    scrollBar.pack(side=RIGHT, fill=Y)
    textWidget = Text(win, width=100, yscrollcommand=scrollBar.set)
    textWidget.pack(side=LEFT, fill=BOTH)
    textWidget.insert(END, '''
Select a directory.
It does not recursively search through. 
It will only grab the files in the same folder.
If the folder has hundreds of songs, the program might not respond. That's okay.
Keep waiting and it will be okay.
This message will dissapear if you hit the Display & Save button.
If you have more folders, you can select them afterwards.
The song names will be saved to a text file to be used with the other program.

If there are any issues, please let me know.

Right now, only ''' + arrayString + ' formats are supported')
    textWidget.config(state='disabled')
    
    win.mainloop()

main()