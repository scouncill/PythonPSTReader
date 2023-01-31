
import win32com.client
import pandas as pd

df = pd.DataFrame(columns = ['Sender', 'Sent', 'Subject', 'EntryID', 'MessageID', 'To','CC','BCC','Folder'])

def find_pst_folder(OutlookObj, pst_filepath) :
    print("Looking for " + pst_filepath)
    OutlookObj.AddStore(pst_filepath)
    for Store in OutlookObj.Stores :
        if Store.IsDataFileStore and Store.FilePath == pst_filepath :
            print("Found pst " + pst_filepath)
            return Store.GetRootFolder()
        print("Did not find " + pst_filepath)
    return None

#def enumerate_folders(FolderObj, DestObj) :
def enumerate_folders(FolderObj) :
    print("enumerating ")
    
    count_folders = 0
    for ChildFolder in FolderObj.Folders :
        print("Folder = " + str(ChildFolder))
        #enumerate_folders(ChildFolder, DestObj)
        enumerate_folders(ChildFolder)
        #count_folders = count_folders + 1
        #if count_folders >= 10: break

        #iterate_messages(FolderObj, DestObj)
    iterate_messages(FolderObj)


#def iterate_messages(FolderObj, DestObj) :
def iterate_messages(FolderObj) :
    global df
    print("iterating")
    count_messages = 0
    for item in FolderObj.Items :
        print("***************************************")
        print("Type = " + item.MessageClass)
        # https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.messageclass
        # https://learn.microsoft.com/en-us/office/vba/outlook/concepts/forms/item-types-and-message-classes

        if( item.MessageClass != "IPM.Note"): continue
        mySender = item.SenderName + " <" + item.SenderEmailAddress + ">"
        myDate = str(item.SentOn)
        mySubject = item.Subject
        myEntryID = item.EntryID
        myMessageID = item.propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E")
        # https://learn.microsoft.com/en-us/office/vba/api/outlook.propertyaccessor
        myPath = item.Parent.FullFolderPath

        myTo = ""
        myCc = ""
        myBcc = ""

        for xrecip in item.Recipients: 
            if(xrecip.Type == 1):
                # https://learn.microsoft.com/en-us/office/vba/api/outlook.olmailrecipienttype
                if(myTo ==""):
                    myTo = xrecip.Name + " <" + xrecip.Address + ">"
                else:
                    myTo = myTo + ", " + xrecip.Name + " <" + xrecip.Address + ">"
                    #myTo = myTo+xrecip.Name + " <" + xrecip.Address + ">"
            if(xrecip.Type == 2):
                if(myCc == ""):
                    myCc = xrecip.Name + " <" + xrecip.Address + ">"
                else:
                    myCc = myCc + ", " + xrecip.Name + " <" + xrecip.Address + ">"
            if(xrecip.Type == 3):
                if(myBcc ==""):
                    myBcc = xrecip.Name + " <" + xrecip.Address + ">"
                else:
                    myBcc = myBcc + ", " + xrecip.Name + " <" + xrecip.Address + ">"

        newline = {'Sender' : mySender, 'Sent' : myDate, 'Subject' : mySubject, 'EntryID' : myEntryID, 'MessageID': myMessageID, 'To': myTo,'CC':myCc,'BCC':myBcc, 'Folder': myPath}

        df_new_row = pd.DataFrame(newline, index=[0])
        df = pd.concat([df, df_new_row])


Outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

pst = r"C:\folder\myPSTfilename.pst"
Outlook.AddStore(pst)

#pst2= r"C:\folder\outfile.pst"
#Outlook.AddStore(pst)

PSTFolderObj = find_pst_folder(Outlook,pst)
#DestFolderObj = find_pst_folder(Outlook,pst2)


try :
    #enumerate_folders(PSTFolderObj, DestFolderObj )
    enumerate_folders(PSTFolderObj )
except Exception as exc :
    print(exc)
finally :
    print(df)
    df.to_excel("C:\folder\output.xlsx") 
    print("Done")
    #Outlook.RemoveStore(PSTFolderObj)





