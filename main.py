import win32com.client
import ctypes # for the VM_QUIT to stop PumpMessage()
import pythoncom
import re
import time
import psutil

class Handler_Class(object):

    def __init__(self):
        # First action to do when using the class in the DispatchWithEvents     
        inbox = self.Application.GetNamespace("MAPI").GetDefaultFolder(6)
        messages = inbox.Items
        # Check for unread emails when starting the event
        for message in messages:
            if message.UnRead:
                print( message.Subject )# Or whatever code you wish to execute.
                print(message.Body)
    def OnQuit(self):
        # To stop PumpMessages() when Outlook Quit
        # Note: Not sure it works when disconnecting!!
        ctypes.windll.user32.PostQuitMessage(0)

    def OnNewMailEx(self, receivedItemsIDs):
    # RecrivedItemIDs is a collection of mail IDs separated by a ",".
    # You know, sometimes more than 1 mail is received at the same moment.
        for ID in receivedItemsIDs.split(","):
            mail = self.Session.GetItemFromID(ID)
            subject = mail.Subject
            print (subject)   
            try: 
                command = re.search(r"%(.*?)%", subject).group(1)
                print( command) # Or whatever code you wish to execute.
            except:
                pass

# Function to check if outlook is open
def check_outlook_open ():
    list_process = []
    for pid in psutil.pids():
        p = psutil.Process(pid)
        # Append to the list of process
        list_process.append(p.name())
    # If outlook open then return True
    if 'OUTLOOK.EXE' in list_process:
        return True
    else:
        return False

# Loop 
while True:
    try:
        outlook_open = check_outlook_open()
    except: 
        outlook_open = False
    # If outlook opened then it will start the DispatchWithEvents
    if outlook_open == True:
        outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
        pythoncom.PumpMessages()
    # To not check all the time (should increase 10 depending on your needs)
    time.sleep(10)
