import win32com.client
import ctypes # for the VM_QUIT to stop PumpMessage()
import pythoncom
import time
import psutil
import requests
import google.auth.transport.requests
from google.oauth2 import service_account
import tkinter as tk
import json
from tkinter import messagebox as mb

class Handler_Class(object):

    def __init__(self):
        # First action to do when using the class in the DispatchWithEvents    
        inbox = self.Application.GetNamespace("MAPI").GetDefaultFolder(6)
        messages = inbox.Items
        # Check for unread emails when starting the event
        Scopes = ['https://www.googleapis.com/auth/cloud-platform']
        SERVICE_ACCOUNT_FILE = 'My First Project-8abdd3d807a6.json'
        self.cred = service_account.Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=Scopes)
        self.auth_req = google.auth.transport.requests.Request()
       
        for message in messages:
            if message.UnRead:
                print("unread")
             
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
            body = mail.Body
            text = subject + " "+ body
            URL = "https://automl.googleapis.com/v1/projects/407905356473/locations/us-central1/models/TCN1472406598289719296:predict"
            data = {'payload': {
                        'textSnippet': {
                            'content': text,
                            'mime_type': 'text/plain'
                            }
                        }
                }
            if not self.cred.valid:
                self.cred.refresh(self.auth_req)
            token = 'Bearer' + ' '+ self.cred.token
            http_headers = {'Authorization': token,
                       'Content-Type': 'application/json'}
           
            resp = requests.post(url = URL, data = json.dumps(data,indent=2), headers = http_headers)
            payload = json.loads(resp.text)['payload']
            sorted_payload = sorted(payload, key = lambda i: i['classification']['score'],reverse=True)
            classification = sorted_payload[0]['displayName']
            print(classification)
            tk.Tk().withdraw()
            tk.messagebox.showwarning(title='Important', message='An important email has been received, please read')
           
 
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
