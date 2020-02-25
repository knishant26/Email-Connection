# -*- coding: utf-8 -*-
"""
Created on Thu Feb  12 22:35:14 2020

@author: knishant
"""

import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

your_folder = mapi.Folders['YOUR ROOT FOLDER NAME'].Folders['SUB FOLDER']#.Folders["Important"]
messages = inbox.Items #Get all emails in the folder
message = messages.GetLast() #Get Latest Email

#Iterating over the mails in the folder
for message in your_folder.Items:
    print(message.Subject) #Get the Subject
    print(message.Body) #Get the Body of the Mail 
