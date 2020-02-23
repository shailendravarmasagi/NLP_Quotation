# -*- coding: utf-8 -*-
"""
Created on Sun Feb 23 20:02:07 2020

@author: Shailendra

"""
import win32com.client

def get_Mail_Messages():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder

    for folder in outlook.Folders:
        print("****Folder Name**********************************")
        print(folder)
    
        print("*************************************************")
        for folder1 in folder.Folders:
            if str(folder1)=='POC':
                req_folder=folder1
    Messages = req_folder.Items
    return(Messages)
#message = messages.GetFirst()
#body_content = message.HTMLBody