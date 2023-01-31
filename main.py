import os
import win32com.client

# Connecting to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
root_folder = inbox.Folders(6)
messages = root_folder.Items



