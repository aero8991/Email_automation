import datetime as dt
import os
from win32com.client import Dispatch



networked_directory = 'Determine file path'
file = os.path.join(networked_directory,"filename.csv")
today = dt.date.today()

#Search through my inbox. Note  win32com uses 6 as a users inbox.
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder("6")
messages = inbox.Items
olApp = Dispatch("Outlook.Application")

def save_attachment():
    for message in messages:
        if message.Subject == 'Enter Subject name here' and message.Senton.date() == today:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(networked_directory,str(attachment)))
                if message.Unread:
                    message.Unread = False
                    
                break

 #optional email notification                  
def Email():
    #note, using win32 to email a notification to a user/users
    #enter users seperated by ; and Importance level 2 sets the email important to high or (!)
    msg = olApp.CreateItem(0)
    msg.To = 'user@user.com' 
    msg.Subject = "Name of your email"
    msg.BodyFormat = 2
    msg.CC = 'user@user.com'
    msg.Importance = 2
    msg.HTMLBody = """ <p>Enter HTML code here</p>"""
   
    
    try:
        msg.Send()
        print('email sent!')
        print(msg.SentOn)
    except Exception as ex:
        print(ex, "Email Message was not sent!")            
            
        
def main():
    if os.path.exists(file_test) == False:
        save_attachment()
        Mail_PS_Link() 
        print("File copied for processing! ")
    else:
        print("File Already exists! Will not copy over existing file")
    

if __name__=="__main__":
    main()       
        
