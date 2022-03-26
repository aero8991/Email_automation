import os
import win32com.client as wc
import pandas as pd
import datetime as dt

#use this area to adjust the length of the time in the inbox search. Example used is 7 days. 
last_week = dt.date.today()- dt.timedelta(days = 7)
last_week = last_week.strftime('%m/%d/%Y %H:%M %p')


#Search through the inbox
outlook = wc.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")

#add ".Folders['name_goes_here]" to go deeper within the inbox
root_folder = mapi.Folders['username'].Folders['Inbox']
print(root_folder.Name)

messages = root_folder.Items
messages = messages.Restrict("[ReceivedTime] >= '" + last_week + "'")
messages.Sort("[ReceivedTime]", True)

#grab message and convert to pandas DF
counter = 0
for message in messages:
    if message.Subject == 'Enter Subject of Search here':
        print(counter, message.Subject)
        html_str = message.HTMLBody
        get_table = pd.read_html(html_str)[0]
        df = get_table.to_dict()
        get_table['time_sent'] = message.SentOn.strftime("%m-%d-%y")
        if counter == 0:  
            # choose to do an action with the results. For example create a running file.. Note on 2nd run to change mode to 'a'
            # get_table.to_csv("File.csv", index=False)
            counter += 1 
        else:
            # choose to do an action with the results. For example create a running file..
            #get_table.to_csv("File.csv", mode='a', index=False, header=False)
            counter += 1
        
print("got results!")
