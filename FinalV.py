import win32com.client
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

def Menu():
  print('\n-----------------------------------------------------------Menu-----------------------------------------------------------------\n')
  print('If you want to respond with new created emails automatically, please type 1.\n')
  print('If you want to respond with new created emails and save all the drafts in the drafts folder to update later on, please type 2.\n')
  print('If you want to respond with new created emails and display all the drafts at once, please type 3.\n')
  print('If you want to respond with new created emails and display 1 draft at a time, please type 4.\n')
  print('If you want to respond with replies to answer the emails automatically, please type 5.\n')
  print('If you want to respond with replies and save all the drafts in the drafts folder to update later on, please type 6.\n')
  print('If you want to respond with replies and display all the drafts at once, please type 7.\n')
  print('If you want to respond with replies and display 1 draft at a time, please type 8.\n')
  print('If you want to check the emails of the new applicants, please type 9.\n')
  print('----------------------------------------------------------------------------------------------------------------------------------\n')
  Choice = int(input('Please choose an option from the above choices: '))
  return (Choice)

#To check available emails
print('The available emails for usage are: ')
for account in mapi.Accounts:
	print('--> '+account.DeliveryStore.DisplayName)
print('\n')

email = input('Please enter the email that you want to use: ')

#In case of multiple users in the outlook app
inbox = mapi.Folders(email).Folders('Inbox')

#Check emails in the inbox folder (n = 6) - only 1 user
#inbox = mapi.GetDefaultFolder(6)

messages = inbox.Items

Choice = Menu()

if Choice == 1 or  Choice == 2 or  Choice == 3 or  Choice == 4:

    subject = input('\nPlease enter the Subject of the Email to send to the new applicants: ')
    Body = input('\nPlease input the Body of the Email to send to the new applicants: ')
    AttChoice = int(input('\nPlease type 1 if you want to attach a document, otherwise type 2: '))

    if AttChoice == 1:
        AttName = input('\nPlease enter the name of your attachment in addition to its extension: ')

    #Date restriction
    time = input('\nPlease enter the duration you want to check in hours: ')
    received_dt = datetime.now() - timedelta(hours=float(time)) #seconds, minutes, hours, days, weeks
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

    #Email restriction
    #messages = messages.Restrict("[SenderEmailAddress] = 'contact@codeforests.com'")

    #Subject restriction exactly 
    #messages = messages.Restrict("[Subject] = 'Prayer Times'")

    #Subject restriction including keywords
    messages = messages.Restrict("@SQL=(urn:schemas:httpmail:subject LIKE '%Alex%' or urn:schemas:httpmail:subject LIKE '%Max%'  )")

    Sent = 0

    print('\nThe email addresses of the senders are: \n')
    #Response Creation 
    for message in messages:
        if message.Class==43:
            #Get the senders Emails
            if message.SenderEmailType=='EX':
                print ('--> '+message.Sender.GetExchangeUser().PrimarySmtpAddress)
                recipient = message.Sender.GetExchangeUser().PrimarySmtpAddress
                
                #Create email
                olmailitem=0x0 #size of the new email
                ToBeSent=outlook.CreateItem(olmailitem)
                
                #The body of the response
                ToBeSent.Body = Body
                
                #The subject of the response
                ToBeSent.Subject = subject
                
                #The recipient (Email senders)
                ToBeSent.To = 'a.abdou@aui.ma'
                
                #In case of an attachement
                if AttChoice == 1:
                    ToBeSent.Attachments.Add(os.path.join(os.getcwd(), AttName))
                
    #Send the emails directly
                if Choice == 1:
                    ToBeSent.Send()
                    print('Email was sent to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' successfully\n')
                    Sent += 1
                
    #Save the emails in the drafts folder
                elif Choice == 2:
                    ToBeSent.Save()
                    print('Email to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' was saved in the draft folder\n')
                    Sent += 1
                
    #Display All drafts at once
                elif Choice == 3:
                    ToBeSent.Display(False)
                    print('Email to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' popped out successfully \n')
                    Sent += 1
                
    #Display One draft at a time
                elif Choice == 4:  
                    ToBeSent.Display(True)
                    print('Email to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' popped out successfully \n')
                    Sent += 1
                    
                else: 
                    print('\nPlease try again and choose an accurate choice from the offered chocies.')
                    quit()  
                
            else:
                print (message.SenderEmailAddress)

    print('\nThe number of applicants is: '+str(messages.Count)+ ' and the number of answered applicants is: '+str(Sent))

elif Choice == 5 or Choice == 6 or Choice == 7 or Choice == 8:
    
    Body = input('\nPlease input the Body of the Reply to send to the new applicants: ')
    AttChoice = int(input('\nPlease type 1 if you want to attach a document, otherwise type 2: '))

    if AttChoice == 1:
        AttName = input('\nPlease enter the name of your attachment in addition to its extension: ')

    #Date restriction
    time = input('\nPlease enter the duration you want to check in hours: ')
    received_dt = datetime.now() - timedelta(hours=float(time)) #seconds, minutes, hours, days, weeks
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

    #Email restriction
    #messages = messages.Restrict("[SenderEmailAddress] = 'contact@codeforests.com'")

    #Subject restriction exactly 
    #messages = messages.Restrict("[Subject] = 'Prayer Times'")

    #Subject restriction including keywords
    messages = messages.Restrict("@SQL=(urn:schemas:httpmail:subject LIKE '%Alex%' or urn:schemas:httpmail:subject LIKE '%Max%'  )")

    Sent = 0

    print('\nThe email addresses of the senders are: \n')
    #Response Creation 
    for message in messages:
        if message.Class==43:
            #Get the senders Emails
            if message.SenderEmailType=='EX':
                print ('--> '+message.Sender.GetExchangeUser().PrimarySmtpAddress)
                recipient = message.Sender.GetExchangeUser().PrimarySmtpAddress
                    
    #Send the replies directly
                if Choice == 5:
                    reply = message.Reply() 
                    reply.Body = Body
                    if AttChoice == 1:
                        reply.Attachments.Add(os.path.join(os.getcwd(), AttName))
                    reply.Send()
                    print('Reply was sent to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' successfully\n')
                    Sent += 1
                
    #Save the replies in the drafts folder
                elif Choice == 6:
                    reply = message.Reply() 
                    reply.Body = Body
                    if AttChoice == 1:
                        reply.Attachments.Add(os.path.join(os.getcwd(), AttName))
                    reply.Save()
                    print('Reply to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' was saved in the draft folder\n')
                    Sent += 1
                
    #Display All drafted replies at once
                elif Choice == 7:
                    reply = message.Reply() 
                    reply.Body = Body
                    if AttChoice == 1:
                        reply.Attachments.Add(os.path.join(os.getcwd(), AttName))
                    reply.Display(False)
                    print('Reply to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' popped out successfully \n')
                    Sent += 1
                
    #Display One drafted reply at a time
                elif Choice == 8:
                    reply = message.Reply() 
                    reply.Body = Body
                    if AttChoice == 1:
                        reply.Attachments.Add(os.path.join(os.getcwd(), AttName))
                    reply.Display(True)
                    print('Reply to: '+message.Sender.GetExchangeUser().PrimarySmtpAddress+ ' popped out successfully \n')
                    Sent += 1        
                
                elif Choice == 9:
                    print() 
                    
                else: 
                    print('\nPlease try again and choose an accurate choice from the offered chocies.')
                    quit()  
                
            else:
                print (message.SenderEmailAddress)

    print('\nThe number of applicants is: '+str(messages.Count)+ ' and the number of answered applicants is: '+str(Sent))

elif Choice == 9:
    time = input('\nPlease enter the duration you want to check in hours: ')
    received_dt = datetime.now() - timedelta(hours=float(time)) #seconds, minutes, hours, days, weeks
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

    #Email restriction
    #messages = messages.Restrict("[SenderEmailAddress] = 'contact@codeforests.com'")

    #Subject restriction exactly 
    #messages = messages.Restrict("[Subject] = 'Prayer Times'")

    #Subject restriction including keywords
    messages = messages.Restrict("@SQL=(urn:schemas:httpmail:subject LIKE '%Alex%' or urn:schemas:httpmail:subject LIKE '%Max%'  )")
    for message in messages:
        if message.Class==43:
            #Get the senders Emails
            if message.SenderEmailType=='EX':
                print ('--> '+message.Sender.GetExchangeUser().PrimarySmtpAddress)
                recipient = message.Sender.GetExchangeUser().PrimarySmtpAddress
    print('\nThe number of applicants is: '+str(messages.Count))    
       
else: 
    print('\nPlease try again and choose an accurate choice from the offered options.')
    quit()               
