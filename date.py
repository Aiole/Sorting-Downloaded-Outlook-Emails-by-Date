import win32com.client 
import os, time
outlook=win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNameSpace("MAPI")
your_folder = mapi.Folders['jhershey@tpm-hawaii.com'].Folders['Inbox'].Folders['test']


def replaceSubjectLine(email:object):
        #print("starting subject = " + email.Subject)
        #print("sent on date = " + str(message.senton.date()) )
        #email.Subject = str(message.senton.date())
        #email.Save
        bad_chars = [':', '!', '*','&']
        new = 'C:/Users/JHershey/Documents/SortDate/' + str(message.senton.date()) + '.msg'
        old = email.Subject
        for i in bad_chars :
                old = old.replace(i, '') #' '
        old = 'C:/Users/JHershey/Documents/SortDate/' + old + '.msg'
        print('OLD = ' + old)

           
        try:
                os.rename(old,new)
        except:
                ph_old = old
                ph_new = new
                x = 1
                while x < 50:
                        temp_insert = '-' + str(x) #temp_insert = ' (' + str(x) + ')' 
                        idx = ph_old.index('.msg')
                        old = ph_old[:idx] + temp_insert + ph_old[idx:]
                        #old = 'C:\\Users\\JHershey\\Documents\\SortDate\\' + 'FW  Insured  Koha Foods  Claim No  20168612 (1)' + '.msg'
                        x+=1
                        try:
                                os.rename(old,new)
                                break;
                        
                        except FileNotFoundError:
                                print('not found old name = ' + old)
                                

                        except:
                                j = 1
                                while j < 50:
                                        temp_insert = ' (' + str(j) + ')' 
                                        idx = ph_new.index('.msg')
                                        new = ph_new[:idx] + temp_insert + ph_new[idx:]
                                        j+=1
                                        try:
                                                os.rename(old,new)
                                                break;
                                        except:
                                                print('looking for a name new = ' + new)

        
#old = 'C:\Users\JHershey\Documents\SortDate\FW Insured  Koha Foods; Claim No. 20168612-3.msg'
#new = 'C:\Users\JHershey\Documents\SortDate\newdate2016.msg' + str(message.senton.date())
#os.rename(r'C:\Users\JHershey\Documents\SortDate\FW Insured  Koha Foods; Claim No. 20168612-3.msg', r'C:\Users\JHershey\Documents\SortDate\' )


for message in your_folder.Items:
    replaceSubjectLine(message)
  




    

    
'''message2=message.GetLast() 
subject=message2.Subject 
body=message2.body 
date=message2.senton.date()    
sender=message2.Sender 
attachments=message2.Attachments 
print(date)'''  
