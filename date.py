import win32com.client 
import os, time

outlook=win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNameSpace("MAPI")
your_folder = mapi.Folders['jhershey@tpm-hawaii.com'].Folders['Inbox'].Folders['test']


def replaceSubjectLine(email:object):
        #print("starting subject = " + email.Subject)
        #print("sent on date = " + str(email.Subject) )
        #email.Subject = str(message.senton.date())
        #email.Save
        bad_chars = [':', '!', '*','&','/','?']
        
        try:
                date = str(email.senton.date())

        except:
                date = str("unknown date")
                
        new = 'C:/Users/JHershey/Documents/SortDate/' + date + '.msg'
        old = email.Subject
        for i in bad_chars :
                old = old.replace(i, '') #' '
        old = 'C:/Users/JHershey/Documents/SortDate/' + old + '.msg'

           
        try:
                os.rename(old,new)
        except:
                ph_old = old
                ph_new = new
                x = 1
                print('not found old names = ' + old)
                while x < 50:
                        temp_insert = '-' + str(x) #temp_insert = ' (' + str(x) + ')' 
                        temp_insert2 = ' (' + str(x) + ')'
                        idx = ph_old.index('.msg')
                        old = ph_old[:idx] + temp_insert + ph_old[idx:]
                        old2 = ph_old[:idx] + temp_insert2 + ph_old[idx:]
                        #old = 'C:\\Users\\JHershey\\Documents\\SortDate\\' + 'FW  Insured  Koha Foods  Claim No  20168612 (1)' + '.msg'
                        x+=1
                        try:
                                os.rename(old,new)
                                break;
                        
                        except FileNotFoundError:
                                try:
                                        os.rename(old2,new)
                                except:
                                        pass
                                

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
                                                print('looking for a name old = ' + old)
                                                print('looking for a name new = ' + new)

        
#old = 'C:\Users\JHershey\Documents\SortDate\FW Insured  Koha Foods; Claim No. 20168612-3.msg'
#new = 'C:\Users\JHershey\Documents\SortDate\newdate2016.msg' + str(message.senton.date())
#os.rename(r'C:\Users\JHershey\Documents\SortDate\FW Insured  Koha Foods; Claim No. 20168612-3.msg', r'C:\Users\JHershey\Documents\SortDate\' )


for email in your_folder.Items:
        replaceSubjectLine(email)
  




    

    
'''message2=message.GetLast() 
subject=message2.Subject 
body=message2.body 
date=message2.senton.date()    
sender=message2.Sender 
attachments=message2.Attachments 
print(date)'''  
