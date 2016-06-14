import sys
import imaplib
import getpass
import email
import email.header
import email.utils
import datetime
import base64
import smtplib
import os.path
import ConfigParser

def sendMail(subject, message, To):
    msg = "\r\n".join(["Subject: " + subject,"",message])
    smtp = smtplib.SMTP('smtp.' + Config.get('overig', 'serverName') +'.com:587') # Verbind met Gmail
    smtp.starttls()
    smtp.login(Config.get('s1', 'usrnm'), Config.get('s1', 'passwrd')) # Login bij Gmail
    smtp.sendmail(Config.get('s1', 'usrnm'), To, msg) # Verstuur de mail
    smtp.quit

def subjectCheck(subject, count = 0):
    if count >= len(subjectsp):
        return False
    if subject == Config.get('sort', subjectsp[count]):
        return True
    else:
        return False | subjectCheck(subject, count + 1)


path = os.getcwd() # current path
Config = ConfigParser.ConfigParser()
Config.read(path + "/config.INI") # less het config bestand
Config.sections()
subjectsp = Config.options('sort') # selecteer de sort sectie in het config file 
print "start"
#execfile("MailSort.conf") # Selecteer .config file
M = imaplib.IMAP4_SSL('imap.gmail.com') # Verbinging met IMAP van Gmail
M.login(Config.get('s1', 'usrnm'), Config.get('s1', 'passwrd')) # Login bij Gmail
M.select() # Kies de inbox
typ, data = M.search(None, 'ALL') # Laad de mailbox in

if typ != 'OK': # Geen 'OK' -> lege inbox
    print "Geen mail."

if typ == 'OK': # Wel 'OK' -> emails
    for num in data:#[0].split(): # Voor elk bericht in de inbox:
        typ, data = M.fetch(num, '(RFC822)')
        print "mail nummer" + num 
        mail = email.message_from_string(data[0][1]) # Selecteer de mail
        receiver = email.utils.parseaddr(mail['From']) # Zender van de te sorteren mail
        decodeHeader = email.header.decode_header(mail['Subject'])[0] # Selecteer onderwerp
        subject = unicode(decodeHeader[0]) # Zet onderwerp om in ASCII
        if(subjectCheck(subject)): # Als subject begint met de waarde van subj (in config.INI)
            hasSubject = False
            attBase64 = ""
            filename = ""
            print "subject is goed"
            for p in mail.walk(): # Loop door de email parts
                c_type = p.get_content_type() # Haal de content part van de verschillende parts
                c_disp = p.get('Content-Disposition') # Haal de content disposition (attachment deel van de mail)
                if c_disp != None and c_type ==  'text/plain' : 
                    hasSubject = True
                    attBase64 = p.get_payload() # Haal encoded attachment binnen
                    filename = p.get_filename() # Haal de filename op
                    attDecode = base64.b64decode (attBase64) # Decode attachment
                    filePath = path + Config.get('overig', 'savePath')
                    completeName = os.filePath.join(filePath, filename) # Set savepath voor filename
                    text_file = open(completeName, 'w') # Maak een nieuw .txt bestand met filename
                    text_file.write(attDecode) # Schrijf de inhoud van attachment
                    text_file.close() # Sluit bestand

            if hasSubject == True: # Wanneer er een attachment is met *.txt:
                print "stuur bevestiging"   
                sendMail(Config.get('overig','bevestigingSubject') + subject, Config.get('overig', 'bevestigingMsg'), receiver)
                
            else:
                print "stuur error"
                sendMail(Config.get('overig','errorSubject') + subject, Config.get('overig', 'errorMsg'), receiver)

        print "Deleted"
        #M.store(num, '+FLAGS', '\\Deleted') # Verwijder de berichten na uitvoeren programma

M.close()
M.logout()


	









