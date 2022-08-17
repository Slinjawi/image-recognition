import pip as pip
import pytesseract as pytesseract
from PIL import ImageGrab
import time
from PIL import ImageGrab, ImageEnhance
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


#Capture a Screenshot every 5 sec
#while True:
    #time.sleep(5)
    #img = ImageGrab.grab()

#Process the image
def get_time_in_queue(img):
    # pre-process image
    width, height = img.size
    img = img.crop((width*.3, height*.4, width*.7, height*.6))
    img = img.point(lambda p: p > 170 and 255)
    img = ImageEnhance.Color(img).enhance(0)

    # extract text
    imgdata = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    imgtext = ' '.join([x for i,x in enumerate(imgdata['text']) if int(imgdata['conf'][i]) >= 70])

    # extract minutes remaining
    imgtime = re.search(r'time[^\d+]*(\d+)', imgtext)
    if imgtime:
        return int(imgtime.group(1))

# function to send alert
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
def send_alert(alarm=True, email=False, playsong=False):
    if alarm:
        print(chr(7))

    if email:
        # set up the SMTP server
        from myinfo import outlook # note: you need to have created myinfo.py
        emailserver = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
        emailserver.starttls()
        emailserver.login(outlook.username, outlook.password)

        # set up the message
        msg = MIMEMultipart()
        msg['From'] = outlook.username
        msg['To'] = outlook.username
        msg['Subject'] = "Time to WoW!"
        message = 'Your WoW queue is about to pop!'
        msg.attach(MIMEText(message, 'plain'))

        # send the message
        emailserver.send_message(msg)
        print(f'Email notification sent to {outlook.username}')

    if playsong:
        songfile = 'C:\\NBUK26X-alarm'
        from playsound import playsound
        playsound(songfile)

# final function to monitor the queue

def queue_alert(alertat=2, interval=5, occurences=3, alarm=True, emailalert = False, playsong = False):
    print('\nScanning for queue. Hit CTRL+C to exit.')
    counter = 0
    while True:
        # capture screenshot of the queue
        time.sleep(interval)
        img = ImageGrab.grab()

        # extract the queue time
        minutes = get_time_in_queue(img)

        # alert when queue reaches cutoff
        if not minutes:
            print(f'Could not extract queue')
        else:
            print(f'Current queue: {minutes}')
            if minutes <= alertat:
                counter += 1
                if counter >= occurences:
                    print(f'Your queue is about to pop!')
                    send_alert(alarm=alarm, email=emailalert, playsong=playsong)
                    emailalert = False

        # stop after 15 alarms
        if counter > 15:
            break

queue_alert(alarm=True, emailalert=False, playsong=False)