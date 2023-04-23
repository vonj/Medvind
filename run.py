import json
import os
import sys
import time
import pathlib
import shutil
from subprocess import PIPE, Popen
from bs4 import BeautifulSoup
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
import selenium
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import pyautogui
import smtplib
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import ssl
from icalendar import Calendar, Event


config = configparser.ConfigParser()
config.read('medvind.ini')

username = config.get('email', 'username')
password = config.get('email', 'password')
sender   = config.get('email', 'sender')
to       = config.get('email', 'to')
server   = config.get('email', 'server')

medvind_username = config.get('medvind', 'username')
medvind_password = config.get('medvind', 'password')

try:
    firefox = config.get('firefox', 'exepath')
except configparser.NoOptionError:
    firefox = '/usr/bin/firefox'


def sendreport(report):

    port = 588
    context = ssl.create_default_context()

    toaddrs = [sender, to]

    msg = MIMEMultipart('alternative')
    msg['From'] = formataddr((str(Header('Medvind', 'utf-8')), sender))
    msg['To'] = to
    html = '<html><body><pre>' + report + '</pre></body></html>'
    msg.attach(MIMEText(html, 'html'))
    print(msg)

    with smtplib.SMTP_SSL(server, port, context=context) as s:
        s.login(username, password)
        s.sendmail(sender, toaddrs, msg.as_string())





found_today = False

def extract_date(txt):
    global found_today
    year = now.year

    day, month = [int(x) for x in txt.split()[0].strip().split('/') ]

    if month == now.month and day == now.day:
        found_today = True
    
    if not found_today and month > now.month:
        year -= 1
    
    if found_today and month < now.month:
        year += 1

    date = datetime(year, month, day)    

    return date

def match_clock(s):
    try:        
        if s[2] != ':':
            return False
        if s[5] != '-':
            return False
        if s[8] != ':':
            return False
    except IndexError:
        return False

    try:
        start_hour   = int(s[0:2])
        start_minute = int(s[3:5])
        end_hour     = int(s[6:8])
        end_minute   = int(s[9:11])
        
        start_minute += start_hour * 60
        end_minute   += end_hour   * 60
        return [start_minute, end_minute]
    except ValueError as e:
        print('ValueError', e)
        pass

    return False

def extract_hours(txt):

    sta_minute = 23 * 60
    end_minute = 0

    i = 0
    while len(txt[i:]) > 0:
        result = match_clock(txt[i:])
        if False == result:
            pass
        else:
            if result[0] < sta_minute:
                sta_minute = result[0]
            if result[1] > end_minute:
                end_minute = result[1]
        i += 1
    
    sta_hour    = sta_minute // 60
    sta_minute %= 60
    end_hour    = end_minute // 60
    end_minute %= 60

    return (
        "{0:0>2}".format(sta_hour) + ":" + "{0:0>2}".format(sta_minute),
        "{0:0>2}".format(end_hour) + ":" + "{0:0>2}".format(sta_minute)
    )


def convert_to_ical(days):

    cal = Calendar()

    for dayk in days:
        start = days[dayk]['start']
        end   = days[dayk]['end']
        if start != '23:00' and end != '00:00':
            event = Event()
            event.add('summary', 'jobb')

            splits = day.split('-')
            year  = int(splits[0])
            month = int(splits[1])
            day   = int(splits[2])

            start_hour   = start.split(':')[0]
            start_minute = start.split(':')[1]
            event.add('dtstart', year, month, day, start_hour, start_minute, 0, 0)

            end_hour     = end.split(':')[0]
            end_minute   = end.split(':')[1]
            event.add('dtend',   year, month, day, end_hour, end_minute, 0, 0)

            cal.add_component(event)

    with open('jobschedule.ics', 'wb') as f:
        f.write(cal.to_ical())


if __name__ == '__main__':
    option = webdriver.FirefoxOptions()
    option.binary_location = firefox
    driverService = Service('/path/to/geckodriver')
    drv = webdriver.Firefox(service=driverService, options=option)
    drv.set_page_load_timeout(60)

    # ff = webdriver.Firefox(executable_path=r'your\path\geckodriver.exe')
    # "Medvind-ux-field-Edit-1014-inputEl"
    # "Medvind-ux-field-Edit-1015-inputEl"
    drv.get('https://medvind-mobil.kronansapotek.se/MvWeb/')
    time.sleep(1)

    original_window = drv.current_window_handle
    assert len(drv.window_handles) == 1

    # To catch <input type="text" id="passwd" />
    login = drv.find_element(By.ID, "Medvind-ux-field-Edit-1014-inputEl")
    # To catch <input type="text" name="passwd" />
    password = drv.find_element(By.ID, "Medvind-ux-field-Edit-1015-inputEl")

    login.send_keys(medvind_username)
    time.sleep(1)
    password.send_keys(medvind_password)
    time.sleep(1)


    try:
        # drv.find_element(By.ID, 'button-1017-btnWrap').click()
        # drv.find_element(By.ID, 'button-1017-btnEl').click()
        # drv.find_element(By.ID, 'button-1017-btnIconEl').click()
        drv.find_element(By.ID, 'button-1017-btnInnerEl').click()
        pass
    except selenium.common.exceptions.NoSuchElementException:
        pass

    downloads = os.path.join(pathlib.Path.home(), 'Downloads')
    
    filename = 'Medvind.html'
    fullpath = os.path.join(downloads, filename)
    try:
        os.remove(fullpath)
    except:
        pass

    shutil.rmtree(os.path.join(downloads, 'Medvind_files'), ignore_errors=True)

    time.sleep(30)

    pyautogui.hotkey('ctrl', 's')
    time.sleep(10)

    #with pyautogui.hold('command'):
    #    pyautogui.press('s')
    
    # pyautogui.keyDown('command')
    # pyautogui.keyUp('command')
    # pyautogui.press('tab', presses = 11, interval = 0.2)
    pyautogui.press('tab', presses = 2, interval = 3)
    pyautogui.press('enter')
    # pyautogui.typewrite(SEQUENCE + '.html')

    time.sleep(20)
    drv.quit()
    time.sleep(2)
    os.system("killall firefox")

    #drv.close()

    f = open(fullpath, 'r')
    t = os.path.getmtime(fullpath)
    now = datetime.fromtimestamp(t)

    filecontent = f.read()
    soup = BeautifulSoup(filecontent, 'html.parser')

    stored_file = 'latest.json'
    with open(stored_file, 'r') as f:
        previous = json.load(f)

    samples = {
        'days': {}
    }

    logfile = 'medvind_log.txt'
    logcontents = pathlib.Path(logfile).read_text()
    changes = []
    
    for content in soup.find_all('div', {'class': 'mv-daycell'}):
        txt = content.text
        day = extract_date(txt).strftime('%Y-%m-%d')
        (start, end) = extract_hours(txt)

        if day in previous['days']:
            that_day = previous['days'][day]
            if start != that_day['start'] or end != that_day['end']:
                changes.append(
                    f"{day} Gammal: {that_day['start']}-{that_day['end']}  Ny: {start}-{end}"
                )

        samples['days'][day] = {
            'last_change': now.isoformat(),
            'start': start,
            'end':   end
        }

    convert_to_ical(samples['days'])
    
    if len(changes) > 0:
        with open(logfile, 'w') as f:
            f.write('\n'.join(changes))
            f.write("\n")
            f.write(f"   (Ändringar upptäckta: {str(now)[:16]})\n")
            f.write("\n")
            f.write("\n")
            f.write(logcontents)


        sendreport(pathlib.Path(logfile).read_text())

    # print('Previous run at: ', datetime.fromisoformat(previous['sample_time']))

    with open(stored_file, 'w') as f:
        f.write(json.dumps(samples, indent = 2))

