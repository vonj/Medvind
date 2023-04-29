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
import re


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

def extract_date(txt, sampletime):
    global found_today
    year = sampletime.year

    day, month = [int(x) for x in txt.split()[0].strip().split('/') ]

    if month == sampletime.month and day == sampletime.day:
        found_today = True

    if not found_today and month > sampletime.month:
        year -= 1

    if found_today and month < sampletime.month:
        year += 1

    date = datetime(year, month, day)

    return date

def match_clock(s):
    if len(s) < 5:
        return False
    if s[2] != ':':
        return False
    try:
        hour   = int(s[0:2])
        minute = int(s[3:5])
    except ValueError:
        return False
    except IndexError:
        return False

    return 60 * hour + minute


def extract_working_hours(s):
    lowest = 60 * 23 + 59
    highest = 0

    clocks = []

    while len(s):
        clock = match_clock(s)
        if False != clock:
            if lowest > clock:
                lowest = clock
            if highest < clock:
                highest = clock
        s = s[1:]

    if (highest - lowest) < 30:
        return (None, None)

    sta_hour    = lowest // 60
    sta_minute  = lowest % 60
    end_hour    = highest // 60
    end_minute  = highest % 60

    return (
        "{0:0>2}".format(sta_hour) + ":" + "{0:0>2}".format(sta_minute),
        "{0:0>2}".format(end_hour) + ":" + "{0:0>2}".format(end_minute)
    )


def convert_to_ical(days):

    cal = Calendar()

    for dayk in days:
        start = days[dayk]['start']
        end   = days[dayk]['end']
        if start != '23:00' and end != '00:00':
            event = Event()
            event.add('summary', 'Krn')

            splits = dayk.split('-')
            year  = int(splits[0])
            month = int(splits[1])
            day   = int(splits[2])

            start_hour   = int(start.split(':')[0])
            start_minute = int(start.split(':')[1])
            event.add('dtstart', datetime(year, month, day, start_hour, start_minute, 0, 0))

            end_hour     = int(end.split(':')[0])
            end_minute   = int(end.split(':')[1])
            event.add('dtend',   datetime(year, month, day, end_hour, end_minute, 0, 0))

            cal.add_component(event)
            print(event)
        else:
            print('skip empty day')

    with open('jobschedule.ics', 'wb') as f:
        f.write(cal.to_ical())


def download_html():
    option = webdriver.FirefoxOptions()
    option.binary_location = firefox
    driverService = Service('/path/to/geckodriver')
    drv = webdriver.Firefox(service=driverService, options=option)
    drv.set_page_load_timeout(80)

    # ff = webdriver.Firefox(executable_path=r'your\path\geckodriver.exe')
    # "Medvind-ux-field-Edit-1014-inputEl"
    # "Medvind-ux-field-Edit-1015-inputEl"
    drv.get('https://medvind-mobil.kronansapotek.se/MvWeb/')
    time.sleep(3)

    original_window = drv.current_window_handle
    assert len(drv.window_handles) == 1

    # To catch <input type="text" id="passwd" />
    login = drv.find_element(By.ID, "Medvind-ux-field-Edit-1014-inputEl")
    # To catch <input type="text" name="passwd" />
    password = drv.find_element(By.ID, "Medvind-ux-field-Edit-1015-inputEl")

    login.send_keys(medvind_username)
    time.sleep(1)
    password.send_keys(medvind_password)
    time.sleep(5)

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

    time.sleep(10)

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

    return fullpath


def parse_calendar(htmlfile, logfile):
    f = open(htmlfile, 'r')

    t = os.path.getmtime(htmlfile)
    sampletime = datetime.fromtimestamp(t)

    filecontent = f.read()
    soup = BeautifulSoup(filecontent, 'html.parser')

    stored_file = 'latest.json'
    with open(stored_file, 'r') as f:
        previous = json.load(f)

    samples = {
        'days': {}
    }

    logcontents = pathlib.Path(logfile).read_text()
    changes = []

    for content in soup.find_all('div', {'class': 'mv-daycell'}):
        day = extract_date(content.text, sampletime).strftime('%Y-%m-%d')

        times = ''
        for sub in content.find_all('div', {'style': re.compile('.*FF00FF')}):
            times += sub.text

        (start, end) = extract_working_hours(times)
        if None == start:
            (start, end) = extract_working_hours(content.text)

        if day in previous['days']:
            that_day = previous['days'][day]
            if start != that_day['start'] or end != that_day['end']:
                changes.append(
                    f"{day} Gammal: {that_day['start']}-{that_day['end']}  Ny: {start}-{end}"
                )

        samples['days'][day] = {
            'last_change': sampletime.isoformat(),
            'start': start,
            'end':   end
        }

    with open(stored_file, 'w') as f:
        f.write(json.dumps(samples, indent = 2))


    return samples, changes, logcontents


def run_all():
    fullpath = download_html()

    logfile = 'medvind_log.txt'

    samples, changes, logcontents = parse_calendar(fullpath, logfile)

    convert_to_ical(samples['days'])

    if len(changes) > 0:
        with open(logfile, 'w') as f:
            f.write('\n'.join(changes))
            f.write("\n")
            # f.write(f"   (Ändringar upptäckta: {str(now)[:16]})\n")
            f.write("\n")
            f.write("\n")
            f.write(logcontents)


        sendreport(pathlib.Path(logfile).read_text())

    # print('Previous run at: ', datetime.fromisoformat(previous['sample_time']))

    return 0


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'test':
        # samples, changes, logcontents = parse_calendar('Medvind.html', 'testlog.txt')
        # print(samples)
        # print('---------')
        # print(samples['days']['2023-04-23'])
        txt = '23/4 sönLedig10:45-11:00Öp11:00-15:00Ar    Egenvård15:00-15:20St    Kassaräkning'
        print(txt)
        # t = os.path.getmtime('latest.json')
        t = 1682461710.568336
        sampletime = datetime.fromtimestamp(t)
        day = extract_date(txt, sampletime).strftime('%Y-%m-%d')
        (start, end) = extract_working_hours(txt)
        print('day:', day)
        print('hours:', start, end)

        parse_calendar('Medvind.html', 'testlog.txt')
        sys.exit(0)


    sys.exit(run_all())