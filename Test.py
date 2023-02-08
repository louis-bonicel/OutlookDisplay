import os
import signal
import socket
import sys
import time
from datetime import datetime
import datetime as dt
import win32com.client

begin = dt.datetime.today()
end = begin + dt.timedelta(1)

# Import only the modules for LCD communication
from library.lcd_comm_rev_a import LcdCommRevA, Orientation
from library.log import logger

COM_PORT = "COM4"
REVISION = "A"

stop = False

def getDailyRdvs( begin, end ):
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort('[Start]')
    restriction = "[Start] >= '" + begin.strftime('%d/%m/%Y') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
    calendar = calendar.Restrict(restriction)

    appointments = [app for app in calendar]    

    cal_subject = [app.subject for app in appointments]
    cal_start = [app.start for app in appointments]
    cal_end = [app.end for app in appointments]

    rdvs = []

    for app in appointments:
        rdvs.append([app.subject, dt.datetime.strptime(str(app.start)[:-6], '%Y-%m-%d %H:%M:%S'), dt.datetime.strptime(str(app.end)[:-6], '%Y-%m-%d %H:%M:%S')])

    return rdvs


if __name__ == "__main__":

    def sighandler(signum, frame):
        global stop
        stop = True

    # Set the signal handlers, to send a complete frame to the LCD before exit
    signal.signal(signal.SIGINT, sighandler)
    signal.signal(signal.SIGTERM, sighandler)
    is_posix = os.name == 'posix'
    if is_posix:
        signal.signal(signal.SIGQUIT, sighandler)

    # Build your LcdComm object based on the HW revision
    lcd_comm = None

    prevNumbRDV = 0
    refresh = 0
    retries = 5
    # Display the current time and some progress bars as fast as possible
    while not stop:       
        rdvs = getDailyRdvs(begin, end)
        try:
            lcd_comm = LcdCommRevA(com_port=COM_PORT,
                            display_width=320,
                            display_height=480)
            if len(rdvs) != prevNumbRDV or refresh == 20:
                lcd_comm.Reset()
                refresh = 0
            prevNumbRDV = len(rdvs)
            
            lcd_comm.InitializeComm()
            lcd_comm.SetBrightness(level=15)
            orientation = Orientation.REVERSE_LANDSCAPE
            lcd_comm.SetOrientation(orientation=orientation)
            
            lcd_comm.DisplayText(str(datetime.now().strftime('%H:%M')), 190, 2)

            line = 50
            for r in rdvs:
                if r[1].strftime('%H') == r[2].strftime('%H'):
                    toPrint = r[1].strftime('%H:%M') + '-' + r[2].strftime('%M') + ' ' + r[0]
                else:    
                    toPrint = r[1].strftime('%H:%M') + '-' + r[2].strftime('%H:%M') + ' ' + r[0]
                if r[2] < datetime.now():
                    lcd_comm.DisplayText( toPrint, 10, line, font_size=15, background_color=(0xBC, 0xBC, 0xBC))
                    line = line + 20
                elif r[1] < datetime.now():
                    lcd_comm.DisplayText( toPrint, 10, line, background_color=(0x9F, 0xC5, 0xE8))
                    line = line + 30
                else:
                    lcd_comm.DisplayText( toPrint, 10, line)
                    line = line + 30
                
            
            lcd_comm.closeSerial()
            retries = 5
            refresh = refresh + 1
        except:
            print('Could not open COM port')
            retries = retries - 1
            if retries == 0:
                break
        
        time.sleep(5)
        

    # Close serial connection at exit
    if lcd_comm != None:
        lcd_comm.closeSerial()
