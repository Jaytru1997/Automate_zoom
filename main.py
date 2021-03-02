import os
import pandas as pd
import pyautogui
import time
from datetime import datetime


def signIn(meeting_id,password):

    #Opens Zoom Application from the specified location
    os.startfile("C:/Users/user/AppData/Roaming/Zoom/bin/Zoom.exe")
    time.sleep(3)

    #Clicks home button
    homebtn=pyautogui.locateCenterOnScreen("homebtn.png")
    pyautogui.moveTo(homebtn)
    pyautogui.click()
    time.sleep(1)

    #Clicks join button
    joinbtn=pyautogui.locateCenterOnScreen("join.PNG")
    pyautogui.moveTo(joinbtn)
    pyautogui.click()
    time.sleep(1)

    #Types the meeting id
    meetingidbtn=pyautogui.locateCenterOnScreen("meetingid.png")
    pyautogui.moveTo(meetingidbtn)
    pyautogui.write(meeting_id)
    time.sleep(1)

    #Clicks Join confirmation button
    joinconfirm=pyautogui.locateCenterOnScreen("joinconfirm.png")
    pyautogui.moveTo(joinconfirm)
    pyautogui.click()
    time.sleep(1)

    #Enters passcode to join meeting
    passcode=pyautogui.locateCenterOnScreen("meetingPasscode.png")
    pyautogui.moveTo(passcode)
    pyautogui.write(password)
    time.sleep(1)

    #Clicks Join Meeting confirm button
    meetingconfirm=pyautogui.locateCenterOnScreen("joinameeting.PNG")
    pyautogui.moveTo(meetingconfirm)
    pyautogui.click()
    time.sleep(1)

df = pd.read_excel('timings.xlsx', index_col=False)

while True:
    #To get current time
    now = datetime.now().strftime("%H:%M")
    if now in str(df['Timings']):
        
        mylist=df["Timings"]
        mylist=[i.strftime("%H:%M") for i in mylist]
        c= [i for i in range(len(mylist)) if mylist[i]==now]
        row = df.loc[c] 
        meeting_id = str(row.iloc[0,1])  
        password= str(row.iloc[0,2])  
        time.sleep(1)
        signIn(meeting_id, password)
        time.sleep(1)
        print('signed in')
        while True:
        #To allow computer audio
            audioBtn=pyautogui.locateAllOnScreen("joinwithcomputeraudio.png")
            for btn in audioBtn:
                pyautogui.moveTo(btn)
                pyautogui.click()
                time.sleep(1)
                print("audio connected successfully")
            break
