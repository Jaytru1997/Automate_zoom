# Automate_zoom
Automate zoom meeting where meeting id and password is required.


Checks the "timings.xlsx" file to look for meetings that are going to start.
As soon as the current time matches any meeting time it opens the Zoom Desktop application.
Navigates the cursor automatically to various steps to join the meeting.
The meeting ID and passcode are extracted from "timings.xlsx" and entered into the Zoom app automatically.


## Prerequisites


Zoom app must be installed in your system.
Meeting time for the day along with Meeting ID and passcode must be entered manually into the "timings.xlsx"
pyautogui, python, pandas


## Behind the scenes


An infinite loop keeps checking the current time of the system using "datetime.now" function.
The zoom app is opened using "os.startfile()" function as soon as current time matches the time mentioned in "timings.xlsx".
"pyautogui.locateOnScreen()" function locates the image of join button on the screen and returns the position.
"pyautogui.moveTo()" moves the cursor to that location.
"pyautogui.click()" performs a click operation.
The meeting Id and Passcode are entered using the "pyautogui.write()" command.
Once the meeting is successfully signed in the computer audio authorisation is automated.
