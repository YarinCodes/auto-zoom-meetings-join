import time, sys, os
from datetime import datetime
import meeting_details
try:
    import pyautogui, openpyxl
except ImportError as e:
    print(f'error, install the requirements:\n{e}') 
    sys.exit(e)
    

XL_FILE_PATH = r"Paste xl file path HERE"
ZOOM_PATH = r"Paste Zoom Shourtcut Path Here"

wb = openpyxl.load_workbook(XL_FILE_PATH)
sheet = wb.active


def process_meeting_details(meeting_row):
    """extract the meeting information from al file
    :parm meeting_row: the meeting's row to get information about
    :type meeting_row: int
    """
    meeting_id = str(sheet.cell(row=meeting_row, column=3).value)
    meeting_passcode = str(sheet.cell(row=meeting_row, column=4).value)
    meeting_link = sheet.cell(row=meeting_row, column=5).value

    if meeting_link is not None: # in case only the link was given
        meeting_id = meeting_details.get_meeting_id(meeting_link)
        meeting_passcode = meeting_details.get_meeting_passcode(meeting_link)

    join_meeting(meeting_id, meeting_passcode)


def join_meeting(meeting_id, meeting_passcode):
    """join the meeting by pressing images on screen
    :parm meeting_id: the meeting's id
    :parm meeting_passcode: the meeting's passcode
    :type meeting_id: str
    :type meeting_passcode: str
    """

    os.startfile(ZOOM_PATH)

    # locate the '+' icon to join a meeting
    button = pyautogui.locateOnScreen('images/join_a_meeting.png') 
    while button is None:
        button = pyautogui.locateOnScreen('images/join_a_meeting.png')
        
    pyautogui.click(button)


    # loctae 'enter meeting link' text box
    button = pyautogui.locateOnScreen('images/enter_link.png')
    while button is None:
        button = pyautogui.locateOnScreen('images/enter_link.png')

    pyautogui.click(button)
    pyautogui.write(meeting_id) # enter meeting id

    # join without audio
    button = pyautogui.locateCenterOnScreen('images/check_box.png')
    while button is None:
        button = pyautogui.locateCenterOnScreen('images/check_box.png')

    pyautogui.click(button)

    # join without video
    button = pyautogui.locateCenterOnScreen('images/check_box.png')
    while button is None:
        button = pyautogui.locateCenterOnScreen('images/check_box.png')

    pyautogui.click(button)

    # join button in 'enter meeting link' screen
    button = pyautogui.locateOnScreen('images/join_btn.png')
    while button is None:
        button = pyautogui.locateOnScreen('images/join_btn.png')

    pyautogui.click(button)

    # type meeting's passcode if it has one
    if meeting_passcode is not None:
        button = pyautogui.locateOnScreen('images/enter_passcode.png')
        while button is None:
            button = pyautogui.locateOnScreen('images/enter_passcode.png')

        pyautogui.click(button)
        pyautogui.write(meeting_passcode)

        button = pyautogui.locateOnScreen('images/join_btn.png')
        while button is None:
            button = pyautogui.locateOnScreen('images/join_btn.png')
        
        pyautogui.click(button)

    
def main():

    meetings_start_time = {}
    #  save for every meeting its starting time as a key
    #  and number of row as value
    for i in range(2, sheet.max_row + 1):
        meetings_start_time[str(sheet.cell(row=i, column=1).value)[:-3]] = i


    while True:
        current_time = datetime.now()
        current_time = current_time.strftime('%H:%M')
        
        if current_time in meetings_start_time:
            meeting_row = meetings_start_time[current_time]
            process_meeting_details(meeting_row)
            time.sleep(5 * 60) # sleep until next meeting

            
if __name__ == '__main__':
    print('app is running...')
    main()