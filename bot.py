try:
    import pyautogui, webbrowser, json, time
    from win10toast import ToastNotifier
    from datetime import datetime
    import meeting_details
except ImportError:
    print('error, install the requirements')


notifier = ToastNotifier()


def join_meeting(meeting_name, meeting_link):
    webbrowser.open(r"C:\Users\EdenG\AppData\Roaming\Zoom\bin\Zoom.exe")
    notifier.show_toast(title=f"{meeting_name} Meeting Is Starting!", msg="Joining Meeting...", duration=10)

    meeting_id = meeting_details.get_meeting_id(meeting_link)
    meeting_passcode = meeting_details.get_meeting_passcode(meeting_link)

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


    button = pyautogui.locateCenterOnScreen('images/check_box.png')
    while button is None:
        button = pyautogui.locateCenterOnScreen('images/check_box.png')

    pyautogui.click(button)

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
        print(meeting_passcode)

        button = pyautogui.locateOnScreen('images/join_btn.png')
        while button is None:
            button = pyautogui.locateOnScreen('images/join_btn.png')
        
        pyautogui.click(button)

    notifier.show_toast(title=f"Joined to {meeting_name} meeting", msg="successfully joined to the meeting!", duration=10)



def main():
    mettings = {}
    with open("meetings.json", "r") as mettings_file:
        mettings = json.load(mettings_file)

    while True:
        if time.strftime("%H:%M") in mettings.keys():
            meeting_name = mettings[time.strftime("%H:%M")]["name"]
            meeting_link = mettings[time.strftime("%H:%M")]["link"]
            join_meeting(meeting_name, meeting_link)
            time.sleep(60 * 5)


if __name__ == '__main__':
    print('app is running...')
    main()
    
