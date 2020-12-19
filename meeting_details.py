def get_meeting_id(meeting_link):
    """get the meeting id from meeting link
    :parm meeting_link: the link for a meeting
    :type meeting_link: str
    :return: the meeting id
    :rtype: str 
    """

    meeting_id = meeting_link.split('/')[-1]
    meeting_id = meeting_id.split('?pwd=')[0]
    return meeting_id


def get_meeting_passcode(meeting_link):
    """get the meeting passcode from meeting link
    :parm meeting_link: the link for a meeting
    :type meeting_link: str
    :return: the meeting passcode
    :rtype: str 
    """

    if '?pwd=' not in meeting_link:
        return None

    url_parts = meeting_link.split('/')
    meeting_passcode = url_parts[-1].split('?pwd=')[1]
    return meeting_passcode
