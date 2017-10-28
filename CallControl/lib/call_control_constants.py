#-------------------------------------------------------------------------------
# Name:        call_control_constants
# Purpose:     Defines all constants related to call control: Softphone
#              and media player states, operation types, status codes,
#              and audio locations.
#
# Author:      rsalsburry/rashforth
#
# Created:     19/07/2012
# Copyright:   (c) Plantronics 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

# Constants representing specific iUI_CallControl methods
# Used to perform validation of operation types permitted for each method
#
# CallControl:
PLACE_CALL = 1
ANSWER_CALL = 2
DECLINE_CALL = 3
END_CALL = 4
SET_CALL_MUTING = 5
HOLD_RESUME_CALL = 6
TRANSFER_CALL_AUDIO = 7
TRANSFER_CALL = 8
ADJUST_CALL_VOLUME = 9
CONFERENCE_CALL = 10
SET_AUTO_ANSWER = 11
#
# MediaControl:
PLAY_MEDIA = 101
PAUSE_MEDIA = 102
STOP_MEDIA = 103
REPEAT_MEDIA = 104
MUTE_MEDIA = 105
ADJUST_VOLUME_MEDIA = 106


# operation_type - this integer type is used to describe the type of operation performed on the softphone?s UI.
# The operation_type constants are as follows:
#
# CallControl:
OT_AUTO_ANSWER_ON = 10 # turn on the auto-answer feature of the softphone (Skype only feature)
OT_AUTO_ANSWER_OFF = 11 # turn off the auto-answer feature of the softphone (Skype only feature)
OT_CONFERENCE_DRAGDROP = 20 # conference call, conference by drag and drop in the softphone UI
OT_CONFERENCE_TWO_CALLS = 30 # conference call, conference together two separate calls in the softphone UI
OT_CONFERENCE_TWO_CONTACTS = 40 # conference call, call two different contacts at once in the softphone UI
OT_CONTACT_NAME = 50 # the name of the softphone contact, for example, Pam White
OT_DRAGDROP_TO_CONVERSATION_WND = 60 # drag-and-drop softphone contact to the softphone?s conversation window
OT_DRAGDROP_TO_PHONE_MENU = 70 # drag-and-drop softphone contact to the softphone?s phone menu
OT_EXTENSION_NUMBER_DIALPAD = 80 # the extension number dialed from the softphone?s dial pad
OT_EXTENSION_NUMBER_MAIN_WND = 90 # the extension number dialed from the softphone?s main window, for example, 2001
OT_HOLD = 100 # put call on-hold from the softphone UI
OT_MAIN_WND = 110 # the softphone?s main window
OT_MUTE = 120 # mute a call from the softphone UI
OT_PHONE_NUMBER_DIALPAD = 130 # the phone number dialed from the softphone?s dial pad
OT_PHONE_NUMBER_MAIN_WND = 140 # the phone number dialed from the softphone?s main window
OT_PICKUP = 150 # the pick-up feature of the softphone (CIPC only feature)
OT_RESUME = 160 # resume held call from the softphone UI
OT_START_CONFERENCE_CALL = 170 # the conference call made to two different softphone contacts (?call group? feature in Skype, ?start conference call? feature in Lync)
OT_TOAST = 180 # the softphone?s toast, or pop-up window
OT_TRANSFER_AUDIO_NEXT = 190 # transfer audio from the Plantronics? headset to PC Speakers in the softphone UI
OT_TRANSFER_AUDIO_PREV = 200 # transfer audio from the PC Speakers to Plantronics? headset in the softphone UI
OT_TRANSFER_CALL = 210 # transfer call to another contact in the softphone UI
OT_UNMUTE = 220 # unmute a call from the softphone UI
OT_VOLUME_DOWN_SPEAKER = 230 # during a call decrease volume of the speaker in the softphone UI
OT_VOLUME_UP_SPEAKER = 240 # during a call increase volume of the speaker in the softphone UI
OT_CHECK_ISDIALING = 250 # Is a call being dialed on the softphone device?
OT_CHECK_ISRINGING = 260 # Is a call ringing on the softphone device?
OT_CHECK_HASCALLS = 270 # Does the softphone have any connected calls?
OT_CHECK_ISCONNECTED = 280 # Is a call connected for the specified contact?
OT_CHECK_ISMUTED = 290 # Is the call for the specific contact muted?
OT_CHECK_ISHELD = 300 # Is the call for the specific contact on hold?
OT_CHECK_ISINCONFERENCE = 310 # Is the call for the specific contact in a conference?
OT_CHECK_USERSTATE = 320 # Return the current state of the softphone user

# User state constants
#
USERSTATE_UNDEFINED = 0 # If the user state matches none of the specifically identified states
USERSTATE_INCALL = 1 # The user is in a call
USERSTATE_INCONFERENCECALL = 2 # The user is in a conference call
USERSTATE_AVAILABLE = 3 # The user is available
USERSTATE_BUSY = 4 # The user is busy
USERSTATE_AWAY = 5 # The user is away

# Dictionaries of strings corresponding to user state constants
# and the reverse
#
userstate_strings = {
    USERSTATE_UNDEFINED : 'Undefined',
    USERSTATE_INCALL : 'In a call',
    USERSTATE_INCONFERENCECALL : 'In a conference call',
    USERSTATE_AVAILABLE : 'Available',
    USERSTATE_BUSY : 'Busy',
    USERSTATE_AWAY : 'Away',
    }

userstate_constants = {
    'Undefined' : USERSTATE_UNDEFINED,
    'In a call' : USERSTATE_INCALL,
    'In a conference call' : USERSTATE_INCONFERENCECALL,
    'Available' : USERSTATE_AVAILABLE,
    'Busy' : USERSTATE_BUSY,
    'Away' : USERSTATE_AWAY,
    }

#
# MediaControl:
OT_MEDIA_MUTE = 5
OT_MEDIA_UNMUTE = 15
OT_MEDIA_REPEAT_ON = 25
OT_MEDIA_REPEAT_OFF = 35
OT_MEDIA_CHECK_IS_PLAYING = 45
OT_MEDIA_CHECK_IS_PAUSED = 55
OT_MEDIA_CHECK_IS_STOPPED = 65
OT_MEDIA_CHECK_IS_MUTED = 75
OT_MEDIA_CHECK_IS_REPEATMODEON = 85
OT_MEDIA_VOLUME_UP = 95
OT_MEDIA_VOLUME_DOWN = 105

# return_code - this integer type is used to describe the output parameter of the function call as a success of failure.
# The return_code constants are as follows:
STATUS_SUCCESS = 0
STATUS_BOOL_TRUE = 0    # Note that this value is also used as 'SUCCESS'
STATUS_BOOL_FALSE = 1   # Consistent with standard int value for False
STATUS_FAILURE = 2
STATUS_FEATURE_NOT_SUPPORTED = 3
STATUS_FEATURE_NOT_IMPLEMENTED = 4
STATUS_INVALID_OPERATION_TYPE = 5

sp_platforms= ['windows', 'linux', 'osx']

SP_PORTS_BASE = 5003
SQUISHPORTS_BASE = 4404
CCS_ADMIN_PORT = 7001

#==================================
# The following (long) section specifies all required settings for each supported softphone.
# The sp_dict which appears at the end is the 'dictionary of dictionaries' with the string sp_id
# as the key.
# This section also includes accessor functions for each softphone attribute.

avaya_ip_agent_dict = {
    'name' : 'Avaya IP Agent',      # Actual product name
    'port' : SP_PORTS_BASE + 10,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 10,  # Port number for communication with squishserver
    }

avaya_ip_softphone_dict = {
    'name' : 'Avaya IP Softphone',  # Actual product name
    'port' : SP_PORTS_BASE + 20,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 20,  # Port number for communication with squishserver
    }

avaya_onex_agent_dict = {
    'name' : 'Avaya one-X Agent',   # Actual product name
    'port' : SP_PORTS_BASE + 30,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 30,  # Port number for communication with squishserver
    }

avaya_onex_communicator_dict = {
    'name' : 'Avaya one-X Communicator',    # Actual product name
    'port' : SP_PORTS_BASE + 40,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 40,  # Port number for communication with squishserver
    }

cisco_ip_communicator_dict = {
    'name' : 'Cisco IP Commmunicator',    # Actual product name
    'port' : SP_PORTS_BASE + 50,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 50,  # Port number for communication with squishserver
    }

cucilync_dict = {
    'name' : 'CUCILync',            # Actual product name
    'port' : SP_PORTS_BASE + 60,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 60,  # Port number for communication with squishserver
    }

cucimoc_dict = {
    'name' : 'CUCIMOC',             # Actual product name
    'port' : SP_PORTS_BASE + 70,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 70,  # Port number for communication with squishserver
    }

cupc_dict = {
    'name' : 'CUPC',                # Actual product name
    'port' : SP_PORTS_BASE + 80,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 80,  # Port number for communication with squishserver
    }

cisco_webex_connect_dict = {
    'name' : 'Cisco WebEx Connect', # Actual product name
    'port' : SP_PORTS_BASE + 90,    # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 90,  # Port number for communication with squishserver
    }

ibm_sametime_dict = {
    'name' : 'IBM Sametime',        # Actual product name
    'port' : SP_PORTS_BASE + 100,   # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 100,    # Port number for communication with squishserver
    }

microsoft_lync_dict = {
    'name' : 'Microsoft Lync',      # Actual product name
    'port' : SP_PORTS_BASE + 110,   # Port for communication with CallControl server
    'aut_path' : 'C:\\Program Files\\Microsoft Lync\\communicator.exe',  # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 110,    # Port number for communication with squishserver
    }

microsoft_lync_2013_dict = {
    'name' : 'Microsoft Lync 2013',      # Actual product name
    'port' : SP_PORTS_BASE + 120,   # Port for communication with CallControl server
    'aut_path' : 'C:\\Program Files\\Microsoft Office\\Office15\\lync.exe',  # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 120,    # Port number for communication with squishserver
    }

microsoft_oc_r2_dict = {
    'name' : 'Microsoft OC R2',     # Actual product name
    'port' : SP_PORTS_BASE + 130,   # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 130,    # Port number for communication with squishserver
    }

nec_sp350_dict = {
    'name' : 'NEC SP350',   # Actual product name
    'port' : SP_PORTS_BASE + 140,   # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 140,    # Port number for communication with squishserver
    }

shoretel_call_manager_dict = {
    'name' : 'ShoreTel Call Manager',   # Actual product name
    'port' : SP_PORTS_BASE + 150,   # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 150,    # Port number for communication with squishserver
    }

shoretel_communicator_dict = {
    'name' : 'ShoreTel Communicator',   # Actual product name
    'port' : SP_PORTS_BASE + 160,   # Port for communication with CallControl server
    'aut_path' : 'TBD',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 160,    # Port number for communication with squishserver
    }

skype_dict = {
    'name' : 'Skype',               # Actual product name
    'port' : SP_PORTS_BASE + 170,   # Port for communication with CallControl server
    'aut_path' : 'C:\\Program Files\\Skype\\Phone\\Skype.exe',             # Path to softphone executable
    'squishport' : SQUISHPORTS_BASE + 170,    # Port number for communication with squishserver
    }


'''
This is the dictionary of all softphone-specific dictionaries
'''
sp_dict = {
            'avaya_ip_agent' : avaya_ip_agent_dict,
            'avaya_ip_softphone' : avaya_ip_softphone_dict,
            'avaya_onex_agent' : avaya_onex_agent_dict,
            'avaya_onex_communicator' : avaya_onex_communicator_dict,
            'cisco_ip_communicator' : cisco_ip_communicator_dict,
            'cucilync' : cucilync_dict,
            'cucimoc' : cucimoc_dict,
            'cupc' : cupc_dict,
            'cisco_webex_connect' : cisco_webex_connect_dict,
            'ibm_sametime' : ibm_sametime_dict,
            'microsoft_lync' : microsoft_lync_dict,
            'microsoft_lync_2013' : microsoft_lync_2013_dict,
            'microsoft_oc_r2' : microsoft_oc_r2_dict,
            'nec_sp350' : nec_sp350_dict,
            'shoretel_call_manager' : shoretel_call_manager_dict,
            'shoretel_communicator' : shoretel_communicator_dict,
            'skype' : skype_dict,
            }


'''
This may seem perverse, but the dictionary below is provided so that softphone
names can be used to reference softphone attributes as easily (or almost as
easily) as softphone id strings. Providing this facility removes the need to
modify upper layers of the framework to pass down the softphone id string rather
than the product name string.
'''
sp_name_to_id_dict = {
            'Avaya IP Agent' : 'avaya_ip_agent',
            'Avaya IP Softphone' : 'avaya_ip_softphone',
            'Avaya one-X Agent' : 'avaya_onex_agent',
            'Avaya one-X Communicator' : 'avaya_onex_communicator',
            'Cisco IP Commmunicator' : 'cisco_ip_communicator',
            'CUCILync' : 'cucilync',
            'CUCIMOC' : 'cucimoc',
            'CUPC' : 'cupc',
            'Cisco WebEx Connect' : 'cisco_webex_connect_dict',
            'IBM Sametime' : 'ibm_sametime',
            'Microsoft Lync' : 'microsoft_lync',
            'Microsoft Lync 2013' : 'microsoft_lync_2013',
            'Microsoft OC R2' : 'microsoft_oc_r2',
            'NEC SP350' : 'nec_sp350',
            'ShoreTel Call Manager' : 'shoretel_call_manager',
            'ShoreTel Communicator' : 'shoretel_communicator',
            'Skype' : 'skype',
            }


def get_sp_id(sp_name):
    if sp_name not in sp_name_to_id_dict:
        sp_id = None
    else:
        sp_id = sp_name_to_id_dict[sp_name]
    return sp_id

def get_sp_name(sp_id):
    if sp_id not in sp_dict:
        pdict = None
    else:
        pdict = sp_dict[sp_id]
    if 'name' in pdict:
        return pdict['name']
    else:
        print 'Error, key \'name\' not in softphone dictionary'
        return None

def get_sp_port(sp_id):
    if sp_id not in sp_dict:
        pdict = None
    else:
        pdict = sp_dict[sp_id]
    if 'port' in pdict:
        return pdict['port']
    else:
        print 'Error, key \'port\' not in softphone dictionary'
        return None

def get_sp_aut_path(sp_id):
    if sp_id not in sp_dict:
        pdict = None
    else:
        pdict = sp_dict[sp_id]
    if 'aut_path' in pdict:
        return pdict['aut_path']
    else:
        print 'Error, key \'aut_path\' not in softphone dictionary'
        return None

def get_sp_squishport(sp_id):
    if sp_id not in sp_dict:
        pdict = None
    else:
        pdict = sp_dict[sp_id]
    if 'squishport' in pdict:
        return pdict['squishport']
    else:
        print 'Error, key \'squishport\' not in softphone dictionary'
        return None

# End softphone dict section
#==================================


# This dictionary tracks the operation types which are valid for a given
# function call. Specific validation of supported operations per call
# for individual soft phones will be performed within the corresponding
# function calls
valid_func_op_dict = {
        PLACE_CALL:
                    [OT_CONTACT_NAME,
                    OT_PHONE_NUMBER_MAIN_WND,
                    OT_PHONE_NUMBER_DIALPAD,
                    OT_EXTENSION_NUMBER_MAIN_WND,
                    OT_EXTENSION_NUMBER_DIALPAD,
                    OT_DRAGDROP_TO_PHONE_MENU,
                    OT_DRAGDROP_TO_CONVERSATION_WND,
                    OT_START_CONFERENCE_CALL],
        ANSWER_CALL:
                    [OT_TOAST,
                    OT_MAIN_WND,
                    OT_PICKUP],
        DECLINE_CALL:
                    [OT_TOAST,
                    OT_MAIN_WND],
        END_CALL:
                    [OT_TOAST,
                    OT_MAIN_WND],
        SET_CALL_MUTING:
                    [OT_MUTE,
                    OT_UNMUTE],
        HOLD_RESUME_CALL:
                    [OT_HOLD,
                    OT_RESUME],
        TRANSFER_CALL_AUDIO:
                    [OT_TRANSFER_AUDIO_NEXT,
                    OT_TRANSFER_AUDIO_PREV],
        TRANSFER_CALL:
                    [OT_TRANSFER_CALL],
        ADJUST_CALL_VOLUME:
                    [OT_VOLUME_UP_SPEAKER,
                    OT_VOLUME_DOWN_SPEAKER],
        CONFERENCE_CALL:
                    [OT_CONFERENCE_TWO_CALLS,
                    OT_CONFERENCE_TWO_CONTACTS],
        SET_AUTO_ANSWER:
                    [OT_AUTO_ANSWER_ON,
                    OT_AUTO_ANSWER_OFF],
        }


#==================================
# The following section specifies all required settings for each supported media player.
# The mp_dict which appears at the end is the 'dictionary of dictionaries' with the string sp_id
# as the key.

itunes_dict = {
    'name' : 'iTunes',    # Actual product name
    'port' : SP_PORTS_BASE + 5,         # Port for communication with MediaControl server
    'aut_path' : 'C:\\Program Files\\iTunes\\iTunes.exe',  # Path to media player executable
    'squishport' : SQUISHPORTS_BASE + 5,    # Port number for communication with squishserver
    }

windows_media_player_dict = {
    'name' : 'Windows Media Player',    # Actual product name
    'port' : SP_PORTS_BASE + 15,         # Port for communication with MediaControl server
    'aut_path' : 'C:\\Program Files\\Windows Media Player\\wmplayer.exe',  # Path to media player executable
    'squishport' : SQUISHPORTS_BASE + 15,    # Port number for communication with squishserver
    }

winamp_dict = {
    'name' : 'iTunes',    # Actual product name
    'port' : SP_PORTS_BASE + 25,         # Port for communication with MediaControl server
    'aut_path' : 'C:\\Program Files\\Winamp\\winamp.exe',  # Path to media player executable
    'squishport' : SQUISHPORTS_BASE + 25,    # Port number for communication with squishserver
    }

'''
This is the dictionary of all mediaplayer-specific dictionaries
'''
mp_dict = {
            'itunes' : itunes_dict,
            'windows_media_player' : windows_media_player_dict,
            'winamp' : winamp_dict,
            }

def get_mp_id(mp_name):
    if mp_name not in mp_name_to_id_dict:
        mp_id = None
    else:
        mp_id = mp_name_to_id_dict[mp_name]
    return mp_id

def get_mp_name(mp_id):
    if mp_id not in mp_dict:
        pdict = None
    else:
        pdict = mp_dict[mp_id]
    if 'name' in pdict:
        return pdict['name']
    else:
        print 'Error, key \'name\' not in softphone dictionary'
        return None

def get_mp_port(mp_id):
    if mp_id not in mp_dict:
        pdict = None
    else:
        pdict = mp_dict[mp_id]
    if 'port' in pdict:
        return pdict['port']
    else:
        print 'Error, key \'port\' not in softphone dictionary'
        return None

def get_mp_aut_path(mp_id):
    if mp_id not in mp_dict:
        pdict = None
    else:
        pdict = mp_dict[mp_id]
    if 'aut_path' in pdict:
        return pdict['aut_path']
    else:
        print 'Error, key \'aut_path\' not in softphone dictionary'
        return None

def get_mp_squishport(mp_id):
    if mp_id not in mp_dict:
        pdict = None
    else:
        pdict = mp_dict[mp_id]
    if 'squishport' in pdict:
        return pdict['squishport']
    else:
        print 'Error, key \'squishport\' not in softphone dictionary'
        return None

