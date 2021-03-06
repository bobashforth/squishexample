#-------------------------------------------------------------------------------
# Name:        CallControlAPI
# Purpose:     This class defines the API for the iUI_CallControl interface,
#              used to drive the CallControl servers representing specific
#              softphones within the Plantronics TDTF.
#              (Target Device Test Framework).
#
# Author:      rashforth
#
# Created:     19/07/2012
# Copyright:   (c) Plantronics 2012
#-------------------------------------------------------------------------------

import call_control_constants as ccc

import socket
import sys
import time

from multiprocessing import connection


class CallControlAPI(object):

    def __init__(self, phone_id, host_name=None):
        self.phone_id = phone_id
        if host_name is None:
            self.host_name = socket.gethostbyname(socket.gethostname())
        else:
            self.host_name = host_name
        print 'host_name is %s' % self.host_name
        self.port = ccc.get_sp_port(phone_id)
        #host_ip = socket.gethostbyname(host_name)
        try:
            print 'Establishing client connection at %s, %d' % (self.host_name, self.port)
            self.conn = connection.Client((self.host_name, self.port))
            print 'Connection established'
        except Exception as e:
            print 'Exception in attempt to connect to %s at %d' % (self.host_name, self.port)
            print e.args


    def Terminate(self):
        self.conn.close()


    def sendMessage(self, argdict):
        status = ccc.STATUS_SUCCESS
        try:
            self.conn.send(argdict)
        except Exception as e:
            print 'Error in sendMessage() send'
            print e.args

        time.sleep(1)

        try:
            statusdict = self.conn.recv()
        except Exception as e:
            print 'Error in sendMessage() receiving status'
            print e.args
            statusdict = None

        if statusdict:
            status_string = statusdict['status']
            if status_string == 'True':
                status = True
            elif status_string == 'False':
                status = False
            else:
                status = int(status_string)
        else:
            status = ccc.STATUS_FAILURE

        return status


    def PlaceCall(self, operation_type, contact):
        argdict = {
            'method' : 'PlaceCall',
            'optype' : str(operation_type),
            'contact' : contact,
            }
        status = self.sendMessage(argdict)
        return status


    def setAutoAnswer(self, operation_type):
        argdict = {
        'method' : 'setAutoAnswer',
        'optype' : str(operation_type),
        }
        status = self.sendMessage(argdict)
        return status


    def AnswerCall(self, operation_type):
        argdict = {
        'method' : 'AnswerCall',
        'optype' : str(operation_type),
        }
        status = self.sendMessage(argdict)
        return status


    def DeclineCall(self, operation_type):
        argdict = {
            'method' : 'DeclineCall',
            'optype' : str(operation_type),
            }
        status = self.sendMessage(argdict)
        return status


    def EndCall(self, operation_type):
        argdict = {
            'method' : 'EndCall',
            'optype' : str(operation_type),
            }
        status = self.sendMessage(argdict)
        return status


    def SetCallMuting(self, operation_type):
        argdict = {
            'method' : 'SetCallMuting',
            'optype' : str(operation_type),
            }
        status = self.sendMessage(argdict)
        return status


    def HoldResumeCall(self, operation_type):
        argdict = {
            'method' : 'HoldResumeCall',
            'optype' : str(operation_type),
            }
        status = self.sendMessage(argdict)
        return status


    def TransferCallAudio(self, operation_type):
        argdict = {
            'method' : 'TransferCallAudio',
            'optype' : str(operation_type),
            }
        status = self.sendMessage(argdict)
        return status


    def TransferCall(self, operation_type, newcontact, currentcontact=''):
        argdict = {
            'method' : 'TransferCall',
            'optype' : str(operation_type),
            'contact_two' : newcontact,
            'contact_one' : currentcontact
            }
        status = self.sendMessage(argdict)
        return status


    def AdjustCallVolume(self, operation_type):
        argdict = {
            'method' : 'AdjustCallVolume',
            'optype' : str(operation_type),
            }
        status = self.sendMessage(argdict)
        return status


    # Functionality not clearly defined, TBD
    def ConferenceCall(self, operation_type, contact_one, contact_two):
        argdict = {
            'method' : 'ConferenceCall',
            'optype' : str(operation_type),
            'contact_one' : contact_one,
            'contact_two' : contact_two,
            }
        status = self.sendMessage(argdict)
        return status


    def stateCheck(self, operation_type, contact=''):
        argdict = {
            'method' : 'stateCheck',
            'optype' : str(operation_type),
            'contact' : contact,
            }
        status = self.sendMessage(argdict)
        return status


    # This group of state checks pertain to the softphone entity
    #===========================================================
    def isDialing(self):
        status = self.stateCheck(ccc.OT_CHECK_ISDIALING)
        return status


    def isRinging(self):
        status = self.stateCheck(ccc.OT_CHECK_ISRINGING)
        return status


    def hasCalls(self):
        status = self.stateCheck(ccc.OT_CHECK_HASCALLS)
        return status


    def checkUserState(self):
        argdict = {
            'method' : 'checkUserState',
            }
        stateconstant = self.sendMessage(argdict)

        # The return value is a state-defining constant; map it to the
        # corresponding string for the caller's convenience
        if stateconstant not in ccc.userstate_strings:
            print 'Error, unrecognized state constant %d returned by checkUserState()' % stateconstant
            stateconstant = ccc.USERSTATE_UNDEFINED

        return ccc.userstate_strings[stateconstant]


    #===========================================================
    # This group of state checks pertain to each individual call
    # If no contact is specified, they apply to the 'current call'.
    #===========================================================
    def isConnected(self, contact=''):
        status = self.stateCheck(ccc.OT_CHECK_ISCONNECTED, contact)
        return status


    def isMuted(self, contact=''):
        status = self.stateCheck(ccc.OT_CHECK_ISMUTED, contact)
        return status


    def isHeld(self, contact=''):
        status = self.stateCheck(ccc.OT_CHECK_ISHELD, contact)
        return status


    def isInConference(self, contact=''):
        status = self.stateCheck(ccc.OT_CHECK_ISINCONFERENCE, contact)
        return status
    #===========================================================


    def SetSoftphoneToIdle(self):
        return False


    def SetMediaPlayerToIdle(self):
        return False


    def RestartSoftphone(self):
        return False


    def RestartMediaPlayer(self):
        return False


    def RestartSpokes(self):
        return False


    def MediaPlay(self):
        return STATUS_SUCCESS


    def MediaPause(self):
        return STATUS_SUCCESS


    def MediaStop(self):
        return STATUS_SUCCESS


    def MediaRepeat(self):
        return STATUS_SUCCESS


    def MediaMute(self, bMute):
        return STATUS_SUCCESS


    def MediaAdjustVolume(self, bVolume):
        return STATUS_SUCCESS


    def GetState_FirstCall(self):
        return SPSTATE_UNKNOWN


    def GetState_SecondCall(self):
        return SPSTATE_UNKNOWN


    def GetState_MediaPlayer(self):
        return MP_SPCALL_UNKNOWN


    def GetState_DialTone(self):
        return 0;


def main():
    pass


if __name__ == '__main__':
    main()
