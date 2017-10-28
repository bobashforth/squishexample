#-------------------------------------------------------------------------------
# Name:        iUI_CallControl
# Purpose:     This is the core class for the iUI_CallControl interface,
#              used to drive the UI of softphones within the Plantronics
#              TDTF (Target Device Test Framework).
#
# Author:      rsalsbury/rashforth
#
# Created:     19/07/2012
# Copyright:   (c) Plantronics 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import test
import testData
import object
import objectMap
import squishinfo
import squish

import exceptions

import call_control_constants as ccc
from pdctest.CallControl.lib.iUI_Control import UI_Control

import __builtin__

class UI_CallControl(UI_Control):

    def __init__(self, sp_id):
        if sp_id not in ccc.sp_dict:
            test.log('Error, invalid softphone name')
        else:
            self.sp_id = sp_id
            self.name = ccc.get_sp_name(sp_id)

        self.debug = True

        # This shared variable lets all call manipulation functions know the ID of the current caller
        self.currentcall_contact = None
        self.calls = {}
        self.callstack = []

        self.incomingcall_contact = None
        self.currentcall_contact = None


    def phonify(self, pnumber):

        areacode = pnumber[1:4]
        exchange = pnumber[4:7]
        number = pnumber[7:11]
        contact = '+1 (%s) %s-%s' % (areacode, exchange, number)

        return contact


    '''
    Note that this method provides the algorithm for using a softphone instance
    to dial an input string of digits. The actual implementation of each
    specific digit 'dial' is provided by each derived softphone class.
    '''
    def dial_digitstring(self, dstring):
        status = ccc.STATUS_SUCCESS
        if dstring is not None:
            i = 0
            while i < len(dstring):
                # Doublecheck that we haven't hit an error
                if status is ccc.STATUS_SUCCESS:
                    digit = dstring[i]
                    if digit < '0' or digit > '9':
                        status = ccc.STATUS_FAILURE
                    else:
                        if digit is '0':
                            status = self.dial_zero()
                        elif digit is '1':
                            status = self.dial_one()
                        elif digit is '2':
                            status = self.dial_two()
                        elif digit is '3':
                            status = self.dial_three()
                        elif digit is '4':
                            status = self.dial_four()
                        elif digit is '5':
                            status = self.dial_five()
                        elif digit is '6':
                            status = self.dial_six()
                        elif digit is '7':
                            status = self.dial_seven()
                        elif digit is '8':
                            status = self.dial_eight()
                        elif digit is '9':
                            status = self.dial_nine()
                        else:
                            # Should never get here, but fail anyway
                            status = ccc.STATUS_FAILURE
                    i += 1
                else:
                    break;
        else:
            status = ccc.STATUS_FAILURE

        return status


    def press_call_button(self):
        return ccc.STATUS_SUCCESS


    # This base-class method depends on softphone-specific methods
    def isIdle(self):
        if self.hasCalls() or self.isRinging() or self.isDialing():
            return True
        else:
            return False


    def PlaceCall(self, operation_type, contact):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def setAutoAnswer(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def AnswerCall(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def DeclineCall(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def EndCall(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def SetCallMuting(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def HoldResumeCall(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def TransferCallAudio(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def TransferCall(self, operation_type, newcontact, contact):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def AdjustCallVolume(self, operation_type):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def ConferenceCall(self, operation_type, contact_one, contact_two):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    def SetSoftphoneToIdle(self):
        return ccc.STATUS_FEATURE_NOT_SUPPORTED


    '''
    The following routines comprise the internal call-state tracking
    system. They should be valid across different platforms and softphones.
    '''
    def addCall(self, contact_id, status, isConferenceCall=False):
        # If we have an entry for this contact_id already, keep it and
        # just update status
        if contact_id in self.calls:
            newCall = self.calls[contact_id]
            newCall.status = status
        else:
            newCall = self.SoftphoneCall(contact_id, status)
            self.callstack.append(contact_id)

        # The rest is done whether it is a new call or being updated
        if isConferenceCall:
            newCall.isconferencecall = True
        self.calls[contact_id] = newCall
        self.currentcall_contact = contact_id

        return newCall


    def removeCall(self, contact_id):
        # If we have an entry for this contact_id already, free the object memory
        # and overwrite with a new object
        if contact_id in self.calls:
            del self.calls[contact_id]
        if contact_id in self.callstack:
            self.callstack.remove(contact_id)
            if self.hasCalls():
                currentcall_contact = self.callstack[len(self.callstack)-1]
            else:
                currentcall_contact = ''


    def getCall(self, contact_id):
        if contact_id not in self.calls:
            test.log('Error, no call connected for %s' % contact_id)
            return None
        else:
            return self.calls[contact_id]

    # Convenience routine for API access to all state checking methods
    def stateCheck(self, optype, contact):

        state = True
        if optype == ccc.OT_CHECK_ISRINGING:
            state = self.isRinging()
        elif optype == ccc.OT_CHECK_ISDIALING:
            state = self.isDialing()
        elif optype == ccc.OT_CHECK_HASCALLS:
            state = self.hasCalls()
        elif optype == ccc.OT_CHECK_ISCONNECTED:
            if not contact:
                contact = self.currentcall_contact
            state = self.isConnected(contact)
        elif optype == ccc.OT_CHECK_ISMUTED:
            if not contact:
                contact = self.currentcall_contact
            state = self.checkMuteState(contact)
        elif optype == ccc.OT_CHECK_ISHELD:
            if not contact:
                contact = self.currentcall_contact
            state = self.checkHoldState(contact)
        elif optype == ccc.OT_CHECK_ISINCONFERENCE:
            if not contact:
                contact = self.currentcall_contact
            state = self.checkInConferenceState(contact, 'In a conference call')
        else:
            # Eventually we should return a dictionary that contains both
            # status and any additional return args as appropriate. In the
            # meantime, failure handling of state checks is 'suboptimal'
            test.log('Error, invalid optype passed to stateCheck')
            status = ccc.STATUS_FAILURE

        return state


    def checkUserState(self):
        pass


    # Utility class, used solely for tracking call state
    class SoftphoneCall(__builtin__.object):


        def __init__(self, contact_id, status):

            self.contact_id = contact_id
            self.muted = False
            self.held = False
            self.inconference = False

            # This variable distinguishes between a direct call and a conference
            # call. It is needed because the windows and interactions are different,
            # and we need to know this to accurately track state and state transitions.
            self.isconferencecall = False

            self.status = status


        def logInfo(self):
            test.log('contact_id: %s' % self.contact_id)
            test.log('muted: %s' % str(self.muted))
            test.log('held: %s' % str(self.held))
            test.log('inconference: %s' % str(self.inconference))
            test.log('status: %s' % str(self.status))


