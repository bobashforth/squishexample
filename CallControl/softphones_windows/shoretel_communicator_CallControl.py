#-------------------------------------------------------------------------------
# Name:        shoretel_communicator_CallControl
# Purpose:     This is the shoretel_communicator-specific iUI_CallControl
#              interface class, derived from the core iUI_CallControl class.
#
# Author:      rsalsbury/rashforth
#
# Created:     26/08/2012
# Copyright:   (c) Plantronics 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import test
import testData
import object
import objectMap
import squishinfo
import squish

import pdctest.CallControl.lib.call_control_constants as ccc
import pdctest.CallControl.lib.iUI_CallControl as uicc

class Shoretel_communicator_CallControl(uicc.UI_CallControl):

    def PlaceCall(self, operation_type, contact):
        status = ccc.STATUS_SUCCESS
        methodname = 'PlaceCall'

        super(Shoretel_communicator_CallControl, self).log_method_called(self.phone_name, methodname)

        # Validate operation type for function and for this softphone
        if not isvalid_func_op(PLACE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_PHONE_NUMBER_DIALPAD,
                                ccc.OT_EXTENSION_NUMBER_DIALPAD,
                                ccc.OT_DRAGDROP_TO_PHONE_MENU,
                                ccc.OT_DRAGDROP_TO_CONVERSATION_WND]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_CONTACT_NAME:
                pass
            if operation_type == ccc.OT_PHONE_NUMBER_MAIN_WND:
                pass
            elif operation_type == ccc.OT_EXTENSION_NUMBER_MAIN_WND:
                pass
            elif operation_type == ccc.OT_START_CONFERENCE_CALL:
                pass
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status

    def AnswerCall(self, operation_type):
        status = ccc.STATUS_SUCCESS
        methodname = 'AnswerCall'

        super(Shoretel_communicator_CallControl, self).log_method_called(self.phone_name, methodname)

        # Validate operation type for function and for this softphone
        if not isvalid_func_op(ANSWER_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_PICKUP]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
           if operation_type == ccc.OT_TOAST:
                pass
           elif operation_type == ccc.OT_MAIN_WND:
                pass
           else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status

    def DeclineCall(self, operation_type):
        status = ccc.STATUS_SUCCESS
        methodname = 'DeclineCall'

        super(Shoretel_communicator_CallControl, self).log_method_called(self.phone_name, methodname)

        # Validate operation type for function and for this softphone
        if not isvalid_func_op(DECLINE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:
           if operation_type == ccc.OT_TOAST:
                pass
           elif operation_type == ccc.OT_MAIN_WND:
                pass
           else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status

    def EndCall(self, operation_type):
        status = ccc.STATUS_SUCCESS
        methodname = 'EndCall'

        super(Shoretel_communicator_CallControl, self).log_method_called(self.phone_name, methodname)

        # Validate operation type for function and for this softphone
        if not isvalid_func_op(END_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_TOAST]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_MAIN_WND:
                pass
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status

    def MuteHoldTransferCall(self, operation_type):
        status = ccc.STATUS_SUCCESS
        methodname = 'MuteHoldTransferCall'

        super(Shoretel_communicator_CallControl, self).log_method_called(self.phone_name, methodname)

        # Validate operation type for function and for this softphone
        if not isvalid_func_op(MUTE_HOLD_TRANSFER_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_TRANSFER_AUDIO_HS_TO_PC,
                                ccc.OT_TRANSFER_AUDIO_PC_TO_HS,
                                ccc.OT_VOLUME_UP_SPEAKER,
                                ccc.OT_VOLUME_DOWN_SPEAKER]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_MUTE:
                pass
            elif operation_type == ccc.OT_UNMUTE:
                pass
            elif operation_type == ccc.OT_HOLD:
                pass
            elif operation_type == ccc.OT_RESUME:
                pass
            elif operation_type == ccc.OT_TRANSFER_CALL:
                pass
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status

    def TwoCalls(self, operation_type, contact):
        status = ccc.STATUS_SUCCESS
        methodname = 'TwoCalls'

        super(Shoretel_communicator_CallControl, self).log_method_called(self.phone_name, methodname)

        # Validate operation type for function and for this softphone
        if not isvalid_func_op(TWO_CALLS, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_TWO_CALLS_INCOMING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ACTIVE_INCOMING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ACTIVE_ONHOLD:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ACTIVE_DIALING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ACTIVEMUTED_INCOMING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ACTIVEMUTED_ONHOLD:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ACTIVEMUTED_DIALING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_INCOMING_DIALING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_DIALING_INCOMING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ONHOLD_DIALING:
                pass
            elif operation_type == ccc.OT_TWO_CALLS_ONHOLDMUTED_DIALING:
                pass
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status

    def ConferenceCall(self, operation_type):
        status = ccc.STATUS_SUCCESS
        methodname = 'ConferenceCall'

        super(Shoretel_communicator_CallControl, self).log_method_called(self.phone_name, methodname)

        # Validate operation type for function and for this softphone
        if not isvalid_func_op(CONFERENCE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_CONFERENCE_TWO_CONTACTS,
                                ccc.OT_CONFERENCE_DRAGDROP]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_CONFERENCE_TWO_CALLS:
                pass
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status

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
        return Failure_not_supported

    def MediaPause(self):
        return Failure_not_supported

    def MediaStop(self):
        return Failure_not_supported

    def MediaRepeat(self):
        return Failure_not_supported

    def MediaMute(self, bMute):
        return Failure_not_supported

    def MediaAdjustVolume(self, bVolume):
        return Failure_not_supported

    def GetState_FirstCall(self):
        return SPSTATE_UNKNOWN

    def GetState_SecondCall(self, softphone_name):
        return SPSTATE_UNKNOWN

    def GetState_MediaPlayer(self):
        return MP_SPCALL_UNKNOWN

    def GetState_DialTone(self):
        return 0;

def main():
    pass

if __name__ == '__main__':
    main()
