#-------------------------------------------------------------------------------
# Name:        microsoft_lync_CallControl
# Purpose:     This is the microsoft_lync-specific iUI_CallControl interface
#               class, derived from the core iUI_CallControl class.
#
# Author:      rashforth
#
# Created:     19/07/2012
# Copyright:   (c) Plantronics 2012
#-------------------------------------------------------------------------------

import test
import testData
import object
import objectMap
import squishinfo
import squish

import __builtin__
import os.path
import sys
import re
import exceptions
import pdctest.CallControl.lib.call_control_constants as ccc
import pdctest.CallControl.lib.iUI_CallControl as uicc


class Microsoft_lync_CallControl(uicc.UI_CallControl):


    def __init__(self, sp_id):

        # First call the __init__() method of the base class
        status = super(Microsoft_lync_CallControl, self).__init__(sp_id)

        # Then attach to the AUT associated with this softphone application
        try:

            autpath = ccc.get_sp_aut_path(sp_id)
            exename = os.path.basename(autpath)
            autname = (exename.split('.'))[0]

            squish.attachToApplication(autname)
        except Exception as e:
            test.log('Error in attaching to application %s' % autname)
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        # Compile regular expressions that will be used repeatedly
        self.incoming_call_regex = re.compile(r'Phone Call from +.*(.*) Lync 2010\.  Press Windows key plus A to accept.' )
        self.volume_extractor_regex = re.compile('Speakers.*:\s*(\d*)%')


    '''
    The methods below are the UI actions needed to dial each specific digit
    on the keypad. The algorithm for sequentially dialing all digits in an
    input string is provided by the base UI_CallControl class.
    '''

    def dial_zero(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .0_Button"))
        except Exception as e:
            test.log('Error in dial_zero() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_one(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .1_Button"))
        except Exception as e:
            test.log('Error in dial_one() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_two(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .2_Button"))
        except Exception as e:
            test.log('Error in dial_two() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_three(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .3_Button"))
        except Exception as e:
            test.log('Error in dial_three() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_four(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .4_Button"))
        except Exception as e:
            test.log('Error in dial_four() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_five(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .5_Button"))
        except Exception as e:
            test.log('Error in dial_five() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_six(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .6_Button"))
        except Exception as e:
            test.log('Error in dial_six() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_seven(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .7_Button"))
        except Exception as e:
            test.log('Error in dial_seven() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_eight(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .8_Button"))
        except Exception as e:
            test.log('Error in dial_eight() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_nine(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .9_Button"))
        except Exception as e:
            test.log('Error in dial_nine() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def press_call_button(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject(":Microsoft Lync          .Call_Button"))
        except Exception as e:
            test.log('Error in press_call_button() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def PlaceCall(self, operation_type, contact_id):
        status = ccc.STATUS_SUCCESS
        methodname = 'PlaceCall'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.PLACE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_DRAGDROP_TO_PHONE_MENU]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:

            # Strip the '+' from the beginning of the contact_id if it is present
            contact_id = contact_id.strip('+')

            # Make sure we're not hidden
            squish.setForegroundWindow(squish.waitForObject("{text='Microsoft Lync          ' type='Window'}"))

            if operation_type == ccc.OT_CONTACT_NAME:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED

            elif operation_type == ccc.OT_PHONE_NUMBER_MAIN_WND:
                try:
                    areacode = contact_id[1:4]
                    exchange = contact_id[4:7]
                    number = contact_id[7:11]
                    squish.clickButton(squish.waitForObject(":Microsoft Lync          .Contacts_Button"))
                    squish.type(squish.waitForObject(":Microsoft Lync          _Window_2"), contact_id)
                    propstring = "{container=':Microsoft Lync          _Window' text?='?*%s?*%s?*%s' type='ListItem'}" % (areacode, exchange, number)
                    squish.mouseClick(squish.waitForObject(propstring))
                except Exception as e:
                    test.log('Error in PlaceCall() method')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    self.press_call_button()

            elif operation_type == ccc.OT_PHONE_NUMBER_DIALPAD:
                try:
                    # Put the softphone in dialpad mode
                    squish.clickButton(squish.waitForObject(":Microsoft Lync          .Phone_Button"))

                    # Call the base class to translate the digit string into digit presses
                    status = super(Microsoft_lync_CallControl, self).dial_digitstring(contact_id)
                    squish.snooze(1)
                except Exception as e:
                    test.log('Error in PlaceCall() method')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

                # Finally, press the call button
                if status == ccc.STATUS_SUCCESS:
                    status = self.press_call_button()

            elif operation_type == ccc.OT_EXTENSION_NUMBER_MAIN_WND:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED

            elif operation_type == ccc.OT_EXTENSION_NUMBER_DIALPAD:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED

            elif operation_type == ccc.OT_DRAGDROP_TO_CONVERSATION_WND:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED

            elif operation_type == ccc.OT_START_CONFERENCE_CALL:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED

            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        # Validation: Voice Controls Status shows text 'Calling...'
        calling_window = None
        if status == ccc.STATUS_SUCCESS:
            try:
                call_status_label = squish.waitForObject("{text='Voice Controls Status' description?='Calling?*'}")
                test.log('PlaceCall() method successfully placed a call')
                self.addCall(contact_id, 'dialing')
            except LookupError:
                test.log('Error, PlaceCall() operation did not place a call')
                status = ccc.STATUS_FAILURE

        return status


    def getIncomingCallContact(self):

        incoming_contact = ''
        try:
            call_window = squish.waitForObject("{text~='Phone Call from .*\.  Press Windows key plus A to accept.' type='Window'}")

            # Make sure the window is not covered at all
            squish.setForegroundWindow(call_window)

            regex_object = self.incoming_call_regex.match(call_window.text)
            if regex_object is None:
                test.log('Error, text of incoming call window is not formatted correctly')
                incoming_contact = ''
            else:
                incoming_contact = regex_object.group(1)

        except Exception as e:
            test.log('Error in PlaceCall() method')
            for item in e.args:
                test.log(str(item))
            incoming_contact = ''

        return incoming_contact


    def AnswerCall(self, operation_type):
        status = ccc.STATUS_SUCCESS
        methodname = 'AnswerCall'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op( ccc.ANSWER_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_PICKUP]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:

            incoming_contact = ''

            if not self.isRinging():
                test.log('Error, no incoming call to answer')
            else:

                incoming_contact = self.incomingcall_contact

                incoming_call_propstring = "{text~=r'Phone Call from %s.*\.  Press Windows key plus A to accept.' type='Window'}" % incoming_contact
                answer_propstring = "{container=%s text='' type='Button'}" % incoming_call_propstring

                try:
                    call_answer_button = squish.waitForObject("{text='' type='Button'}")
                except Exception as e:
                    test.log('Error in AnswerCall() method')
                    for item in e.args:
                        test.log(str(item))

                    status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:
                if operation_type == ccc.OT_TOAST:

                    try:
                        incoming_call_window = squish.waitForObject(answer_propstring)

                        # Make sure the window is not covered at all
                        squish.setForegroundWindow(incoming_call_window)

                        squish.clickButton(call_answer_button)
                    except Exception as e:
                        test.log('Error in AnswerCall() method')
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

                elif operation_type == ccc.OT_MAIN_WND:

                    try:
                        # Switch to call list view
                        squish.clickButton(squish.waitForObject(":Microsoft Lync          .Conversations_Button"))

                        # Now doubleclick on incoming call in list
                        list_item_text = '%s Lync 2010' % incoming_contact
                        property_string ="{container=':Microsoft Lync          _Window' text='%s' type='ListItem'}" % list_item_text
                        squish.doubleClick(squish.waitForObject(property_string))

                        # Click on answer call button
                        call_button_text = 'Answer incoming call from %s Lync 2010 (%s@?*) (Alt+C)' % (incoming_contact, incoming_contact.lower())
                        #call_button_text = 'Answer incoming call from %s Lync 2010 (?*) (Alt+C)' % incoming_contact
                        property_string = "{text?='%s' type='Button'}" % call_button_text
                        test.log('property string for Answer Call button: %s' % property_string)
                        squish.clickButton(squish.waitForObject(property_string))
                    except Exception as e:
                        test.log('Error in AnswerCall() method')
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE


                else:
                    # Should never get here, given the above checks... being safe
                    status = ccc.STATUS_INVALID_OPERATION_TYPE

            # Validation
            if status == ccc.STATUS_SUCCESS:
                # Wait a second to make sure the window has appeared
                squish.snooze(1)
                if self.connectionWindow(incoming_contact):
                    test.log('AnswerCall() successfully answered the call')
                    if incoming_contact in self.calls:
                        # This means that someone checked isRinging(); just change the status
                        self.calls[incoming_contact].status = 'connected'
                        self.currentcall_contact = incoming_contact
                    else:
                        # We need to create a new call object and add it to calls and callstack
                        self.addCall(incoming_contact, 'connected')
                else:
                    test.log('Error, could not validate results of AnswerCall() method')
                    status = ccc.STATUS_FAILURE

        return status


    def DeclineCall(self, operation_type):

        status = ccc.STATUS_SUCCESS
        methodname = 'DeclineCall'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.DECLINE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_MAIN_WND]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:

            incoming_contact = ''

            if self.isRinging():

                incoming_contact = self.incomingcall_contact
                try:
                    incoming_call_propstring = "{text~=r'Phone Call from %s.*.  Press Windows key plus A to accept.' type='Window'}" % incoming_contact
                    decline_propstring = "{container=%s occurrence='3' text='' type='Button'}" % incoming_call_propstring
                    decline_button = squish.waitForObject(decline_propstring)

                    incoming_call_window = squish.waitForObject(incoming_call_propstring)

                    # Make sure we're not hidden
                    squish.setForegroundWindow(incoming_call_window)

                    if decline_button is not None:
                        squish.clickButton(decline_button)

                except Exception as e:
                    test.log('Error in DeclineCall() method')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

            else:
                test.log('Error, no ringing call to decline')
                status = ccc.STATUS_FAILURE

            # Validation: incoming_call_window disappears
            if status == ccc.STATUS_SUCCESS:
                try:
                    # Wait a second to be sure the incoming_call_window has been removed
                    squish.snooze(1)
                    dummy = squish.findObject(incoming_call_window)

                    # Note the inverted logic: The findObject() call will FAIL on SUCCESS,
                    # since if the window is found it means that we did not end the call.
                    test.log('Error, DeclineCall() operation did not successfully decline a call')
                    status = ccc.STATUS_FAILURE
                except LookupError:
                    test.log('DeclineCall() method successfully declined a call')

            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status


    def EndCall(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'EndCall'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.END_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        else:
            if operation_type in [ccc.OT_MAIN_WND]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:

            if not contact_id:
                contact_id = self.currentcall_contact
            if contact_id not in self.calls:
                test.log('Error, contact %s is not a valid call' % contact_id)
                status = ccc.STATUS_FAILURE

            else:
                # TBD: Make code below specific to the contact_id passed in (or current contact)
                currentcall = self.getCall(contact_id)
                if not currentcall:
                    test.log('Error, no valid call for %s, cannot end call')
                    status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_TOAST:
                try:
                    endcall_button = squish.waitForObject("{text='End call (Alt+Q)' type='Button'}")

                    # Make sure we're not hidden
                    squish.setForegroundWindow(endcall_button)

                    squish.clickButton(endcall_button)

                except Exception as e:
                    test.log('Error in EndCall() method')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

        # Validation: callInProgress_Window disappears
        if status == ccc.STATUS_SUCCESS:
            try:
                # Wait a second to be sure the call window has been removed
                squish.snooze(1)
                dummy = squish.findObject(":callInProgressWindow")

                # Note the inverted logic: The findObject() call will FAIL on SUCCESS,
                # since if the window is found it means that we did not end the call.
                test.log('Error, EndCall() operation did not successfully end a call')
                status = ccc.STATUS_FAILURE
            except LookupError:
                test.log('EndCall() method successfully ended a call')

        return status


    def SetCallMuting(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'SetCallMuting'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.SET_CALL_MUTING, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        currentcall = None

        if status == ccc.STATUS_SUCCESS:

            if not contact_id:
                contact_id = self.currentcall_contact

            currentcall = self.getCall(contact_id)
            if not currentcall:
                test.log('Error, no valid call for %s, cannot set call muting')
                status = ccc.STATUS_FAILURE

            titlebar = self.connectionWindow(contact_id)
            if titlebar:
                squish.setForegroundWindow(titlebar)

        if status == ccc.STATUS_SUCCESS:



            muteState = self.checkMuteState(contact_id)

            if operation_type == ccc.OT_MUTE:
                if muteState:
                    test.log('Call for %s is already muted' % contact_id)
                    currentcall.muted = True
                else:
                    try:
                        squish.clickButton(squish.waitForObject("{text='Mute microphone' type='Button'}"))
                        currentcall.muted = True
                    except Exception as e:
                        test.log('Error muting call')
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_UNMUTE:

                if not muteState:
                    test.log('Call for %s is already unmuted' % contact_id)
                    currentcall.muted = False
                else:
                    try:
                        squish.clickButton(squish.waitForObject("{text='Unmute microphone' type='Button'}"))
                        currentcall.muted = False
                    except Exception as e:
                        test.log('Error unmuting call')
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                # Validation - Check that the label of the button has toggled to the
                # value appropriate to the operation type
                # Wait a second to make sure the window reflects the updated status
                squish.snooze(1)
                muteState = self.checkMuteState(contact_id)

                if operation_type == ccc.OT_MUTE:
                    if muteState:
                        test.log('SetCallMuting() method successfully muted the call')
                        currentcall.muted = True
                    else:
                        test.log('SetCallMuting() method failed to mute the call')
                        currentcall.muted = False
                        status = ccc.STATUS_FAILURE

                elif operation_type == ccc.OT_UNMUTE:
                    if not muteState:
                        test.log('SetCallMuting() method successfully unmuted the call')
                        currentcall.muted = False
                    else:
                        test.log('SetCallMuting() method failed to unmute the call')
                        currentcall.muted = True
                        status = ccc.STATUS_FAILURE

        return status


    def HoldResumeCall(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'HoldResumeCall'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.HOLD_RESUME_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        currentcall = None

        if status == ccc.STATUS_SUCCESS:

            if not contact_id:
                contact_id = self.currentcall_contact
            if contact_id not in self.calls :
                test.log('Error, contact %s is not a valid call' % contact_id)
                status = ccc.STATUS_FAILURE

            # TBD: Make code below specific to the contact_id passed in (or current contact)
            currentcall = self.getCall(contact_id)
            if not currentcall:
                test.log('Error, no valid call for %s, cannot hold or resume call' % contact_id)
                status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                holdState = self.checkHoldState(contact_id)
                if operation_type == ccc.OT_HOLD:
                    try:
                        if holdState:
                            test.log('Error, call for %s is already on hold' % contact_id)
                        else:
                            holdbutton = squish.waitForObject("{text='Hold (Ctrl+Shift+H)' type='Button'}")

                            # Make sure we're not hidden
                            squish.setForegroundWindow(holdbutton)

                            squish.clickButton(holdbutton)

                    except Exception as e:
                        test.log('Error in putting call on hold for contact %s' % contact_id)
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

                elif operation_type == ccc.OT_RESUME:
                    try:
                        if not holdState:
                            test.log('Error, call for %s is already resumed' % contact_id)
                        else:
                            resumebutton = squish.waitForObject("{text='Resume call' type='Button'}")

                            # Make sure we're not hidden
                            squish.setForegroundWindow(resumebutton)

                            squish.clickButton(resumebutton)

                    except Exception as e:
                        test.log('Error in resuming call for contact %s' % contact_id)
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

                else:
                    # Should never get here, given the above checks... being safe
                    status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            # Validate operation
            # Wait a second to make sure the window reflects updated status
            squish.snooze(1)
            holdState = self.checkHoldState(contact_id)
            if operation_type == ccc.OT_HOLD:
                if holdState:
                    test.log('HoldResumeCall() method successfully put the call on hold')
                    currentcall.held = True
                else:
                    currentcall.held = False
                    test.log('HoldResumeCall() method failed to put the call on hold')
                    status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_RESUME:
                if not holdState:
                    currentcall.held = False
                    test.log('HoldResumeCall() method successfully resumed the call')
                else:
                    currentcall.held = True
                    test.log('HoldResumeCall() method failed to resume the call')
                    status = ccc.STATUS_FAILURE

        return status


    def findAllMenuItems(self):

        done = False
        occurrenceCount = 0
        device_itemlist =[]

        try:
            audiochangebutton = squish.findObject("{text='Change the audio device for this call' type='Button'}")
            # Make sure we're not hidden
            squish.setForegroundWindow(audiochangebutton)

            squish.clickButton(audiochangebutton)

        except LookupError:
            test.log('Error, cannot select Device Changer button')
            done = True

        # Give UI a second to ensure that transition is complete
        squish.snooze(1)

        while not done:
            occurrenceCount += 1
            propstring = "{type='MenuItem' text~='^Headset|PC|Custom.*$' occurrence='%s'}" % str(occurrenceCount)

            try:
                thisItem = squish.findObject(propstring)
                device_itemlist.append(thisItem)
            except LookupError:
                done = True

        return device_itemlist


    def TransferCallAudio(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'TransferCallAudio'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.TRANSFER_CALL_AUDIO, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type != ccc.OT_TRANSFER_AUDIO_NEXT and operation_type != ccc.OT_TRANSFER_AUDIO_PREV:
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            if not contact_id:
                contact_id = self.currentcall_contact
            if contact_id not in self.calls:
                test.log('Error, contact %s is not a valid call' % contact_id)
                status = ccc.STATUS_FAILURE

            # TBD: Make code below specific to the contact_id passed in (or current contact)
            currentcall = self.getCall(contact_id)
            if not currentcall:
                test.log('Error, no valid call for %s, cannot transfer call audio')
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
            inUseIndex = 0
            itemCount = 0
            newSelectionIndex = 0
            newSelectedMenuItem = None
            newSelectionText = ''

            menuItemList = self.findAllMenuItems()

            if not menuItemList:
                test.log('Error obtaining list of connected devices')
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
                # Find item in list which is marked as "In Use"
                i = 0
                inUseIndex = 0
                for menuItem in menuItemList:
                    # The list will be zero-indexed, though the occurrences begin with 1
                    if re.match('.*In Use.', menuItem.text):
                        inUseIndex = i
                    i += 1
                itemCount = i

                if operation_type == ccc.OT_TRANSFER_AUDIO_NEXT:
                    # Mod with the item count to wrap around to the beginning
                    newSelectionIndex = (inUseIndex + 1) % itemCount
                    optext = 'next'
                elif operation_type == ccc.OT_TRANSFER_AUDIO_PREV:
                    newSelectionIndex = (inUseIndex + itemCount - 1) % itemCount
                    optext = 'previous'
                newSelectedMenuItem = menuItemList[newSelectionIndex]

                if status == ccc.STATUS_SUCCESS:
                    try:
                        squish.mouseClick(newSelectedMenuItem)
                    except LookupError:
                        test.log('Error selecting new audio device')
                        status == ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Validate - Look for the new 'In Use' index and compare to new selection index
                    testMenuItemList = self.findAllMenuItems()
                    i = 0
                    newInUseIndex = 0
                    for menuItem in testMenuItemList:
                        if re.match('.*In Use.', menuItem.text):
                            newInUseIndex = i
                        i += 1

                    if newInUseIndex == newSelectionIndex:
                        test.log('TransferCallAudio() successfully transferred audio to %s device' % optext)
                    else:
                        test.log('TransferCallAudio() failed to transfer audio to %s device' % optext)
                        status = ccc.STATUS_FAILURE

                    '''
                    # Validation: Look for the MenuItem marked as 'In Use' and check a match of occurrence number
                    squish.clickButton(squish.waitForObject("{text='Change the audio device for this call' type='Button'}"))
                    propstring = "{type='MenuItem' text~='^.*In Use.*$' occurrence='%d'}" % inUseOccurrence
                    testItem = None

                    try:
                        testItem = squish.waitForObject(propstring)
                    except LookupError:
                        test.log('Error locating in use device after new selection')
                        status = ccc.STATUS_FAILURE
                    '''
                    # Pop down the menu
                    squish.clickButton(squish.waitForObject("{text='Change the audio device for this call' type='Button'}"))

        return status


    def TransferCall(self, operation_type, newcontact, current_contact=None):
        status = ccc.STATUS_SUCCESS
        methodname = 'TransferCall'

        areacode = ''
        exchange = ''
        number = ''

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.TRANSFER_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if not current_contact:
                current_contact = self.currentcall_contact
            if current_contact not in self.calls:
                test.log('Error, current_contact %s is not a valid call' % current_contact)
                status = ccc.STATUS_FAILURE
            else:
                currentcall = self.getCall(current_contact)
                if not currentcall:
                    test.log('Error, no valid call for %s, cannot set call muting')
                    status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            if newcontact is None:
                test.log('No contact number specified, TransferCall() failed.')
                status = ccc.STATUS_FAILURE
            else:
                areacode = newcontact[1:4]
                exchange = newcontact[4:7]
                number = newcontact[7:11]
                test.log('areacode=%s, exchange=%s, number=%s' % (areacode, exchange, number))
                # {container=':Transfer Call_Window' text='+1 (831) 704-7167' type='ListItem'}
                contact_listitemname = "{container=':Transfer Call_Window' text?='?*%s?*%s?*%s' type='ListItem'}" % (areacode, exchange, number)

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_TRANSFER_CALL:
                try:

                    transfercallbutton = squish.waitForObject("{container=':callInProgress_Window' text='TRANSFER' type='Button'}")

                    # Make sure we're not hidden
                    squish.setForegroundWindow(transfercallbutton)

                    squish.mouseClick(transfercallbutton)

                    squish.mouseClick(squish.waitForObject("{occurrence='2' text='' type='MenuItem'}"))

                    squish.mouseClick(squish.waitForObject(contact_listitemname))

                    squish.clickButton(squish.waitForObject("{container=':Transfer Call_Window' text='OK' type='Button'}"))

                except Exception as e:
                    test.log('Error in TransferCall() method')
                    for item in e.args:
                        test.log(str(item))

                    status = ccc.STATUS_FAILURE
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

            if status == ccc.STATUS_SUCCESS:
                # Validation: The call window should show the 'Transferring...' label text
                # Give the UI a second to reflect transition
                squish.snooze(1)

                transfermatchtext = 'Transferring call to +1 (%s) %s-%s?*' % (areacode, exchange, number)
                label_propstring = "{text='Voice Controls Status' type='Label' description?='%s'}" % transfermatchtext

                try:
                    callstatuslabel = squish.findObject(label_propstring)
                    test.log('TransferCall() successfully transferred the call')
                except LookupError:
                    status = ccc.STATUS_FAILURE
                    test.log('TransferCall() failed to successfully transfer the call')

        return status


    def AdjustCallVolume(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'AdjustCallVolume'
        slider=None
        slider_height = 0
        slider_value = 0
        new_slider_value = 0

        # Microsoft-determined; this value reflects reality, it doesn't define it. :-)
        bumpIncrement = 25

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.ADJUST_CALL_VOLUME, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if not contact_id:
                contact_id = self.currentcall_contact
            if contact_id not in self.calls:
                test.log('Error, contact %s is not a valid call' % contact_id)
                status = ccc.STATUS_FAILURE
            else:
                # TBD: Make code below specific to the contact_id passed in (or current contact)
                currentcall = self.getCall(contact_id)
                if not currentcall:
                    test.log('Error, no valid call for %s, cannot set call muting')
                    status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            volumeadjustbutton = squish.waitForObject("{text='Adjust volume or mute speakers' type='Button'}")

            # Make sure we're not hidden
            squish.setForegroundWindow(volumeadjustbutton)

            # Get the current slider value, which is the volume percentage
            try:
                # First pop up the slider, then get the object
                squish.clickButton(volumeadjustbutton)
            except Exception as e:
                test.log('Error 1 in getting slider object and value in AdjustCallVolume() method')
                for item in e.args:
                    test.log(str(item))
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            try:
                popup = squish.waitForObject("{name='Speakers'}")
            except Exception as e:
                test.log('Error 2 in getting slider object and value in AdjustCallVolume() method')
                for item in e.args:
                    test.log(str(item))
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            try:
                slider = squish.waitForObject(":Speaker Volume Popup_Slider")
            except Exception as e:
                test.log('Error 3 in getting slider object and value in AdjustCallVolume() method')
                for item in e.args:
                    test.log(str(item))
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            try:
                # Origin of objects is upper left, so a 'slider_height' value for y represents zero volume
                # and a y value of 0 represents maximum volume.
                slider_height = __builtin__.int(slider.height)
                slider_value = __builtin__.int(slider.value)

            except Exception as e:
                test.log('Error 4 in getting slider object and value in AdjustCallVolume() method')
                for item in e.args:
                    test.log(str(item))
                status = ccc.STATUS_FAILURE


        if status == ccc.STATUS_SUCCESS:
            # Set the click position. We always click at the extremes, and the resulting 'bump'
            # in volume change is determined by the application.
            if operation_type == ccc.OT_VOLUME_DOWN_SPEAKER:
                yClickValue = __builtin__.int(slider_height) - 1
                bumpIncrement *= -1
            elif operation_type == ccc.OT_VOLUME_UP_SPEAKER:
                yClickValue = 0
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

            try:
                squish.mouseClick(slider,
                        __builtin__.int(slider.width)/2,
                        yClickValue,
                        squish.MouseButton.LeftButton)

                # Get new volume percentage
                new_slider_value = __builtin__.int(slider.value)

                # Click again to remove the slider
                squish.clickButton(squish.waitForObject("{text='Adjust volume or mute speakers' type='Button'}"))

            except Exception as e:
                test.log('Error in adjusting volume')
                for item in e.args:
                    test.log(str(item))
                status = ccc.STATUS_FAILURE

        test.log('Initial value-%d, new value=%d' % (slider_value, new_slider_value))

        # Validation - Note that if Microsoft ever changes the 'bump' increment it must be updated
        if new_slider_value != (slider_value + bumpIncrement):
            status = ccc.STATUS_FAILURE

        return status


    # Note that this method does not recognize an implicit 'current contact,' but requires
    # specification of both contacts.
    def ConferenceCall(self, operation_type, contact_one, contact_two):

        methodname = 'ConferenceCall'

        super(Microsoft_lync_CallControl, self).log_method_called(self.sp_id, methodname)

        status = ccc.STATUS_SUCCESS

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.CONFERENCE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            status = ccc.STATUS_SUCCESS
            contact_one_listitem = None
            contact_two_listitem = None

            # Make sure we're not hidden
            squish.setForegroundWindow(squish.waitForObject(":Microsoft Lync          _Window"))

            if operation_type == ccc.OT_CONFERENCE_TWO_CONTACTS:
                # First, select conference contacts using Ctrl+left-button-select
                #

                listitem_one_propstring = "{type='ListItem' text?='%s?*'}" % contact_one
                listitem_two_propstring = "{type='ListItem' text?='%s?*'}" % contact_two
                try:
                    contact_one_listitem = squish.findObject(listitem_one_propstring)
                    contact_two_listitem = squish.findObject(listitem_two_propstring)
                except LookupError:
                    test.log('Error, could not locate contact listitems for conference call')
                    status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    try:
                        squish.mouseClick(contact_one_listitem, squish.MouseButton.LeftButton, squish.Modifier.Control)
                        squish.mouseClick(contact_two_listitem, squish.MouseButton.LeftButton, squish.Modifier.Control)

                    except Exception as e:
                        test.log('Error, could not select contacts for conference call')
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    # Now ctrl-right-click to bring up the conference call context menu,
                    # then left-click the Conference Call menuitem,
                    # then select 'Lync' as the mechanism for starting the conference call.
                    try:
                        squish.mouseClick(contact_one_listitem, squish.MouseButton.RightButton, squish.Modifier.Control)
                        squish.mouseClick(squish.waitForObject("{text='Start a Conference Call' type='MenuItem'}"))
                        squish.mouseClick(squish.waitForObject("{text?='?*Lync' type='MenuItem'}"))

                    except Exception as e:
                        test.log('Error initiating conference call')
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Wait a second to make sure the window reflects the updated status
                    squish.snooze(1)

                    contact_one_okay = False
                    contact_two_okay = False

                    if self.checkInConferenceState(contact_one, 'Available'):
                        contact_one_okay = True
                        test.log('contact_one okay')
                    if self.checkInConferenceState(contact_two, 'Available'):
                        contact_two_okay = True
                        test.log('contact_two okay')
                    if contact_one_okay and contact_two_okay:
                        test.log('ConferenceCall() successfully initiated a conference call')
                        # Arbitrarily identify the conference call by the first contact id
                        # Add the optional qualifier to addCall() which marks a conference call
                        self.addCall(contact_one, 'dialing', True)
                    else:
                        test.log('ConferenceCall() failed to successfully initiate a conference call')
                        status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_CONFERENCE_TWO_CALLS:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED
            elif operation_type == ccc.OT_CONFERENCE_DRAGDROP:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status


    def checkCallStatus(self, statestring, contact=''):

        returnval = False

        if not contact:
            contact = self.currentcall_contact

        test.log('checking status \'%s\' for contact %s' % (statestring, contact))
        status = self.refreshCallStatus()

        if status != ccc.STATUS_SUCCESS:
            test.log('Error in refreshing call status')
            returnval = False
        else:
            if contact in self.calls:
                if self.calls[contact].status != statestring:
                    returnval = False
                else:
                    returnval = True

        return returnval


    def isDialing(self):

        # We need to check for both direct calls and outgoing conference call requests, which have different
        # windows and text
        directdialing = False
        conferencing = False

        try:
            #call_status_label = squish.waitForObject("{text='Voice Controls Status' description?='Calling?*'}")
            call_status_label = squish.findObject("{text='Voice Controls Status' description?='Calling?*'}")
            directdialing = True
            test.log('Found directdialing calling window')
        except LookupError:
            pass

        conferencing = self.checkInConferenceState(self.currentcall_contact, 'Available')


        return (directdialing or conferencing)


    def isRinging(self):

        self.incomingcall_contact = self.getIncomingCallContact()
        if self.incomingcall_contact is None:
            return False
        else:
            self.addCall(self.incomingcall_contact, 'ringing')
            return True



    def isConnected(self, contact):

        return self.checkCallStatus('connected', contact)


    def isMuted(self, contact):

        return self.checkMuteState(contact)


    def connectionWindow(self, contact):

        # Wait a second in case we were called immediately after iniation of a UI state change
        squish.snooze(1)

        titlebar = None

        testcontact = contact.strip('+')

        # If the contact is an unformatted phone number, 'phoneformat' it
        if testcontact.isdigit():
            areacode = testcontact[1:4]
            exchange = testcontact[4:7]
            number = testcontact[7:11]
            contact = '+1 (%s) %s-%s' % (areacode, exchange, number)

        try:
            titlebar = squish.findObject("{name='Header Pane Primary Text' type='Label' role='StaticText' value?='%s'}" % contact)
        except LookupError:
            # Not finding the connection window is a perfectly valid result
            pass

        return titlebar


    def checkMuteState(self, contact):

        muteState = False

        #titlebar = self.connectionWindow(contact)
        #if titlebar:
        #    squish.setForegroundWindow(titlebar)

        try:
            dummy = squish.findObject("{text='Unmute microphone' type='Button'}")
            squish.setForegroundWindow(dummy)
            muteState = True
        except LookupError:
            pass

        try:
            dummy = squish.findObject("{text='Mute microphone' type='Button'}")
            squish.setForegroundWindow(dummy)
            muteState = False
        except:
            pass

        return muteState


    def checkHoldState(self, contact):

        holdState = False

        titlebar = self.connectionWindow(contact)
        if titlebar:
            squish.setForegroundWindow(titlebar)

        try:
            dummy = squish.findObject("{text='Resume call' type='Button'}")
            holdState = True
        except LookupError:
            pass

        return holdState


    def checkInConferenceState(self, contact, statestring):

        titlebar = self.connectionWindow(contact)
        if titlebar:
            squish.setForegroundWindow(titlebar)

        contact_propstring = "{text?='%s?*(%s)' type='ListItem'}" % (contact, statestring)
        try:
            dummy = squish.findObject(contact_propstring)
            return True
        except LookupError:
            return False


    def refreshCallStatus(self):
        status = ccc.STATUS_SUCCESS

        if self.callstack:
            # Confirm status of each listed call
            # Use a copy of the list to iterate, so we can delete items safely
            listcopy = self.callstack[:]
            for contact in listcopy:
                call = self.calls[contact]

                # Direct calls and conference calls have different windows and text indicating
                # 'dialing' and 'connected' states
                if not call.isconferencecall:
                    if call.status == 'dialing' or call.status is 'ringing':
                        # This is our last known state for this contact,
                        # check its current status
                        if self.connectionWindow(contact):
                            call.status = 'connected'

                    elif call.status == 'connected':
                        if not self.connectionWindow(contact):
                            self.removeCall(contact)
                else:
                    if call.status == 'dialing':
                        # This is our last known state for this contact,
                        # check its current status
                        if self.checkInConferenceState(contact, 'In a conference call'):
                            call.status = 'connected'

                    # For conference calls, we check to see if we're the only one in the 'conference'
                    #elif call.status == 'connected':
                    #    if not self.connectionWindow(contact):
                    #        self.removeCall(contact)

        return status


    def hasCalls(self):

        self.refreshCallStatus()
        if self.callstack:
            return True
        else:
            return False


    def checkUserState(self):

        stateconstant=ccc.USERSTATE_UNDEFINED
        statestring=''

        propstring = "{text='My Status' role='StaticText' type='Label'}"

        try:
            statebutton = squish.findObject(propstring)
            statestring = statebutton.value
            if statestring in ccc.userstate_constants:
                stateconstant = ccc.userstate_constants[statestring]
            else:
                test.log('Error, statestring \'%s\' is not recognized' % statestring)
        except:
            test.log('Error, could not locate user state button')

        return stateconstant


    def SetSoftphoneToIdle(self):

        # Make sure we're not hidden
        squish.setForegroundWindow(squish.waitForObject("{text='Microsoft Lync          ' type='Window'}"))

        for call in self.calls:
            self.removeCall(call.contact_id)

        self.incomingcall_contact = ''
        self.currentcall_contact = ''

        # TBD: Abort any calls that are dialing and decline any calls that are ringing
        # (May implement by logging out and logging back in)

