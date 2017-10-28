#-------------------------------------------------------------------------------
# Name:        microsoft_lync_2013_CallControl
# Purpose:     This is the microsoft_lync_2013-specific iUI_CallControl interface
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


class Microsoft_lync_2013_CallControl(uicc.UI_CallControl):


    def __init__(self, sp_id):

        # First call the __init__() method of the base class
        status = super(Microsoft_lync_2013_CallControl, self).__init__(sp_id)

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
        self.userstate_regex = r'Lync - Signed in (.*)'
        self.incoming_call_regex = re.compile(r'Phone Call from +.*(.*) Lync 201.\.  Press Windows key plus A to accept.' )
        self.volume_extractor_regex = re.compile('Speakers.*:\s*(\d*)%')

        try:
            self.lync_window = squish.findObject("{text='Lync' type='Window'}")
        except LookupError:
            self.lync_window = None


    '''
    The methods below are the UI actions needed to dial each specific digit
    on the keypad. The algorithm for sequentially dialing all digits in an
    input string is provided by the base UI_CallControl class.
    '''

    def dial_zero(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='0' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_zero() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_one(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='1' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_one() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_two(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='2' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_two() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_three(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='3' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_three() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_four(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='4' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_four() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_five(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='5' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_five() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_six(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='6' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_six() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_seven(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='7' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_seven() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_eight(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='8' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_eight() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def dial_nine(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='9' type='Button'}"))
        except Exception as e:
            test.log('Error in dial_nine() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def press_call_button(self):
        status = ccc.STATUS_SUCCESS
        try:
            squish.clickButton(squish.waitForObject("{text='Call' type='Button'}"))
        except Exception as e:
            test.log('Error in press_call_button() method')
            for item in e.args:
                test.log(str(item))
            status = ccc.STATUS_FAILURE

        return status


    def PlaceCall(self, operation_type, contact_id):
        status = ccc.STATUS_SUCCESS
        methodname = 'PlaceCall'

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

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
            squish.setForegroundWindow(self.lync_window)

            if operation_type == ccc.OT_CONTACT_NAME:
                status = ccc.STATUS_SUCCESS

                try:
                # Convert to 'lastname, firstname' if we have a last name
                    (firstname, lastname) = contact_id.split(' ')
                    contact_id = '%s, %s' % (lastname, firstname)
                except:
                    # If we fail, just carry on... no need to convert
                    pass

                try:
                    # Put the softphone in 'Contacts' mode, then select and call the contact
                    squish.clickButton(squish.findObject("{text='Contacts' role='PushButton' type='Button'}"))
                    listitem = squish.findObject("{text?='*%s*' role='ListItem'}" % contact_id)
                    squish.mouseClick(listitem, squish.MouseButton.RightButton)
                    squish.clickButton(squish.waitForObject("{text='Call' type='MenuItem'}"))
                    squish.clickButton(squish.waitForObject("{text?='Work*' type='MenuItem'}"))

                except Exception as e:
                    self.display_all_objects()
                    test.log('Error in PlaceCall() method')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_PHONE_NUMBER_MAIN_WND:

                contact_listitem = None

                areacode = contact_id[1:4]
                exchange = contact_id[4:7]
                number = contact_id[7:11]

                try:
                    # Put the softphone in 'Contacts' mode
                    squish.clickButton(squish.findObject("{text='Contacts' role='PushButton' type='Button'}"))
                    edit_window = squish.waitForObject("{container={text='Lync' type='Window'} type='Edit'}")
                    squish.clickButton(edit_window)
                    squish.type(edit_window, contact_id)
                    squish.snooze(1)
                    squish.nativeType("<Return>")
                    test.log('After pressing return to call contact')
                except Exception as e:
                    test.log('Error, PlaceCall() cannot dial specified contact number')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_PHONE_NUMBER_DIALPAD:
                try:
                    # Put the softphone in dialpad mode
                    squish.clickButton(squish.findObject("{text='Phone' role='PushButton' type='Button'}"))

                    # Call the base class to translate the digit string into digit presses
                    status = super(Microsoft_lync_2013_CallControl, self).dial_digitstring(contact_id)
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

        # Validation: call status label shows text 'Calling...'
        calling_window = None
        if status == ccc.STATUS_SUCCESS:

            confirm_contact = ''
            try:
                if contact_id.isdigit():
                    confirm_contact = self.phonify(contact_id)
                else:
                    confirm_contact = contact_id

                self.display_all_objects()
                #call_status_label = squish.waitForObject("{text='Calling %s...' type='Label'}" % confirm_contact)
                call_status_label = squish.waitForObject("{text?='%s*' type='Label' role='StaticText'}" % confirm_contact)
                test.log('PlaceCall() method successfully placed a call')
                self.addCall(contact_id, 'dialing')

            except LookupError:
                test.log('Error, PlaceCall() operation did not place a call')
                status = ccc.STATUS_FAILURE

        return status


    def getIncomingCallContact(self):

        incoming_contact = ''
        try:
            # Because the significant text is a separate object from the contact info, we need to
            # rely on the matching x value and a unique height to identify the caller contact info.
            # (Both have the same x value and all attributes are identical except for the text
            # and the y and height attributes.)
            incomingcall_label = squish.waitForObject("{text='is calling you' role='StaticText'}")
            incomingcall_propstring = "{role='StaticText' x='%d' height='22'}" % __builtin__.int(incomingcall_label.x)
            incomingcall_contact_label = squish.findObject(incomingcall_propstring)
            incoming_contact = incomingcall_contact_label.text

        except Exception as e:
            test.log('Error in retrieving incoming call data')
            for item in e.args:
                test.log(str(item))
            incoming_contact = ''

        return incoming_contact


    def AnswerCall(self, operation_type):
        status = ccc.STATUS_SUCCESS
        methodname = 'AnswerCall'

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

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

            if status == ccc.STATUS_SUCCESS:
                if operation_type == ccc.OT_TOAST:

                    incomingcall_window = squish.findObject("{role='Window' name='Lync Toast'}")

                    try:
                        call_answer_button = squish.waitForObject("{text='Accept' role='PushButton'}")
                        squish.clickButton(call_answer_button)
                    except Exception as e:
                        test.log('Error answering call in AnswerCall() method')
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

                elif operation_type == ccc.OT_MAIN_WND:

                    try:
                        # Switch to call list view
                        squish.clickButton(squish.waitForObject(":Microsoft Lync          .Conversations_Button"))

                        # Now doubleclick on incoming call in list
                        list_item_text = '%s Lync 201?' % incoming_contact
                        property_string ="{container=':Microsoft Lync 2013          _Window' text?='%s' type='ListItem'}" % list_item_text
                        squish.doubleClick(squish.waitForObject(property_string))

                        # Click on answer call button
                        call_button_text = 'Answer incoming call from %s Lync 201? (%s@*) (Alt+C)' % (incoming_contact, incoming_contact.lower())
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
                    test.log('Error, could not validate results of AnswerCall() method for contact %s' % incoming_contact)
                    status = ccc.STATUS_FAILURE

        return status


    def DeclineCall(self, operation_type):

        status = ccc.STATUS_SUCCESS
        methodname = 'DeclineCall'

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.DECLINE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_MAIN_WND]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:

            incoming_contact = ''

            if self.isRinging():

                    incomingcall_window = squish.findObject("{role='Window' name='Lync Toast'}")

                    try:
                        call_answer_button = squish.waitForObject("{text='Ignore' role='PushButton'}")
                        squish.clickButton(call_answer_button)
                    except Exception as e:
                        test.log('Error declining call in DeclineCall() method')
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
                    dummy = squish.findObject("{role='Window' name='Lync Toast'}")

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

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

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

                endcall_propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='1'}" % contact_id
                try:
                    endcall_button = squish.findObject(endcall_propstring)

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
                dummy = squish.findObject("{type='Window' text?='*%s'}" % contact_id)

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

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

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


            callcontrol_propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='8'}" % contact_id
            mute_propstring = "{role='PushButton' text='Mute'}"
            unmute_propstring = "{role='PushButton' text='Unmute'}"
            targetMutestate = False
            action_propstring = ''

            muteState = self.checkMuteState(contact_id)

            if operation_type == ccc.OT_MUTE:

                # For validation
                targetMutestate = True
                action_propstring = mute_propstring

                if muteState:
                    test.log('Call for %s is already muted' % contact_id)

            elif operation_type == ccc.OT_UNMUTE:

                # For validation
                targetMutestate = False
                action_propstring = unmute_propstring

                if not muteState:
                    test.log('Call for %s is already unmuted' % contact_id)

            # Action is the same for mute and unmute if we don't have targetMutestate
            if muteState != targetMutestate:
                try:
                    callcontrolgraphic = squish.findObject(callcontrol_propstring)
                    squish.mouseMove(callcontrolgraphic)

                    # Wait a sec for the popup to appear
                    squish.snooze(1)

                    squish.clickButton(action_propstring)

                except Exception as e:
                    test.log('Error locating and activating mute/unmute button')
                    for item in e.args:
                        test.log(str(item))
                        status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                # Validation - Check that the current mute state matches the
                # targetMutestate.
                # Move away from the button and wait two seconds to make sure the window reflects the updated status
                squish.mouseMove(10,10)
                squish.snooze(2)
                muteState = self.checkMuteState(contact_id)
                validState = (muteState == targetMutestate)

                if operation_type == ccc.OT_MUTE:
                    if validState:
                        test.log('SetCallMuting() method successfully muted the call')
                    else:
                        test.log('SetCallMuting() method failed to mute the call')
                        currentcall.muted = False
                        status = ccc.STATUS_FAILURE

                elif operation_type == ccc.OT_UNMUTE:
                    if validState:
                        test.log('SetCallMuting() method successfully unmuted the call')
                    else:
                        test.log('SetCallMuting() method failed to unmute the call')
                        currentcall.muted = False
                        status = ccc.STATUS_FAILURE

        return status


    def HoldResumeCall(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'HoldResumeCall'

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.HOLD_RESUME_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        currentcall = None

        if status == ccc.STATUS_SUCCESS:

            if not contact_id:
                contact_id = self.currentcall_contact

            currentcall = self.getCall(contact_id)
            if not currentcall:
                test.log('Error, no valid call for %s, cannot hold/resume call')
                status = ccc.STATUS_FAILURE

            titlebar = self.connectionWindow(contact_id)
            if titlebar:
                squish.setForegroundWindow(titlebar)

        if status == ccc.STATUS_SUCCESS:

            callcontrol_propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='8'}" % contact_id
            hold_propstring = "{role='PushButton' text='Hold Call'}"
            resume_propstring = "{role='PushButton' text='Resume Call'}"
            targetHoldstate = False
            action_propstring = ''

            holdState = self.checkHoldState(contact_id)

            if operation_type == ccc.OT_HOLD:

                # For validation
                targetHoldstate = True
                action_propstring = hold_propstring

                if holdState:
                    test.log('Call for %s is already held' % contact_id)

            elif operation_type == ccc.OT_RESUME:

                # For validation
                targetHoldstate = False
                action_propstring = resume_propstring

                if not holdState:
                    test.log('Call for %s is already resumed' % contact_id)

            # Action is the same for hold and resume if we don't have targetHoldstate
            if holdState != targetHoldstate:
                try:
                    callcontrolgraphic = squish.findObject(callcontrol_propstring)
                    squish.mouseMove(callcontrolgraphic)

                    # Wait a sec for the popup to appear
                    squish.snooze(1)

                    squish.clickButton(action_propstring)

                except Exception as e:
                    test.log('Error locating and activating hold/resume button')
                    for item in e.args:
                        test.log(str(item))
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS and operation_type == ccc.OT_HOLD:

                    try:
                        # Truly annoying: Because the 'Hold' action pops up a new graphic,
                        # which changes the occurrence number of all other graphics, close
                        # the offending window.
                        squish.mouseMove(10,10)
                        squish.snooze(2)

                        closewindow_propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='14'}" % contact_id
                        squish.clickButton(squish.findObject(closewindow_propstring))

                    except Exception as e:
                        test.log('Error closing \'Resume Call\' window')
                        status = ccc.STATUS_FAILURE


            if status == ccc.STATUS_SUCCESS:

                # Validation - Check that the current hold state matches the
                # targetHoldstate.
                # Move away from the button and wait two seconds to make sure the window reflects the updated status
                squish.mouseMove(10,10)
                squish.snooze(2)
                holdState = self.checkHoldState(contact_id)
                validState = (holdState == targetHoldstate)

                if operation_type == ccc.OT_HOLD:
                    if validState:
                        test.log('HoldResumeCall() method successfully held the call')
                    else:
                        test.log('HoldResumeCall() method failed to hold the call')
                        currentcall.held = False
                        status = ccc.STATUS_FAILURE

                elif operation_type == ccc.OT_RESUME:
                    if validState:
                        test.log('HoldResumeCall() method successfully resumed the call')
                    else:
                        test.log('HoldResumeCall() method failed to resume the call')
                        currentcall.held = False
                        status = ccc.STATUS_FAILURE

        return status


    def findAllDevices(self, contact_id):

        done = False
        occurrenceCount = 0
        device_itemlist =[]

        propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='8'}" % contact_id

        try:
            # First we need to trigger display of the "call control" button so we can check its label
            squish.mouseMove(10, 10)
            squish.snooze(1)
            callcontrolgraphic = squish.findObject(propstring)
            squish.mouseMove(callcontrolgraphic, callcontrolgraphic.width/2, callcontrolgraphic.height/2)
            squish.snooze(1)
        except Exception as e:
            test.log('Error locating and activating call control graphic')
            for item in e.args:
                test.log(str(item))
                status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                # Now go to the 'Devices' page
                try:
                    squish.mouseClick("{name='Devices' role='PageTab'}")
                    test.log('Clicked on \'Devices\' tab')
                except LookupError:
                    test.log('Error, cannot locate \'Devices\' tab')
                    status = ccc.STATUS_FAILURE

        while not done:
            occurrenceCount += 1
            propstring = "{type='ListItem' text~='^Headset|PC|Custom.*$' occurrence='%s'}" % str(occurrenceCount)

            try:
                thisItem = squish.findObject(propstring)
                device_itemlist.append(thisItem)
            except LookupError:
                done = True

        return device_itemlist


    def TransferCallAudio(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'TransferCallAudio'

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

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

            currentcall = self.getCall(contact_id)
            if not currentcall:
                test.log('Error, no valid call for %s, cannot transfer call audio')
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
            inUseIndex = 0
            itemCount = 0
            newSelectionIndex = 0
            newSelectedItem = None
            newSelectionText = ''

            itemList = self.findAllDevices(contact_id)

            if not itemList:
                test.log('Error obtaining list of connected devices')
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
                # Find item in list which is in use ('Selected')
                i = 0
                inUseIndex = 0
                for item in itemList:
                    # The list will be zero-indexed, though the occurrences begin with 1
                    if re.match('.*Selected', item.state):
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
                newSelectedItem = itemList[newSelectionIndex]

                if status == ccc.STATUS_SUCCESS:
                    try:
                        squish.mouseClick(newSelectedItem)
                    except LookupError:
                        test.log('Error selecting new audio device')
                        status == ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Validate - Look for the new 'Selected' index and compare to new selection index
                    testItemList = self.findAllDevices(contact_id)
                    i = 0
                    newInUseIndex = 0
                    for item in testItemList:
                        if re.match('.*Selected', item.state):
                            newInUseIndex = i
                        i += 1

                    if newInUseIndex == newSelectionIndex:
                        test.log('TransferCallAudio() successfully transferred audio to %s device' % optext)
                    else:
                        test.log('TransferCallAudio() failed to transfer audio to %s device' % optext)
                        status = ccc.STATUS_FAILURE

                    # Pop down the devices page by moving off the callcontrol graphic
                    squish.mouseMove(10, 10)

        return status


    def TransferCall(self, operation_type, newcontact, current_contact=None):
        status = ccc.STATUS_SUCCESS
        methodname = 'TransferCall'

        areacode = ''
        exchange = ''
        number = ''
        contact_listitemname = ''
        newcontactstring = ''

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

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
                if newcontact.isdigit():
                    newcontactstring = self.phonify(newcontact)
                else:
                    newcontactstring = newcontact

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_TRANSFER_CALL:

                propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='8'}" % current_contact

                try:

                    # First we need to trigger display of the "call control" button so we can check its label
                    squish.mouseMove(10, 10)
                    squish.snooze(1)
                    callcontrolgraphic = squish.findObject(propstring)
                    squish.mouseMove(callcontrolgraphic, callcontrolgraphic.width/2, callcontrolgraphic.height/2)
                    squish.snooze(1)
                except Exception as e:
                    test.log('Error locating and activating call control graphic')
                    for item in e.args:
                        test.log(str(item))
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Now go to the 'Transfer Call' page
                    try:
                        squish.mouseClick(squish.waitForObject("{name='Transfer call' role='PageTab'}"))
                    except LookupError:
                        test.log('Error, cannot locate \'Transfer call\' tab')
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    try:
                        squish.mouseClick(squish.waitForObject("{text='Another Person or Number' role='PushButton'}"))
                    except LookupError:
                        test.log('Error, cannot locate \'Another Person or Number\' pushbutton')
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    try:
                        contact_listitemname = "{text?='%s' type='ListItem'}" % newcontactstring
                        squish.mouseClick(squish.findObject(contact_listitemname))
                        squish.mouseClick(squish.waitForObject("{text='OK' role='PushButton'}"))
                    except Exception as e:
                        test.log('Error, cannot locate list item for contact %s' % newcontact)
                        for item in e.args:
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

            if status == ccc.STATUS_SUCCESS:

                # Validation: The connection window should now be absent
                # Give the UI a second to reflect transition
                squish.snooze(1)

                transfermatchtext = 'Transferring call to %s' % newcontactstring
                label_propstring = "{text?='%s' type='Label'}" % transfermatchtext
                test.log('label_propstring=%s' % label_propstring)

                try:
                    callstatuslabel = squish.waitForObject(label_propstring)
                    test.log('TransferCall() successfully transferred the call')
                except LookupError:
                    status = ccc.STATUS_FAILURE
                    test.log('TransferCall() failed to successfully transfer the call')
                '''
                if not self.connectionWindow(current_contact):
                    test.log('TransferCall() successfully transferred the call')
                else:
                    status = ccc.STATUS_FAILURE
                    test.log('TransferCall() failed to successfully transfer the call')
                '''
        return status


    def AdjustCallVolume(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'AdjustCallVolume'
        slider=None

        # Microsoft-determined; this value reflects reality, it doesn't define it. :-)
        bumpIncrement = 1

        bumps = 5

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

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
                currentcall = self.getCall(contact_id)
                if not currentcall:
                    test.log('Error, no valid call for %s, cannot adjust call volume')
                    status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='8'}" % contact_id
            current_vol = 0
            new_vol = 0

            try:
                # First we need to trigger display of the "call control" button so we can check its label
                squish.mouseMove(10, 10)
                squish.snooze(1)
                callcontrolgraphic = squish.findObject(propstring)
                squish.mouseMove(callcontrolgraphic, callcontrolgraphic.width/2, callcontrolgraphic.height/2)
                squish.snooze(1)
            except Exception as e:
                test.log('Error locating and activating call control graphic')
                for item in e.args:
                    test.log(str(item))
                    status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                try:
                    squish.mouseClick("{name='Devices' role='PageTab'}")
                    test.log('Clicked on \'Devices\' tab')
                except LookupError:
                    test.log('Error, cannot locate \'Devices\' tab')
                    status = ccc.STATUS_FAILURE

                try:
                    # We're finally at the goal: The volume slider.
                    slider = squish.findObject("{role='Slider'}")
                except LookupError:
                    test.log('Error, cannot locate volume slider')
                    status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                # Get the current volume and the measurements we need
                slider_height = __builtin__.int(slider.height)
                slider_width = __builtin__.int(slider.width)
                current_vol = __builtin__.int(slider.value)

                counter = 0
                try:
                    if operation_type == ccc.OT_VOLUME_UP_SPEAKER:
                        new_vol = current_vol + (bumpIncrement * bumps)
                        clickloc = slider_width - 5
                    elif operation_type == ccc.OT_VOLUME_DOWN_SPEAKER:
                        new_vol = current_vol - (bumpIncrement * bumps)
                        clickloc = 5

                    while counter < bumps:
                        # Click near the end towards which we're going. The increment will
                        # always be bumpIncrement.
                        squish.mouseClick(slider, clickloc, slider_height/2, squish.MouseButton.LeftButton)
                        counter += 1
                except Exception as e:
                    test.log('Error, could not adjust volume')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                # Validation consists solely of comparing the new volume to the intended value
                if __builtin__.int(slider.value) != new_vol:
                    test.log('Error, AdjustCallVolume() did not successfully adjust the call volume')
                    status = ccc.STATUS_FAILURE
                else:
                    test.log('AdjustCallVolume() successfully adjusted the call volume')


        return status


    # Note that this method does not recognize an implicit 'current contact,' but requires
    # specification of both contacts.
    def ConferenceCall(self, operation_type, contact_one, contact_two):

        methodname = 'ConferenceCall'

        super(Microsoft_lync_2013_CallControl, self).log_method_called(self.sp_id, methodname)

        status = ccc.STATUS_SUCCESS

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.CONFERENCE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            status = ccc.STATUS_SUCCESS
            contact_one_listitem = None
            contact_two_listitem = None

            # Make sure we're not hidden
            squish.setForegroundWindow(self.lync_window)

            if operation_type == ccc.OT_CONFERENCE_TWO_CONTACTS:
                # First, select conference contacts using Ctrl+left-button-select
                #

                try:
                    # Convert to 'lastname, firstname' if we have a last name
                    (contactone_firstname, contactone_lastname) = contact_one.split(' ')
                    contact_one = '%s, %s' % (lastname, firstname)
                    (contacttwo_firstname, contacttwo_lastname) = contact_one.split(' ')
                    contact_two = '%s, %s' % (contacttwo_lastname, contacttwo_firstname)
                except:
                    # If we fail, just carry on... no need to convert
                    pass

                try:
                    # Put the softphone in 'Contacts' mode, then select and conference the contacts
                    squish.clickButton(squish.findObject("{text='Contacts' role='PushButton' type='Button'}"))
                    listitem_one = squish.findObject("{text?='*%s*' role='ListItem'}" % contact_one)
                    listitem_two = squish.findObject("{text?='*%s*' role='ListItem'}" % contact_two)
                    squish.clickButton(listitem_one)
                    squish.mouseClick(listitem_two, squish.MouseButton.LeftButton, squish.Modifier.Control)
                    squish.mouseClick(listitem_one, squish.MouseButton.RightButton)
                    squish.snooze(1)
                    squish.clickButton(squish.waitForObject("{text='Start a Conference Call' type='MenuItem'}"))
                    squish.snooze(1)
                    squish.clickButton(squish.waitForObject("{text='Lync Call' type='MenuItem'}"))

                except Exception as e:
                    self.display_all_objects()
                    test.log('Error in PlaceCall() method')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Wait a second to make sure the window reflects the updated status
                    squish.snooze(1)

                    try:
                        conferencecall_label = squish.waitForObject("{text='Starting conference call...' type='Label'}")
                        test.log('ConferenceCall() successfully initiated a conference call')
                        self.addCall(contact_one, 'dialing', True)
                        self.addCall(contact_two, 'dialing', True)
                    except LookupError:
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
            call_status_label = squish.findObject("{text?='Calling *...' type='Label'}")
            directdialing = True
            test.log('Found directdialing calling window')
        except LookupError:
            pass

        conferencing = self.checkInConferenceState(self.currentcall_contact)

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

        conn_window = None

        testcontact = contact.strip('+')

        # Once the connection has been made, numeric contacts are displayed raw, not in
        # "phone format," and may have a '9' prepended to them for outside calls

        try:
            conn_window = squish.findObject("{type='Window' text?='*%s'}" % contact)
        except LookupError:
            # Not finding the connection window is a perfectly valid result
            pass

        return conn_window


    def checkMuteState(self, contact):

        status = ccc.STATUS_SUCCESS
        muteState = False

        # With one call, the graphic which represents the mute/unmute function is occurrence '8'.
        # As multiple calls accumulate, the most recent call is the one which will have this
        # occurrence index, and earlier calls will increment this index by one successively.
        # This is a royal pain, and I hope we can identify this button more uniquely moving forward.
        propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='8'}" % contact

        try:
            # First we need to trigger display of the Mute/Unmute button so we can check its label
            squish.mouseMove(10, 10)
            squish.snooze(1)
            callcontrolgraphic = squish.findObject(propstring)
            squish.mouseMove(callcontrolgraphic, callcontrolgraphic.width/2, callcontrolgraphic.height/2)
            squish.snooze(1)
        except Exception as e:
            test.log('Error locating and activating mute/unmute graphic')
            for item in e.args:
                test.log(str(item))
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            try:
                # Now we can check the label and return state accordingly
                dummy = squish.findObject("{text='Unmute' role='PushButton'}")
                squish.setForegroundWindow(dummy)
                muteState = True
            except LookupError:
                pass

        # Move away to let the popup disappear
        squish.mouseMove(10,10)

        return muteState


    def checkHoldState(self, contact):

        status = ccc.STATUS_SUCCESS
        holdState = False

        # With one call, the graphic which represents the mute/unmute function is occurrence '8'.
        # As multiple calls accumulate, the most recent call is the one which will have this
        # occurrence index, and earlier calls will increment this index by one successively.
        # This is a royal pain, and I hope we can identify this button more uniquely moving forward.
        propstring = "{container={type='Window' text?='*%s'} role='Graphic' occurrence='8'}" % contact

        try:
            # First we need to trigger display of the Mute/Unmute button so we can check its label
            squish.mouseMove(10, 10)
            squish.snooze(1)
            callcontrolgraphic = squish.findObject(propstring)
            squish.mouseMove(callcontrolgraphic, callcontrolgraphic.width/2, callcontrolgraphic.height/2)
            squish.snooze(1)
        except Exception as e:
            test.log('Error locating and activating mute/unmute graphic')
            for item in e.args:
                test.log(str(item))
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            try:
                # Now we can check the label and return state accordingly
                dummy = squish.findObject("{text='Resume Call' role='PushButton'}")
                squish.setForegroundWindow(dummy)
                holdState = True
            except LookupError:
                pass

        # Move away to let the popup disappear
        squish.mouseMove(10,10)

        return holdState


    def checkInConferenceState(self, contact):

        self.display_all_objects()
        return False

        titlebar = self.connectionWindow(contact)
        if titlebar:
            squish.setForegroundWindow(titlebar)

        contact_propstring = "{text?='%s*In a conference call' type='ListItem'}" % contact
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
                        if self.checkInConferenceState(contact):
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

        for statestring in ccc.userstate_constants:
            try:
                propstring = "{role='StaticText' text='%s' type='Label'}" % statestring
                statebutton = squish.findObject(propstring)
                stateconstant = ccc.userstate_constants[statestring]
            except:
                pass

        if stateconstant == ccc.USERSTATE_UNDEFINED:
            test.log('Error, could not locate user state button')

        return stateconstant


    def SetSoftphoneToIdle(self):

        # Make sure we're not hidden
        squish.setForegroundWindow(self.lync_window)

        for call in self.calls:
            self.removeCall(call.contact_id)

        self.incomingcall_contact = ''
        self.currentcall_contact = ''

        # TBD: Abort any calls that are dialing and decline any calls that are ringing
        # (May implement by logging out and logging back in)

