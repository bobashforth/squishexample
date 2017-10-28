#-------------------------------------------------------------------------------
# Name:        skype_CallControl
# Purpose:     This is the Skype-specific iUI_CallControl interface
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

import os.path
import sys
import re
import exceptions
import pdctest.CallControl.lib.call_control_constants as ccc
import pdctest.CallControl.lib.iUI_CallControl as uicc


class Skype_CallControl(uicc.UI_CallControl):


    def __init__(self, sp_id):

        # First call the __init__() method of the base class
        status = super(Skype_CallControl, self).__init__(sp_id)

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
        self.incoming_call_regex = re.compile(r'(.*) calling')


    def PlaceCall(self, operation_type, contact_id):
        status = ccc.STATUS_SUCCESS
        methodname = 'PlaceCall'

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.PLACE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_DRAGDROP_TO_PHONE_MENU, ccc.OT_DRAGDROP_TO_CONVERSATION_WND]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_CONTACT_NAME:
                # Bring up the context menu over the selected contact name
                try:
                    contact_listitem = squish.waitForObject("{text='%s' type='ListItem'}" % contact_id)
                    squish.mouseClick(contact_listitem, 102, 15, squish.MouseButton.RightButton)
                except Exception as e:
                    test.log('Error clicking on contact name \'%s\' list item' % contact_id)
                    status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    # Click the 'Call' menu selection
                    try:
                        squish.mouseClick(squish.waitForObject("{text='Call' type='MenuItem' defaultAction='Execute'}"))
                    except Exception as e:
                        test.log('Error selecting \'Call\' menu item')
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Wait one second for the UI to react to the action above
                    squish.snooze(1)

                    # Validate by checking for the presence of the 'Ringing' pane
                    try:
                        propstring = "{role='Pane' type='AccessibleObject' name='%s - Ringing'}" % contact_id
                        callingpane = squish.findObject(propstring)
                        test.log('PlaceCall() successfully placed a call to %s' % contact_id)
                        self.addCall(contact_id, 'dialing')
                    except LookupError:
                        test.log('PlaceCall() failed to place a call to \'%s\'' % contact_id)
                        status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_PHONE_NUMBER_MAIN_WND:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED
            elif operation_type == ccc.OT_PHONE_NUMBER_DIALPAD:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED
            elif operation_type == ccc.OT_EXTENSION_NUMBER_MAIN_WND:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED
            elif operation_type == ccc.OT_EXTENSION_NUMBER_DIALPAD:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED
            elif operation_type == ccc.OT_START_CONFERENCE_CALL:
                status = ccc.STATUS_FEATURE_NOT_IMPLEMENTED
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status


    def getIncomingCallContact(self):

        incoming_contact = ''
        try:
            call_window = squish.waitForObject("{text~='.* calling' type='Label'}")

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


    # Internal method defined to reuse functionality
    def _navigateToCallSettingsDialog(self):

        status = ccc.STATUS_SUCCESS

        try:
            # Get to the relevant settings window
            squish.mouseClick(squish.waitForObject("{text='Tools' type='MenuItem'}"))
            squish.mouseClick(squish.waitForObject("{text='Options...' type='MenuItem'}"))
            squish.mouseClick(squish.waitForObject("{name='Calls' role='OutlineItem'}"), 46, 17, squish.MouseButton.LeftButton)
        except Exception as e:
            test.log('Error in navigating to Call settings dialog')
            for item in e.args:
                test.log(str(item))
                status = ccc.STATUS_FAILURE

        return status


    def setAutoAnswer(self, optype):

        status = ccc.STATUS_SUCCESS
        methodname = 'setAutoAnswer'

        autoanswer_checkbox = None
        autoanswer_state = 0
        checkstate_needed = 0

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op( ccc.SET_AUTO_ANSWER, optype):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            status = self._navigateToCallSettingsDialog()

            if status == ccc.STATUS_SUCCESS:

                # Here we check which operation was requested against the current checkbox state and react accordingly
                if optype == ccc.OT_AUTO_ANSWER_ON:
                    checkstate_needed = 1
                elif optype == ccc.OT_AUTO_ANSWER_OFF:
                    checkstate_needed = 0

                try:
                    autoanswer_checkbox = squish.waitForObject("{text='Answer incoming calls automatically' type='CheckBox'}")

                    autoanswer_state = autoanswer_checkbox.checked
                    test.log('autoanswer_state=%d, checkstate_needed=%d, autoanswer_checkbox.checked=%d' % \
                        (autoanswer_state, checkstate_needed, autoanswer_checkbox.checked))

                except Exception as e:
                    test.log('Error retrieving autoanswer checkbox state')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                try:
                    if autoanswer_state != checkstate_needed:
                        squish.mouseClick(squish.waitForObject(autoanswer_checkbox))
                        squish.clickButton(squish.waitForObject("{text='Save' type='Button'}"))
                    else:
                        test.log('AutoAnswer is already set to the requested state, no action taken')
                        squish.clickButton(squish.waitForObject("{text='Cancel' type='Button'}"))

                except Exception as e:
                    test.log('Error in setting requested AutoAnswer state')
                    for item in e.args:
                        test.log(str(item))
                        status = ccc.STATUS_FAILURE

            if status == ccc.STATUS_SUCCESS:

                # Validation: Compare current checkbox state

                # If this fails the called method logs the problem, just leave failure status
                status = self._navigateToCallSettingsDialog()
                if status == ccc.STATUS_SUCCESS:
                    try:
                        autoanswer_checkbox = squish.waitForObject("{text='Answer incoming calls automatically' type='CheckBox'}")
                    except LookupError:
                        test.log('Error locating AutoAnswer checkbox')
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    autoanswer_state = autoanswer_checkbox.checked
                    test.log('Validation: autoanswer_state=%d, checkstate_needed=%d, autoanswer_checkbox.checked=%d' % \
                        (autoanswer_state, checkstate_needed, autoanswer_checkbox.checked))
                    if autoanswer_state == checkstate_needed:
                        test.log('setAutoAnswer() successfully set the requested AutoAnswer state')
                    else:
                        test.log('setAutoAnswer() failed to set the requested AutoAnswer state')
                        status = ccc.STATUS_FAILURE

                # We need to exit the settings dialog regardless of success or failure above
                try:
                    squish.clickButton(squish.waitForObject("{text='Cancel' type='Button'}"))
                except Exception as e:
                    test.log('Error in locating Cancel button in settings dialog')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

        return status


    def AnswerCall(self, operation_type):

        status = ccc.STATUS_SUCCESS
        methodname = 'AnswerCall'
        propstring = ''

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op( ccc.ANSWER_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_PICKUP]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if not self.isRinging():
                test.log('Error, no incoming call to answer')
            else:
                if operation_type == ccc.OT_TOAST:


                        try:
                            call_answer_button = squish.waitForObject("{text='Answer' type='Button'}")
                            squish.clickButton(call_answer_button)
                        except Exception as e:
                            test.log('Error in AnswerCall() method')
                            for item in e.args:
                                test.log(str(item))

                            status = ccc.STATUS_FAILURE

                elif operation_type == ccc.OT_MAIN_WND:
                    try:
                        propstring = "{text='%s' type='ListItem'}" % self.incomingcall_contact
                        squish.clickButton(squish.waitForObject(propstring))
                    except LookupError:
                        test.log('Error locating %s list item' % incomingcall_contact)
                        status = ccc.STATUS_FAILURE

                    if status == ccc.STATUS_SUCCESS:
                        try:
                            squish.clickButton(squish.waitForObject("{text='Call' type='MenuItem'}"))
                        except LookupError:
                            test.log('Error locating Call menu item')
                            status = ccc.STATUS_FAILURE

                    if status == ccc.STATUS_SUCCESS:
                        try:
                            squish.clickButton(squish.waitForObject("{text='Answer' type='MenuItem'}"))
                        except LookupError:
                            test.log('Error locating Answer menu item')
                            status = ccc.STATUS_FAILURE

                else:
                    # Should never get here, given the above checks... being safe
                    status = ccc.STATUS_INVALID_OPERATION_TYPE
        if status == ccc.STATUS_SUCCESS:

            # Wait a second for the UI to adjust
            squish.snooze(1)

            # Validate by locating hangup button
            if self.isConnected(self.incomingcall_contact):
                test.log('AnswerCall() successfully answered a call')
                self.addCall(self.incomingcall_contact, 'connected')
            else:
                test.log('AnswerCall() did not successfully answer a call')
                status = ccc.STATUS_FAILURE

        return status


    def DeclineCall(self, operation_type):

        status = ccc.STATUS_SUCCESS
        methodname = 'DeclineCall'

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.DECLINE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        elif operation_type in [ccc.OT_MAIN_WND]:
            # Note: Skype has an 'Ignore' menu entry, but that does not
            # end the call, as 'Decline' does, it only ignores it.
            status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if not self.isRinging():
                test.log('Error, no incoming call to decline')
            else:
                if operation_type == ccc.OT_TOAST:

                        try:
                            call_decline_button = squish.waitForObject("{text='Decline' type='Button'}")
                            squish.clickButton(call_decline_button)
                        except Exception as e:
                            test.log('Error in DeclineCall() method')
                            for item in e.args:
                                test.log(str(item))

                            status = ccc.STATUS_FAILURE

                else:
                    # Should never get here, given the above checks... being safe
                    status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:
            # Validate by confirming that the decline button can no longer be located
            try:
                squish.snooze(1)
                call_decline_button = squish.findObject("{text='Decline' type='Button'}")
                test.log('DeclineCall() failed to decline a call')
            except LookupError:
                test.log('DeclineCall() successfully declined a call')

        return status


    def EndCall(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'EndCall'

        if not contact_id:
            contact_id = self.currentcall_contact
        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.END_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            if operation_type == ccc.OT_TOAST:
                try:
                    endcallbutton = squish.waitForObject("{text='End call' role='PushButton' type='Button' occurrence='2'}")
                    squish.clickButton(endcallbutton)
                except LookupError:
                    test.log('Error activating toast End call button')
                    status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_MAIN_WND:
                try:
                    propstring = "{text='%s' type='ListItem'}" % self.currentcall_contact
                    listitem = squish.waitForObject(propstring)
                    squish.clickButton(listitem)
                    self.display_all_properties(listitem)
                    endcallbutton = squish.waitForObject("{text='End call' role='PushButton' type='Button'}")
                    squish.clickButton(endcallbutton)
                    squish.snooze(1)
                    self.display_all_properties(listitem)
                except LookupError:
                    test.log('Error activating End call button')
                    status = ccc.STATUS_FAILURE

                pass
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

            if status == ccc.STATUS_SUCCESS:
                # Validate by confirming failure to find End call button
                try:
                    squish.snooze(1)
                    endcallbutton = squish.findObject("{text='End call' role='PushButton' type='Button'}")
                    test.log('EndCall() failed to end a call')
                    status = ccc.STATUS_FAILURE
                except LookupError:
                    test.log('EndCall() successfully ended a call')
                    self.removeCall(contact_id)

        return status


    def SetCallMuting(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'SetCallMuting'

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        if not contact_id:
            contact_id = self.currentcall_contact

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.SET_CALL_MUTING, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            if operation_type == ccc.OT_MUTE:
                try:
                    squish.clickButton(squish.waitForObject("{text='%s' type='ListItem'}" % contact_id))
                    mutebutton = squish.findObject("{text='Mute' type='Button'}")
                    squish.clickButton(mutebutton)
                except LookupError:
                    test.log('Error locating Mute button')
                    status = ccc.STATUS_FAILURE
                except Exception as e:
                    test.log('Error activating Mute button')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Wait a second to let the UI reflect the change
                    squish.snooze(1)

                    # Validation is toggling to the desired state
                    try:
                        unmutebutton = squish.findObject("{text='Unmute' type='Button'}")
                        test.log('SetCallMuting() successfully muted the call.')
                    except LookupError:
                        test.log('SetCallMuting() failed to mute the call.')
                        status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_UNMUTE:
                try:
                    unmutebutton = squish.findObject("{text='Unmute' type='Button'}")
                    squish.clickButton(unmutebutton)
                except LookupError:
                    test.log('Error locating Unmute button')
                    status = ccc.STATUS_FAILURE
                except Exception as e:
                    test.log('Error activating Unmute button')
                    for item in e.args:
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Wait a second to let the UI reflect the change
                    squish.snooze(1)

                    # Validation is toggling to the desired state
                    try:
                        mutebutton = squish.findObject("{text='Mute' type='Button'}")
                        test.log('SetCallMuting() successfully unmuted the call.')
                    except LookupError:
                        test.log('SetCallMuting() failed to unmute the call.')
                        status = ccc.STATUS_FAILURE

        return status


    def HoldResumeCall(self, operation_type, contact_id=''):
        status = ccc.STATUS_SUCCESS
        methodname = 'HoldResumeCall'

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        if not contact_id:
            contact_id = self.currentcall_contact

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.HOLD_RESUME_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE

        if status == ccc.STATUS_SUCCESS:

            if operation_type == ccc.OT_HOLD:

                try:
                    squish.clickButton(squish.waitForObject("{text='%s' type='ListItem'}" % contact_id))
                    squish.clickButton(squish.waitForObject("{text='Call' type='MenuItem'}"))
                    squish.clickButton(squish.waitForObject("{text='Hold' type='MenuItem'}"))
                except:
                    test.log('Error activating Hold menu item')
                    status == ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Wait a second to let the UI reflect the change
                    squish.snooze(1)

                    if self.checkHoldState(contact_id):
                        test.log('HoldResumeCall() successfully put a call on hold')
                    else:
                        test.log('HoldResumeCall() failed to put a call on hold')

            elif operation_type == ccc.OT_RESUME:

                try:
                    squish.clickButton(squish.waitForObject("{text='Call' type='MenuItem'}"))
                    squish.clickButton(squish.waitForObject("{text='Resume' type='MenuItem'}"))
                except:
                    test.log('Error activating Resume menu item')
                    status == ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    # Wait a second to let the UI reflect the change
                    squish.snooze(1)

                    if not self.checkHoldState(contact_id):
                        test.log('HoldResumeCall() successfully resumed a call')
                    else:
                        test.log('HoldResumeCall() failed to resume a call')

        return status


    def TransferCall(self, operation_type, newcontact, current_contact=None):

        status = ccc.STATUS_SUCCESS
        methodname = 'TransferCall'

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.TRANSFER_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if not current_contact:
                current_contact = self.currentcall_contact
            currentcall = self.getCall(current_contact)
            if not currentcall:
                test.log('Error, no valid call for %s, cannot transfer call')
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:

            if newcontact is None:
                test.log('No contact number specified, TransferCall() failed.')
                status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_TRANSFER_CALL:

                try:
                    squish.clickButton(squish.waitForObject("{text='%s' type='ListItem'}" % current_contact))
                    squish.mouseClick(squish.waitForObject("{text='Call' type='MenuItem'}"))
                    squish.mouseClick(squish.waitForObject("{text='Transfer' type='MenuItem'}"))
                    squish.mouseClick(squish.waitForObject("{text='%s' type='MenuItem'}" % newcontact))
                    squish.mouseClick(squish.waitForObject("{text='Transfer' type='Button'}"))

                except Exception as e:
                    test.log('Error activating Transfer menu selection')
                    for item in e.args():
                        test.log(str(item))
                    status = ccc.STATUS_FAILURE


                if status == ccc.STATUS_SUCCESS:

                    # Wait a second to let the UI reflect the change
                    squish.snooze(1)

                    # Validation - Check for 'Finished' status label
                    try:
                        squish.findObject("{name='Finished' role='StaticText' type='Label'}")
                        test.log('TransferCall() successfully transferred a call')
                    except LookupError:
                        test.log('TransferCall() failed to transfer a call')
                        status = ccc.STATUS_FAILURE

        return status


    # Note that this method does not recognize an implicit 'current contact,' but requires
    # specification of both contacts.
    def ConferenceCall(self, operation_type, contact_one, contact_two):

        methodname = 'ConferenceCall'

        super(Skype_CallControl, self).log_method_called(self.sp_id, methodname)

        status = ccc.STATUS_SUCCESS

        # Validate operation type for function and for this softphone
        if not self.isvalid_func_op(ccc.CONFERENCE_CALL, operation_type):
            status = ccc.STATUS_INVALID_OPERATION_TYPE
        else:
            if operation_type in [ccc.OT_CONFERENCE_DRAGDROP]:
                status = ccc.STATUS_FEATURE_NOT_SUPPORTED

        if status == ccc.STATUS_SUCCESS:
            if operation_type == ccc.OT_CONFERENCE_TWO_CONTACTS:

                try:
                    squish.mouseClick(squish.waitForObject("{text='%s' type='ListItem'}" % contact_one), \
                        130, 12, squish.MouseButton.LeftButton)
                except LookupError:
                    test.log('Error selecting contact_one list item')
                    status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    try:
                        squish.mouseClick(squish.waitForObject("{text='%s' type='ListItem'}" % contact_two), \
                            130, 12, squish.MouseButton.LeftButton, squish.KeyboardModifier.Control)
                    except LookupError:
                        test.log('Error selecting contact_two list item')
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:
                    try:
                        squish.mouseClick(squish.waitForObject("{text='%s' type='ListItem'}" % contact_two), \
                            squish.MouseButton.RightButton, squish.KeyboardModifier.None)
                        squish.mouseClick(squish.waitForObject("{text='Call the Group' type='MenuItem'}"))

                    except Exception as e:
                        test.log('Error activating \'Call the Group\' menu item')
                        for item in e.args():
                            test.log(str(item))
                        status = ccc.STATUS_FAILURE

                if status == ccc.STATUS_SUCCESS:

                    #Validation - Look for the window with both conference call recipients listed in the object properties
                    try:
                        squish.waitForObject("{name='%s, %s' type='Window'}" % (contact_two, contact_one))
                        test.log('ConferenceCall() successfully placed a call to two contacts')
                        thiscall = self.addCall(contact_one, 'dialing')
                        thiscall.inconference = True

                    except LookupError:
                        test.log('ConferenceCall() failed to place a call to two contacts')
                        status = ccc.STATUS_FAILURE

            elif operation_type == ccc.OT_CONFERENCE_TWO_CALLS:
                pass
            else:
                # Should never get here, given the above checks... being safe
                status = ccc.STATUS_INVALID_OPERATION_TYPE

        return status


    def SetSoftphoneToIdle(self):

        # Make sure we're not hidden
        squish.setForegroundWindow("{text='Microsoft Lync          ' type='Window'}")

        for call in self.calls:
            self.removeCall(call.contact_id)

        self.incomingcall_contact = ''
        self.currentcall_contact = ''

        # TBD: Abort any calls that are dialing and decline any calls that are ringing
        # (May implement by logging out and logging back in)


    def isDialing(self):

        # Check for the presence of the 'calling' pane for the current contact
        contact = self.currentcall_contact
        try:
            propstring = "{role='Pane' type='AccessibleObject' name='%s - Ringing'}" % contact
            callingpane = squish.findObject(propstring)
            return True

        except LookupError:
            return False


    def isRinging(self):

        self.incomingcall_contact = self.getIncomingCallContact()
        if self.incomingcall_contact is None:
            return False
        else:
            self.addCall(self.incomingcall_contact, 'ringing')
            return True


    def isConnected(self, contact):
            # Contact is connected if we can locate hangup button
            try:
                propstring = "{text='%s' type='ListItem'}" % contact
                squish.clickButton(squish.findObject(propstring))
                endcallbutton = squish.findObject("{text='End call' type='Button'}")
                return True
            except LookupError:
                return False

    def connectionWindow(self, contact):

        titlebar = None

        try:
            titlebar = squish.findObject("{type='TitleBar' value='%s'}" % contact)
        except LookupError:
            pass

        return titlebar


    def refreshCallStatus(self):
        status = ccc.STATUS_SUCCESS

        if self.callstack:
            # Confirm status of each listed call
            # Use a copy of the list to iterate, so we can delete items safely
            listcopy = self.callstack[:]
            for contact in listcopy:
                test.log('refreshCallStatus, contact = %s' % contact)
                call = self.calls[contact]

                # Direct calls and conference calls have different windows and text indicating
                # 'dialing' and 'connected' states
                if call.status == 'dialing' or call.status is 'ringing':
                    # This is our last known state for this contact,
                    # check its current status
                    if self.connectionWindow(contact):
                        test.log('Setting call status for %s to connected' % contact)
                        call.status = 'connected'

                elif call.status == 'connected':
                    if not self.connectionWindow(contact):
                        test.log('Removing call for contact %s because it is not connected' % contact)
                        self.removeCall(contact)

                # For conference calls, we should check to see if we're the only one in the 'conference'
                #if call.isinconference and call.status == 'connected':
                #    if not self.connectionWindow(contact):
                #        self.removeCall(contact)

        return status


    def hasCalls(self):

        self.refreshCallStatus()
        if self.callstack:
            return True
        else:
            return False


    # Since some softphones take a second string argument, accept and ignore it
    def checkInConferenceState(self, contact, dummy):

        titlebar = self.connectionWindow(contact)
        if titlebar:
            squish.setForegroundWindow(titlebar)

        try:
            squish.findObject("{name?='?*, %s?*' type='Window'}" % contact)
            return True
        except LookupError:
            return False


    def checkMuteState(self, contact):

        muteState = None

        titlebar = self.connectionWindow(contact)
        if titlebar:
            squish.setForegroundWindow(titlebar)

        try:
            dummy = squish.findObject("{text='Mute' type='Button'}")
            muteState = False
        except:
            muteState = True

        return muteState


    def checkHoldState(self, contact):

        # Note that this method is not specific to the 'contact' argument, at least not yet.

        holdState = None

        try:
            dummy = squish.findObject("{name='CallOnHoldWidget' type='AccessibleObject'}")
            holdState = True

        except LookupError:
            holdState = False

        return holdState

