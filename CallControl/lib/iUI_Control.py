#-------------------------------------------------------------------------------
# Name:        iUI_Control
# Purpose:     This is the core class for the all Plantronics modules used to
#              drive specific types of software applications, such as softphones
#              and media players, within the TDTF (Target Device Test Framework).
#
# Author:      rashforth
#
# Created:     12/25/2012
# Copyright:   (c) Plantronics 2012
#-------------------------------------------------------------------------------
import test
import testData
import object
import objectMap
import squishinfo
import squish

import exceptions

import pdctest.CallControl.lib.call_control_constants as ccc

import __builtin__

class UI_Control(__builtin__.object):


    def log_method_called(self, app_id, app_method):
        if self.debug:
            test.log( 'Invoked method \'%s()\' of app \'%s\'' % (app_method, app_id))


    def isvalid_func_op(self, function, operation):
        if operation in ccc.valid_func_op_dict[function]:
            return True
        else:
            return False


    # The following utility methods can be invoked from derived modules to
    # discover all UI objects on the screen at the time of invocation
    def display_object_properties(self, obj):

        status = ccc.STATUS_SUCCESS

        if not obj:
            return (ccc.STATUS_SUCCESS, None)

        try:
            properties = object.properties(obj)
        except Exception as e:
            test.log('Error, cannot get properties for object')
            return (ccc.STATUS_FAILURE, None)

        # A value of 'None' is okay, at least I think so currently ;-)
        if properties:

            test.log('+++++++++++++++++++++++++++++++++++++++++++++++++++++++')
            for prop, value in properties.iteritems():
                test.log( 'property=%s, value=%s' % (prop, str(value)))

        return (status)


    def display_all_properties(self, obj):

        status = ccc.STATUS_SUCCESS

        if status == ccc.STATUS_SUCCESS:
            # Display this object's properties
            status = self.display_object_properties(obj)
            kids = None

            # Iterate over children
            try:
                kids = object.children(obj)
            except Exception as e:
                test.log('Error in getting object children')

            if kids:
                for kid in kids:
                    status = self.display_all_properties(kid)

        return status

    def display_all_objects(self):

        status = ccc.STATUS_SUCCESS

        toplevel_objects = []

        try:
            toplevel_objects = object.topLevelObjects()

        except Exception as e:
            test.log('Error, cannot get application top level objects')
            status = ccc.STATUS_FAILURE

        if toplevel_objects and status == ccc.STATUS_SUCCESS:
            test.log('Displaying properties for all UI objects')
            test.log('========================================')
            for obj in toplevel_objects:
                status = self.display_all_properties(obj)
                if status != ccc.STATUS_SUCCESS:
                    break

        return status


    def display_all_occurrences(self, propstring, startindex=1):

        status = ccc.STATUS_SUCCESS

        # Loop through all occurrence instances, substituting the current index
        # into the propstring passed in
        occurrence_index = startindex
        done = False

        while not done:
            try:
                object_props = "{%s occurrence='%d'}" % (propstring, occurrence_index)
                object = squish.findObject(object_props)
                test.log('Displaying occurrence %d for propstring %s' % (occurrence_index, propstring))
                self.display_object_properties(object)
                squish.mouseMove(object, object.width/2, object.height/2)
                squish.snooze(3)
                occurrence_index += 1
            except LookupError:
                test.log('Completed display of all occurrences for propstring %s' % propstring)
                done = True

        return status

