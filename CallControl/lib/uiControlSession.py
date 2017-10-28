#-------------------------------------------------------------------------------
# Name:         callControlSession
# Purpose:      This module defines the class CallControlSession, used to
#               represent the attributes and methods associated with a
#               CallControl testing session.
#
# Author:      rashforth
#
# Created:     23/10/2012
# Copyright:   (c) Plantronics 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from os import path

import sys
import time

import pdctest.CallControl.lib.serverControl as sc
import pdctest.CallControl.lib.call_control_constants as ccc


class UIControlSession(object):


    def __init__(self, platform, sp_id='', mp_id=''):

        if sp_id not in ccc.sp_dict:
            print 'Error, invalid softphone id \'%s\'entered' % sp_id
            # Raise exception?

        if mp_id and mp_id not in ccc.mp_dict:
            print 'Error, invalid mediaplayer id \'%s\'entered' % mp_id
            # Raise exception?

        if not sp_id and not mp_id:
            print 'Error, must specify at least one of softphone or mediaplayer'
            # Raise exception?

        if platform not in ccc.sp_platforms:
            print 'Error, invalid platform \'%s\' entered' % platform
            # Raise exception?

        self.sp_id = sp_id
        self.mp_id = mp_id
        self.platform = platform

        # Squish server, needed for any UI control functionality
        self.uicontrol_server = None

        # squish-monitored softphone app and sp-specific 'virtual softphone'
        self.sp_autcontrol_server = None
        self.sp_callcontrol_server = None

        # squish-monitored mediaplayer app and sp-specific 'virtual mediaplayer'
        self.mp_autcontrol_server = None
        self.mp_callcontrol_server = None

        self.uicontrol_started = False

        self.sp_autcontrol_started = False
        self.sp_callcontrol_started = False

        self.mp_autcontrol_started = False
        self.mp_callcontrol_started = False


    def start(self):

        # First create the ServerControl instances, then start the servers
        #=================================================================

        # The UI Control server is squishserver
        self.uicontrol_server = sc.ServerControl('squishserver --verbose')
        testsuite_dir = 'C:\\squishdata\\suite_servers_%s' % self.platform

        # Now the softphone-specific setup if needed
        if self.sp_id:
            sp_aut_path = ccc.get_sp_aut_path(self.sp_id)
            sp_squishport = ccc.get_sp_squishport(self.sp_id)

            # Command to start the AUT depends on which one it is. At some point
            # there is always the possibility that the UI control app will not be
            # squish, but while it is this is the startwinaut command, which runs
            # the softphone app under squish control.
            sp_aut_cwd, sp_aut_exename = path.split(sp_aut_path)
            sp_aut_cmdstring = 'startwinaut --port=%d --cwd=\"%s\" \"%s\"' % (sp_squishport, sp_aut_cwd, sp_aut_path)

            # Strip off the trailing '.exe'
            sp_aut_procname = (sp_aut_exename.split('.'))[0]

            self.sp_autcontrol_server = sc.ServerControl(sp_aut_cmdstring, sp_aut_procname)

            sp_testcase_dir = '%s\\tst_%s' % (testsuite_dir, self.sp_id)
            sp_callcontrol_cmdstring = 'squishrunner --testsuite %s --testcase %s' % (testsuite_dir, sp_testcase_dir)
            sp_callcontrol_procname = 'callcontrolserver_%s' % self.sp_id
            self.sp_callcontrol_server = sc.ServerControl(sp_callcontrol_cmdstring, sp_callcontrol_procname)

        # Now the mediaplayer-specific setup if needed
        if self.mp_id:
            mp_aut_path = ccc.get_mp_aut_path(self.mp_id)
            mp_squishport = ccc.get_mp_squishport(self.mp_id)

            # Command to start the AUT depends on which one it is. At some point
            # there is always the possibility that the UI control app will not be
            # squish, but while it is this is the startwinaut command, which runs
            # the softphone app under squish control.
            mp_aut_cwd, mp_aut_exename = path.split(mp_aut_path)
            mp_aut_cmdstring = 'startwinaut --port=%d --cwd=\"%s\" \"%s\"' % (mp_squishport, mp_aut_cwd, mp_aut_path)

            # Strip off the trailing '.exe'
            mp_aut_procname = (mp_aut_exename.split('.'))[0]

            self.mp_autcontrol_server = sc.ServerControl(mp_aut_cmdstring, mp_aut_procname)

            mp_testcase_dir = '%s\\tst_%s' % (testsuite_dir, self.mp_id)
            mp_callcontrol_cmdstring = 'squishrunner --testsuite %s --testcase %s' % (testsuite_dir, mp_testcase_dir)
            mp_callcontrol_procname = 'callcontrolserver_%s' % self.mp_id
            self.mp_callcontrol_server = sc.ServerControl(mp_callcontrol_cmdstring, mp_callcontrol_procname)

        # Now we're ready for the actual startup
        #=======================================
        status = self.uicontrol_server.start()
        if status == ccc.STATUS_SUCCESS:
            print 'uicontrol_server started'
            self.uicontrol_started = True
            time.sleep(2)
        else:
            print 'Error starting uicontrol_server'
            status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
            if self.sp_id:
                status = self.sp_autcontrol_server.start()
                if status == ccc.STATUS_SUCCESS:
                    print 'sp_autcontrol_server started'
                    self.sp_autcontrol_started = True
                else:
                    print 'Error starting sp_autcontrol_server'
                    status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
            if self.mp_id:
                status = self.mp_autcontrol_server.start()
                if status == ccc.STATUS_SUCCESS:
                    print 'mp_autcontrol_server started'
                    self.mp_autcontrol_started = True
                else:
                    print 'Error starting sp_autcontrol_server'
                    status = ccc.STATUS_FAILURE

        if status == ccc.STATUS_SUCCESS:
            if self.sp_id:
                # We need to leave plenty of time for the AUT to start up...
                time.sleep(13)
                status = self.sp_callcontrol_server.start()
                if status == ccc.STATUS_SUCCESS:
                    print 'sp_callcontrol_server started'
                    self.sp_callcontrol_started = True
                else:
                    print 'Error starting sp_autcontrol_server'

        if status == ccc.STATUS_SUCCESS:
            if self.mp_id:
                # We need to leave plenty of time for Lync to start up...
                time.sleep(11)
                status = self.mp_callcontrol_server.start()
                if status == ccc.STATUS_SUCCESS:
                    print 'mp_callcontrol_server started'
                    self.mp_callcontrol_started = True
                else:
                    print 'Error starting sp_autcontrol_server'

        if status is not ccc.STATUS_SUCCESS:
            # If we did not start all servers successfully, clean up everything
            # that did succeed
            self.stop()

        return status


    def stop(self):

        if self.uicontrol_server is not None and self.uicontrol_started:
            self.uicontrol_server.stop()

        if self.sp_autcontrol_server is not None and self.sp_autcontrol_started:
            self.sp_autcontrol_server.stop()

        if self.mp_autcontrol_server is not None and self.mp_autcontrol_started:
            self.mp_autcontrol_server.stop()

        if self.sp_callcontrol_server is not None and self.sp_callcontrol_started:
            self.sp_callcontrol_server.stop()

        if self.mp_callcontrol_server is not None and self.mp_callcontrol_started:
            self.mp_callcontrol_server.stop()

        return ccc.STATUS_SUCCESS


