#-------------------------------------------------------------------------------
# Name:         callControlAdminService
# Purpose:      This module defines the class CallControlAdminService, used to
#               start the CallControl 'server service.' This service will be
#               started from run_tcm, and will start and stop all required
#               CallControl servers on request.
#
# Author:      rashforth
#
# Created:     28/10/2012
# Copyright:   (c) Plantronics 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

from os import path
import sys
import time

import pdctest.CallControl.lib.serverControl as sc
import pdctest.CallControl.lib.call_control_constants as ccc


class CallControlAdminService(object):


    def __init__(self):

        self.admin_server = None
        self.admin_server_started = False

    def start(self):

        self.admin_server = sc.ServerControl('python C:\\Python27\\Scripts\\run_callcontrol.py', 'callControlAdmin')

        status = self.admin_server.start()

        if status == ccc.STATUS_SUCCESS:
            print 'callControlAdmin server started'
            self.admin_server_started = True
        else:
            # If we did not start the server successfully, clean up
            self.stop()

        returncode = self.admin_server.pobject.wait()
        print 'process died, returncode = %s' % str(returncode)

        return status


    def stop(self):

        if self.admin_server is not None and self.admin_server_started:
            self.admin_server.stop()



