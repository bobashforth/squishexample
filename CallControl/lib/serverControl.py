#-------------------------------------------------------------------------------
# Name:         serverControl
# Purpose:      This module embodiest the attributes and methods needed for
#               simple server control operations. At present only start()
#               and stop() methods are provided, but the class can be
#               enhanced as needed.
#
# Author:      rashforth
#
# Created:     23/10/2012
# Copyright:   (c) Plantronics 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import exceptions
import subprocess
import pdctest.CallControl.lib.call_control_constants as ccc

class ServerControl(object):

    def __init__(self, commandline, procname=None):
        self.commandline = commandline
        self.procname = procname
        self.logfile = None
        self.pid = None


    def spawnproc(self, cmdstring, logfile=None):

        try:
            if logfile is not None:
                p = subprocess.Popen(cmdstring, stdout=logfile, stderr=logfile)
            else:
                p = subprocess.Popen(cmdstring)

        except Exception as e:
            print 'Failed to execute command'
            print e.args
            if logfile is not None:
                self.logfile.close()
            p = None

        if p is not None:
            self.pid = p.pid
            self.logfile = None

        return p


    def start(self):

        if self.procname is None:
            # Default procname is basename of executable in command string
            self.procname = self.commandline.split()[0].strip()

        logfile_name = '%s.log' % self.procname
        print 'logfile_name = %s' % logfile_name

        try:
            self.logfile = open(logfile_name, 'w')

        except Exception as e:
            print 'Error opening logfile %s' % logfile_name
            status = ccc.STATUS_FAILURE
            return None

        print 'Executing command string %s ' % self.commandline
        p = self.spawnproc(self.commandline, self.logfile)

        if p is None:
            print 'Error in executing command'
            self.logfile.close()
            status = ccc.STATUS_FAILURE
        else:
            print 'Command executed successfully.'
            self.pid = p.pid
            status = ccc.STATUS_SUCCESS

        return status


    def stop(self):

        if self.pid is not None:
            commandline = 'taskkill /T /F /PID %s' % str(self.pid)

            print 'Killing process %s' % self.procname
            p = self.spawnproc(commandline)

            if p is None:
                status = ccc.STATUS_FAILURE
            else:
                self.pid = None
                if self.logfile is not None:
                    self.logfile.close()
                status = ccc.STATUS_SUCCESS
        else:
            print 'Cannot kill process %s, pid is None' % self.procname
            status = ccc.STATUS_FAILURE

        return status


    def main():
        pass

if __name__ == '__main__':
    main()
