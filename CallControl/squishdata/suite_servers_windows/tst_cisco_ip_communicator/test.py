
import test
import testData
import object
import objectMap
import squishinfo
import squish

import sys

import pdctest.CallControl.lib.callControlServer as ccs

def usage():
    print '''
        Usage:
            squish_start_call_control_server sp_id

            sp_id:     String id of softphone type
        '''

def main():
    sp_id = 'cisco_ip_communicator'
    call_control_server = ccs.CallControlServer(sp_id)
    status= call_control_server.run()
    sys.exit()

if __name__ == '__main__':
    main()