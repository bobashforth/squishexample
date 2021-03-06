#-------------------------------------------------------------------------------
# Name:        callControlFactory.py
# Purpose:     Provides a factory interface for instatiating a softphone
#              object corresponding to the associated name string.
#
# Author:      rashforth
#
# Created:     28/08/2012
# Copyright:   (c) Plantronics 2012
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import sys
import call_control_constants as ccc

import pdctest.CallControl.softphones_windows.microsoft_lync_CallControl
import pdctest.CallControl.softphones_windows.microsoft_lync_2013_CallControl
import pdctest.CallControl.softphones_windows.skype_CallControl
import pdctest.CallControl.softphones_windows.avaya_ip_agent_CallControl
import pdctest.CallControl.softphones_windows.avaya_ip_softphone_CallControl
import pdctest.CallControl.softphones_windows.avaya_onex_agent_CallControl
import pdctest.CallControl.softphones_windows.avaya_onex_communicator_CallControl
import pdctest.CallControl.softphones_windows.cisco_ip_communicator_CallControl
import pdctest.CallControl.softphones_windows.cucilync_CallControl
import pdctest.CallControl.softphones_windows.cucimoc_CallControl
import pdctest.CallControl.softphones_windows.cupc_CallControl
import pdctest.CallControl.softphones_windows.cisco_webex_connect_CallControl
import pdctest.CallControl.softphones_windows.ibm_sametime_CallControl
import pdctest.CallControl.softphones_windows.microsoft_oc_r2_CallControl
import pdctest.CallControl.softphones_windows.nec_sp350_CallControl
import pdctest.CallControl.softphones_windows.shoretel_call_manager_CallControl
import pdctest.CallControl.softphones_windows.shoretel_communicator_CallControl

def getCallControlInstance(softphone_name):

    if softphone_name == "microsoft_lync":
        return pdctest.CallControl.softphones_windows.microsoft_lync_CallControl.Microsoft_lync_CallControl(softphone_name)
    elif softphone_name == "microsoft_lync_2013":
        return pdctest.CallControl.softphones_windows.microsoft_lync_2013_CallControl.Microsoft_lync_2013_CallControl(softphone_name)
    elif softphone_name ==  "skype":
        return pdctest.CallControl.softphones_windows.skype_CallControl.Skype_CallControl(softphone_name)
    elif softphone_name == "avaya_ip agent":
        return pdctest.CallControl.softphones_windows.avaya_ip_agent_CallControl.Avaya_ip_agent_CallControl(softphone_name)
    elif softphone_name == "avaya_ip_softphone":
        return pdctest.CallControl.softphones_windows.avaya_ip_softphone_CallControl.Avaya_ip_softphone_CallControl(softphone_name)
    elif softphone_name == "avaya_onex_agent":
        return pdctest.CallControl.softphones_windows.avaya_onex_agent_CallControl.Avaya_onex_agent_CallControl(softphone_name)
    elif softphone_name == "avaya_onex_communicator":
        return pdctest.CallControl.softphones_windows.avaya_onex_communicator_CallControl.Avaya_onex_communicator_CallControl(softphone_name)
    elif softphone_name == "cisco_ip_communicator":
        return pdctest.CallControl.softphones_windows.cisco_ip_communicator_CallControl.Cisco_ip_communicator_CallControl(softphone_name)
    elif softphone_name == "cucilync":
        return pdctest.CallControl.softphones_windows.cucilync_CallControl.Cucilync_CallControl(softphone_name)
    elif softphone_name == "cucimoc":
        return pdctest.CallControl.softphones_windows.cucimoc_CallControl.Cucimoc_CallControl(softphone_name)
    elif softphone_name == "cupc":
        return pdctest.CallControl.softphones_windows.cupc_CallControl.Cupc_CallControl(softphone_name)
    elif softphone_name == "cisco_webex_connect":
        return pdctest.CallControl.softphones_windows.cisco_webex_connect_CallControl.Cisco_webex_connect_CallControl(softphone_name)
    elif softphone_name == "ibm_sametime":
        return pdctest.CallControl.softphones_windows.ibm_sametime_CallControl.Ibm_sametime_CallControl(softphone_name)
    elif softphone_name == "microsoft_oc_r2":
        return pdctest.CallControl.softphones_windows.microsoft_oc_r2_CallControl.Microsoft_oc_r2_CallControl(softphone_name)
    elif softphone_name == "nec_sp350":
        return pdctest.CallControl.softphones_windows.nec_sp350_CallControl.Nec_sp350_CallControl(softphone_name)
    elif softphone_name == "shoretel_call_manager":
        return pdctest.CallControl.softphones_windows.shoretel_call_manager_CallControl.Shoretel_call_manager_CallControl(softphone_name)
    elif softphone_name == "shoretel_communicator":
        return pdctest.CallControl.softphones_windows.shoretel_communicator_CallControl.Shoretel_communicator_CallControl(softphone_name)
    else:
        print 'Invalid softphone name %s, ' \
            'cannot provide CallControl instance' % softphone_name
        return None
