
import test
import testData
import object
import objectMap
import squishinfo
import squish

import sys

from pdctest.MediaControl.lib.mediaControlServer import MediaControlServer


def main():
    mp_id = 'windows_media_player'
    call_control_server = MediaControlServer(mp_id)
    status= call_control_server.run()
    sys.exit()

if __name__ == '__main__':
    main()