# -*- coding: utf-8 -*-
#
#  main.py
#  Office8n
#
#  Created by francois on 21/10/15.
#  Copyright (c) 2015 ftiff. All rights reserved.
#

# import modules required by application
import objc
import Foundation
import AppKit

from PyObjCTools import AppHelper

# import modules containing classes required to start application and load MainMenu.nib
import AppDelegate

# pass control to AppKit
AppHelper.runEventLoop()
