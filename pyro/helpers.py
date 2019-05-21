# -*- coding: utf-8 -*-

from __future__ import print_function

import os.path
import logging

import ctypes #required for messagebox


#class is compatible to systems.forms
class Forms(object):
    '''
    interface to windows-user32-dll to show simple messages
    '''
    
    class MessageBoxButtons(object):
        OK =                     0x00000000L #OK
        OKCancel =               0x00000001L #OK | Cancel
        AbortRetryIgnore =       0x00000002L #Abort | Retry | Ignore
        YesNoCancel =            0x00000003L #Yes | No | Cancel
        YesNo =                  0x00000004L #Yes | No
        RetryCancel =            0x00000005L #Retry | Cancel 
        CancelTryAgainContinue = 0x00000006L #Cancel | Try Again | Continue

    class MessageBoxIcon(object):
        #None =        0x00000000L
        Stop =        0x00000010L
        Error =       0x00000010L
        Hand =        0x00000010L
        Question =    0x00000020L
        Exclamation = 0x00000030L
        Warning     = 0x00000030L
        Information = 0x00000040L
        Asterisk    = 0x00000040L

    class DialogResult(object):
        OK          = 1
        Yes         = 6
        No          = 7
        Cancel      = 2
        Abort       = 3
        Continue    = 11
        Ignore      = 5
        Retry       = 4
        TryAgain    = 10

    class MessageBox(object):
        @staticmethod
        def Show(text, title, buttons, icon):
            def _get_hwnd():
                try:
                    return ctypes.windll.user32.GetForegroundWindow()
                except:
                    return 0
            
            return ctypes.windll.user32.MessageBoxW(_get_hwnd(), text, title, buttons | icon | 0x00002000L | 0x00010000L) #TASKMODAL | SETFOREGROUND


def message(text, title="PYRO"):
    ''' show simple message box '''
    Forms.MessageBox.Show(text, title, Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information)

def confirmation(text, title="pyro", buttons=Forms.MessageBoxButtons.OKCancel):
    ''' show confirmation-box with buttons '''
    result = Forms.MessageBox.Show(text, title, buttons, Forms.MessageBoxIcon.Question)
    if buttons == Forms.MessageBoxButtons.OKCancel or buttons == Forms.MessageBoxButtons.YesNo:
        if result == Forms.DialogResult.OK or result == Forms.DialogResult.Yes:
            return True
        else:
            return False
    else:
        return result


def exception_as_message():
    ''' show last exceptioon in a message box '''
    from cStringIO import StringIO
    import traceback
    
    fd = StringIO()
    traceback.print_exc(file=fd)
    traceback.print_exc()
    message(fd.getvalue(), title="pyro Exception")


