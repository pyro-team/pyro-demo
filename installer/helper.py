# -*- coding: utf-8 -*-

import ctypes
import os

def is_admin():
    #test for admin
    try:
        return os.getuid() == 0
    except AttributeError:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0


def yes_no_question(question):
    reply = str(raw_input(question + ' (y/n): ')).lower().strip()
    if reply[0] == 'y':
        return True
    else:
        return False


def exception_as_message():
    import StringIO
    import traceback

    fd = StringIO.StringIO()
    traceback.print_exc(file=fd)
    traceback.print_exc()


