# -*- coding: utf-8 -*-

import logging
import os


# wpf basics
import clr
clr.AddReference("IronPython.Wpf")
import wpf

import System
from System.Windows import Application, Window

# property binding
import pyro
from pyro.library.wpf.notify import NotifyPropertyChangedBase, notify_property

# for MahApps.Metro
import sys
import os
assembly_filename = os.path.realpath((os.path.join(os.path.dirname(os.path.realpath(__file__)), '..', '..', 'bin', 'external', 'MahApps.Metro.dll')))
logging.debug('adding assembly: %s' % assembly_filename)
logging.debug('sys paths before clr.Add: %s' % sys.path)
clr.AddReferenceToFileAndPath(assembly_filename)
logging.debug('sys paths after clr.Add: %s' % sys.path)
from MahApps.Metro.Controls import MetroWindow



class ViewModel(NotifyPropertyChangedBase):

    def __init__(self):
        super(ViewModel, self).__init__()
    
    


class TestWindow(MetroWindow):
    
    def __init__(self):
        filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), 'mahapps_dialog.xaml')
        wpf.LoadComponent(self, filename)
        self._vm = ViewModel()
        self.DataPanel.DataContext = self._vm

    def __getattr__(self, name):
        # provides easy access to XAML elements (e.g. self.Button)
        return self.root.FindName(name)
    
    


def show(window_handle):
    logging.debug("dialog.show dialog")
    wnd = TestWindow()
    System.Windows.Interop.WindowInteropHelper(wnd).Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
    wnd.ShowDialog()
    

