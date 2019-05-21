# -*- coding: utf-8 -*-

import logging
import os


# wpf basics
logging.debug('import IronPython.Wpf')
import clr
clr.AddReference("IronPython.Wpf")
import wpf

from System.Windows import Application, Window


# property binding
import pyro
import pyro.helpers
from pyro.library.wpf.notify import NotifyPropertyChangedBase, notify_property

# for generate_image
from System.Windows.Media.Imaging import BitmapImage, BitmapSource
from System.Windows.Media import PixelFormats
import System



class ViewModel(NotifyPropertyChangedBase):

    def __init__(self):
        super(ViewModel, self).__init__()
        
        self.size = '10'
        
        # initialize dynamic image
        self._red = 90
        self._green = 55
        self._blue = 200
        self.generate_image()
        
        
    def generate_image(self):
        # soure: https://docs.microsoft.com/en-us/dotnet/framework/wpf/graphics-multimedia/how-to-create-a-new-bitmapsource
        
        # Define parameters used to create the BitmapSource.
        # 32 bits per pixel / 4 bytes
        pf = PixelFormats.Bgr32 
        width = 16
        height = 16
        rawStride = (width * pf.BitsPerPixel ) / 8
        
        # Initialize the image with data.
        pixel = [self.blue or 0, self.green or 0, self.red or 0,0]
        rawImageValues = pixel * width * height
        rawImage = System.Array[System.Byte](rawImageValues)
        
        # Create a BitmapSource.
        # array len of rawImage show be:  Height * Stride * Format.BitsPerPixel/8
        bitmap = BitmapSource.Create(width, height,
            96, 96, pf, None,
            rawImage, rawStride)
        
        #set source
        self.img_source = bitmap

    @notify_property
    def size(self):
        return self._size

    @size.setter
    def size(self, value):
        self._size = value
    
    
    @notify_property
    def red(self):
        return self._red

    @red.setter
    def red(self, value):
        self._red = value
        self.generate_image()

        
    @notify_property
    def green(self):
        return self._green

    @green.setter
    def green(self, value):
        self._green = value
        self.generate_image()
    
    
    @notify_property
    def blue(self):
        return self._blue

    @blue.setter
    def blue(self, value):
        self._blue = value
        self.generate_image()
    
    
    @notify_property
    def img_source(self):
        return self._img_source

    @img_source.setter
    def img_source(self, value):
        self._img_source = value
    
    
    


class TestWindow(Window):
    
    def __init__(self):
        filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), 'dialog.xaml')
        wpf.LoadComponent(self, filename)
        self._vm = ViewModel()
        self.DataPanel.DataContext = self._vm

    def __getattr__(self, name):
        # provides easy access to XAML elements (e.g. self.Button)
        return self.root.FindName(name)
    
    def reset_initial_size(self, sender, event):
        # must be string to two-way binding work correctly
        self._vm.size = '10'

    def generate_color_image(self, sender, event):
        self._vm.generate_image()
    
    def show_message(self, sender, event):
        pyro.helpers.message("Hello world!")




def show(window_handle):
    logging.debug("dialog.show dialog")
    wnd = TestWindow()
    System.Windows.Interop.WindowInteropHelper(wnd).Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
    wnd.ShowDialog()
    
