# -*- coding: utf-8 -*-

import logging
import os.path
import traceback

import pyro
import pyro.helpers


# =================
# = Configuration =
# =================

CONFIG_LOGGING = True
CONFIG_LOVLEVEL = logging.DEBUG



# ======================
# = Initialize Logging =
# ======================

#FIXME: gleiche Log-Datei wie im .Net-Addin verwenden. Verwendung von pyro-debug.log führt noch zu Fehlern (Verlust von log-Text), da die Logger nicht Zeilenweise schreiben. Alternativ logging über C#-Addin-Klasse durchführen
if CONFIG_LOGGING:
    log_level = CONFIG_LOVLEVEL
    logging.basicConfig(
        filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "pyro-debug-py.log"), 
        filemode='w',
        format='%(asctime)s %(levelname)s: %(message)s', 
        level=log_level
    )





class AddIn(object):
    
    
    # ==================================
    # = initialization with .Net-Addin =
    # ==================================
    
    def __init__(self):
        logging.debug("AddIn.__init__")
        self.dialog_window = None
        self.main_window_handle = 0
    
    
    def on_create(self, context):
        logging.debug("AddIn.on_create")
        self.context = context
        
        #pyro.helpers.message("Python-Addin is active!")
        
    
    def on_destroy(self):
        logging.debug("AddIn.on_destroy")
        
    
    # ===================
    # = Window handling =
    # ===================
    
    def set_window_hwnd(self, hwnd):
        logging.debug("AddIn.set_window_hwnd %s" % hwnd)
        self.main_window_handle = hwnd
    
    
    # ==================================
    # = IRibbonExtensibility interface =
    # ==================================
    
    def get_custom_ui(self, ribbon_id):
        logging.debug("AddIn.get_custom_ui: %s" % ribbon_id)
        
        filename=os.path.join(os.path.dirname(os.path.realpath(__file__)), "..", "resources", "xml", "customui.xml")
        with open(filename, 'r') as myfile:
          customui = myfile.read()
        
        logging.info('customui: %s ' % customui)
        return customui
    
    
    # ======================================
    # = IRibbonExtensibility Ribbon events =
    # ======================================
    
    def on_action(self, control):
        logging.debug("AddIn.on_action, control-id=%s" % control.id)
        
        try:
            # run action from control.tag
            getattr(self,control.tag)()
        except:
            logging.error(traceback.format_exc())
        
    
    
    # ================
    # = Demo actions =
    # ================
    
    def reload_pyro(self):
        import pyro.console
        try:
            addin = self.context.app.COMAddIns["pyro.AddIn"]
            addin.Connect = False
            addin.Connect = True
        except Exception, e:
            pyro.console.show_message(str(e))
    
    def console(self):
        import pyro.console as co
        co.console.Visible = True
        co.console.scroll_down()
        co.console.BringToFront()
        co.console._globals['context'] = self.context
        
    def simple_action(self):
        pyro.helpers.message("Ribbon button action from Python!")

    
    
    
    def show_dialog(self):
        logging.debug("show dialog")
        from demo import dialog
        dialog.show(self.main_window_handle)
    
    def show_fluent_dialog(self):
        logging.debug("show fluent dialog")
        from demo import fluentdialog
        fluentdialog.show(self.main_window_handle)
    
    def show_mahappsmetro_dialog(self):
        logging.debug("show MahApps.Metro dialog")
        from demo import mahapps_dialog
        mahapps_dialog.show(self.main_window_handle)
    
    
    
