# -*- coding: utf-8 -*-

from __future__ import absolute_import, division, print_function

import clr
import os.path
import traceback
import argparse

import reg
import helper

import System.Environment



class AppInfo(object):
    load_behavior = {
        'pyro' : 2
        }
    register_addins = {'pyro'}


class PowerPoint(AppInfo):
    addins_regpath = reg.office_default_path('PowerPoint')
    # register_addins = {'pyro'}
    # load_behavior = {
    #     'pyro' : 3
    #     }


class Word(AppInfo):
    addins_regpath = reg.office_default_path('Word')


class Excel(AppInfo):
    addins_regpath = reg.office_default_path('Excel')


class Outlook(AppInfo):
    addins_regpath = reg.office_default_path('Outlook')


class Visio(AppInfo):
    addins_regpath = reg.PathString('Software') / 'Microsoft' / 'Visio' / 'Addins'


APPS = [
    PowerPoint,
    Excel,
    Word,
    Outlook,
    Visio,
    ] 


class AddinInfo(object):
    pass


class pyro(AddinInfo):
    key = 'pyro'
    prog_id = 'pyro.AddIn'
    # FIXME: IMPORTANT -- change the Guid for your application
    # here and in Addin.cs
    uuid = '{A2BE0273-DF1B-461F-AF89-AA8B32A0C778}'
    name = 'Python riding on Office (pyro) -- DEMO'
    description = 'Python riding on Office (pyro) -- DEMO'
    dll = 'pyro.dll'


# class pyroTaskPane(AddinInfo):
#     key = 'pyro_taskpane'
#     prog_id = 'pyro.TaskPane'
#     uuid = '{76FD3062-86C8-11E4-BE43-6336340000B1}'
#     name = 'pyro Task Pane'
#     description = 'Business Kasper Toolbox Task Pane'
#     dll = 'pyro.dll'
#


ALL_ADDINS = [
    pyro,
    #pyroTaskPane,
    ]


INSTALL_ADDINS = [
    pyro,
    #pyroTaskPane,
    ]


def go_up(path, *directories):
    current = os.path.normpath(os.path.abspath(path))
    for d in directories:
        current, tail = os.path.split(current)
        if tail != d:
            raise ValueError('expected path component %r, got %r' % (d, tail))
    return current


INSTALL_BASE = go_up(os.path.dirname(__file__), 'installer')


class RegistryInfoService(object):
    def __init__(self, apps=None, addins=None, install_base=None, uninstall=False):
        if apps is None:
            apps = list(APPS)
        if addins is None:
            if uninstall:
                addins = list(ALL_ADDINS)
            else:
                addins = list(INSTALL_ADDINS)
        if install_base is None:
            install_base = INSTALL_BASE
            
        self.apps = apps
        self.addins = {a.key: a for a in addins}
        self.install_base = install_base
        self.uninstall = uninstall
            
    def get_addin_assembly_info(self, addin_info):
        return dict(
            prog_id=addin_info.prog_id,
            uuid=addin_info.uuid,
            assembly_path=os.path.join(self.install_base, 'bin', addin_info.dll),
            )
    
    def iter_addin_assembly_infos(self):
        for addin in self.addins.values():
            yield self.get_addin_assembly_info(addin)
                
    def get_application_addin_info(self, app, addin):
        return dict(prog_id=addin.prog_id,
                   friendly_name=addin.name,
                   description=addin.description,
                   addins_regpath=app.addins_regpath,
                   load_behavior=app.load_behavior.get(addin.key, 0),
                   )

    def iter_application_addin_infos(self):
        all_addins = list(self.addins)
        for app in self.apps:
            if self.uninstall:
                addins = all_addins
            else:
                addins = app.register_addins
            
            for addin_key in addins:
                addin = self.addins[addin_key]
                yield self.get_application_addin_info(app, addin)


def check_wow6432():
    ''' returns true if office-32-bit is running on 64 bit machine '''
    iop_base = 'Microsoft.Office.Interop.'        
    
    apps = ['PowerPoint',
            'Excel']
    
    os_64 = System.Environment.Is64BitOperatingSystem
    if os_64 == False:
        return False
    
    office_is_32 = set()
    for app_name in apps:
        iop_name = iop_base + app_name
        try:
            clr.AddReference(iop_name)
            module = None
            # FIXME: this is ugly, but __import__(iop_name) does not seem to work
            exec 'import ' + iop_name + ' as module'
            app = module.ApplicationClass()
            try:
                office_is_32.add(app.OperatingSystem.startswith('Windows (32-bit)'))
            finally:
                app.Quit()
        except:
            traceback.print_exc()
            
    if len(office_is_32) == 0:
        raise AssertionError('failed to get bitness of all tested office applications')
    
    return os_64 and (True in office_is_32)

def fmt_load_behavior(integer):
    return ('%08x' % integer).upper()


class Installer(object):
    def __init__(self, install_base=None, wow6432=None):
        if install_base is None:
            install_base = INSTALL_BASE

        self.install_base = install_base
        
        if wow6432 is None:
            print('checking system and office for 32/64 bit')
            wow6432 = check_wow6432()
        
        self.wow6432 = wow6432
    

    
    def unregister(self):
        reginfo = RegistryInfoService(uninstall=True, install_base=self.install_base)
        for info in reginfo.iter_application_addin_infos():
            reg.AddinRegService(**info).unregister_addin()
            
        for info in reginfo.iter_addin_assembly_infos():
            reg.AssemblyRegService(wow6432=self.wow6432, **info).unregister_assembly()
            
    def install(self):
        self.unregister()
        try:
            self.register()
            print("\nInstallation ready -- addin available after Office restart")
        except:
            self.unregister()
    
    def register(self):
        reginfo = RegistryInfoService(install_base=self.install_base)
        for info in reginfo.iter_addin_assembly_infos():
            reg.AssemblyRegService(wow6432=self.wow6432, **info).register_assembly()

        for info in reginfo.iter_application_addin_infos():
            reg.AddinRegService(**info).register_addin()


def install(apps=["powerpoint"]):
    try:
        # uninstall
        Installer(wow6432=True).unregister()
        Installer(wow6432=False).unregister()

        #app load beavhior
        if "powerpoint" in apps:
            PowerPoint.load_behavior = { 'pyro' : 3, 'pyro_dev': 3 }
        if "excel" in apps:
            Excel.load_behavior = { 'pyro' : 3, 'pyro_dev': 3 }
        if "word" in apps:
            Word.load_behavior = { 'pyro' : 3, 'pyro_dev': 3 }
        if "visio" in apps:
            Visio.load_behavior = { 'pyro' : 3, 'pyro_dev': 3 }
        if "outlook" in apps:
            Outlook.load_behavior = { 'pyro' : 3, 'pyro_dev': 3 }

        # install
        installer = Installer()
        installer.install()
    except:
        helper.exception_as_message()
        

def uninstall():
    try:
        Installer(wow6432=True).unregister()
        Installer(wow6432=False).unregister()
    except:
        helper.exception_as_message()


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('-u', '--uninstall', action='store_true', help='Remove all Pyro registry entries')
    parser.add_argument('-a', '--app', action='append', default=['powerpoint'], help='Application in which Pyro is activated')
    return parser.parse_args()


def main():
    args = parse_args()
    if args.uninstall:
        print('Uninstalling Pyro from current directory...')
        uninstall()
        print("\nPyro successfully uninstalled")
    else:
        if helper.is_admin():
            if helper.yes_no_question('Are you sure to run Pyro installer as admin?'):
                start_install = True
            else:
                start_install = False
                print('Pyro installation cancelled')
        else:
            start_install = True

        if start_install:
            print('Installing Pyro in current directory...')

            # deactivate previous installation
            uninstall()

            # start installation
            install(args.app)

if __name__ == '__main__':
    main()
