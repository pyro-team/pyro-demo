# -*- coding: utf-8 -*-

def import_forms():
    import clr
    clr.AddReference('System.Windows.Forms')
    import System.Windows.Forms as forms
    return forms

def import_drawing():
    import clr
    clr.AddReference('System.Drawing')
    import System.Drawing as drawing
    return drawing

def import_powerpoint():
    import clr
    clr.AddReference('Microsoft.Office.Interop.PowerPoint')
    import Microsoft.Office.Interop.PowerPoint as powerpoint
    return powerpoint

def import_excel():
    import clr
    clr.AddReference('Microsoft.Office.Interop.Excel')
    import Microsoft.Office.Interop.Excel as excel
    return excel

def import_outlook():
    import clr
    clr.AddReference('Microsoft.Office.Interop.Outlook')
    import Microsoft.Office.Interop.Outlook as outlook
    return outlook

def import_officecore():
    import clr
    clr.AddReference('Office')
    import Microsoft.Office.Core as officecore
    return officecore

