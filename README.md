# PYRO DEMO

Python riding on Office - Demo application for Python Office COM Addin.


## Introduction

This Demo project shows how to run Python scripts from a Microsoft Office COM Addin.
It is the basis for the *Python riding on Office* project.

This project demonstrates how to do the following things.

In a C# COM Addin:

  * integrate IronPython as scripting host
  * simple registration of dlls as Office addin
  * implement the ribbon interface (`IRibbonExtensibility`), to add tabs/groups/buttons to the ribbon
  * implement the task pane interface (`xxx`), to add task panes to your addin
  * show and hide task panes from the ribbon

and in Python:
  
  * manipulate office objects
  * handle ribbon and task pane events
  * create WPF windows
  * two-way-binding to Properties from WPF controls
  * add thrid party librarties to your WPF windows and task panes (FluentRibbon, MahApps,  Google Material Icons)
  * use windows-dlls, e.g. user32-dll to show simple windows message boxes


## Quick Start

Start by cloning the repository an run `installer\install.bat`.
The installation will register the COM Addin in the Office products (only active in PowerPoint by default).

If you use Office 2010, you need to build the Addin first by running `dotnet\build2010.bat`.


## System requirements

Pyro runs on Windows with Microsoft Office 2010 or more recent Office versions. There is no Mac-Version for Pyro since COM addins are not supported in Mac-Office versions.


## Contributions

 * [IronPython](https://github.com/IronLanguages/ironpython2)
 * [Fluent.Ribbon](https://github.com/fluentribbon/Fluent.Ribbon)
 * [ControlzEx](https://github.com/ControlzEx/ControlzEx)
 * [MahApps.Metro](https://github.com/MahApps/MahApps.Metro)
 * [MouseKeyHooks](https://github.com/gmamaladze/globalmousekeyhook)
 * [InnoSetup](http://www.jrsoftware.org/isinfo.php)
 * [Google Material Icons](https://material.io/tools/icons/) & [Material Design Icons](https://materialdesignicons.com/)
