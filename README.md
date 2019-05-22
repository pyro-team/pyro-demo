# PYRO DEMO

Python riding on Office - Demo application for Python Office COM Addin.


## Introduction

Pyro is a demo project providing basic concepts on how to run Python scripts from a Microsoft Office COM Addin.
The application demonstrates how to do the following things.

In a C# COM Addin:

  * build a simple Shared COM Addin
  * deployment of dlls as Office addin (registration in Windows registry)
  * integrate IronPython as scripting host
  * implement the ribbon interface (`IRibbonExtensibility`), to add tabs/groups/buttons to the ribbon
  * implement the task pane interface (`ICustomTaskPaneConsumer`), to add task panes to your addin (not provided yet)
  * show and hide task panes from the ribbon (not provided yet)

and in Python:
  
  * manipulate office objects
  * handle ribbon and task pane events (latter not provided yet)
  * create WPF windows
  * two-way-binding to Properties from WPF controls
  * add thrid party librarties to your WPF windows and task panes (FluentRibbon, MahApps,  Google Material Icons)
  * use windows-dlls, e.g. user32-dll to show simple windows message boxes


## Quick Start

Start by cloning the repository an run `installer\install.bat`.
The installation will register the COM addin in the Office products (only active in PowerPoint by default).

If you use Office 2010, you need to build the addin first by running `dotnet\build2010.bat`.
To build the addin you will need .Net Framework SDK to be installed.


## System requirements

Pyro runs on Windows with Microsoft Office 2010 or more recent Office versions. There is no Mac-Version for Pyro since COM addins are not supported in Mac-Office versions.


## Contributions

Pyro currently uses the following third party libraries:

 * [IronPython](https://github.com/IronLanguages/ironpython2)
 * [Fluent.Ribbon](https://github.com/fluentribbon/Fluent.Ribbon)
 * [ControlzEx](https://github.com/ControlzEx/ControlzEx)
 * [MahApps.Metro](https://github.com/MahApps/MahApps.Metro)
 * [MouseKeyHooks](https://github.com/gmamaladze/globalmousekeyhook)
 
 
## Links and resources

The following links and resources can be helpful if you plan to build your own Python COM Addin:

 * [InnoSetup](http://www.jrsoftware.org/isinfo.php)
 * [Google Material Icons](https://material.io/tools/icons/)
 * [Material Design Icons](https://materialdesignicons.com/)
 * [Office UI Fabric](https://developer.microsoft.com/de-de/fabric#/)
 * [Differences between Shared addin and VSTO addin](https://social.msdn.microsoft.com/Forums/vstudio/en-US/3f97705a-6052-4296-a10a-bfa3a39ab4e7/shared-addin-vs-vsto-addin-whats-the-difference-betweenhow-can-i-tell-if-im-developing)

