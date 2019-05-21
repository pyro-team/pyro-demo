
using System;
using System.Diagnostics;

// office libraries
using Microsoft.Office.Core;
// using Microsoft.Office.Tools;

// to determine host application and hook application events
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Visio = Microsoft.Office.Interop.Visio; //Note: Visio Interop need to be embedded!
using Outlook = Microsoft.Office.Interop.Outlook;

internal enum HostApplication {Unknown=0, Excel, PowerPoint, Word, Visio, Outlook}



namespace pyro
{
            
    public partial class AddIn
    {
        // The AddIn class is the main pyro class and implements all necessary interfaces
        // to interact with Office.
        // 
        // This part implents:
        //  - detection which host (office application) is running the addin
        //  - initialization of office application level events
        //

        // Host detection
        private HostApplication host = HostApplication.Unknown;
        private string hostAppName;


        public void DetermineHostApplication(object application)
        {
            try {
                if (application is Excel.Application)
                {
                    host = HostApplication.Excel;
                    hostAppName = ((Excel.Application)application).Name;
                    Trace.TraceInformation("host application: " + hostAppName);
                    //BindExcelEvents((Excel.Application)application);
                }
                else if (application is PowerPoint.Application)
                {
                    host = HostApplication.PowerPoint;
                    hostAppName = ((PowerPoint.Application)application).Name;
                    Trace.TraceInformation("host application: " + hostAppName);
                    //BindPowerPointEvents((PowerPoint.Application)application);
                }
                else if (application is Word.Application)
                {
                    host = HostApplication.Word;
                    hostAppName = ((Word.Application)application).Name;
                    Trace.TraceInformation("host application: " + hostAppName);
                    //BindWordEvents((Word.Application)application);
                }
                else if (application is Visio.Application)
                {
                    host = HostApplication.Visio;
                    hostAppName = ((Visio.Application)application).Name;
                    Trace.TraceInformation("host application: " + hostAppName);
                    // BindVisioEvents((Visio.Application)application);
                }
                else if (application is Outlook.Application)
                {
                    host = HostApplication.Outlook;
                    hostAppName = ((Outlook.Application)application).Name;
                    Trace.TraceInformation("host application: " + hostAppName);
                    // BindOutlookEvents((Outlook.Application)application);
                }
                else
                {
                    Trace.TraceInformation("host application unknown");
                }
            } catch (Exception) {
                Trace.TraceInformation("error dertermining host application (maybe visio interop not installed)");
            }
        }
        
        

    }
}