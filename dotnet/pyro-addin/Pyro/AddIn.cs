
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;

// for reading xml-config
using System.Configuration;

// for com interface
using Extensibility;

// for ComVisible, Guid, etc.
using System.Runtime.InteropServices;

// for MessageBox
using System.Windows.Forms; 


namespace pyro
{
    
    // FIXME: IMPORTANT -- change the Guid for your application
    // here and in install.py
    [ComVisible(true)]
    [ProgId("pyro.AddIn")]
    [Guid("A2BE0273-DF1B-461F-AF89-AA8B32A0C778")]
        
    public partial class AddIn : Extensibility.IDTExtensibility2
    {
        // The AddIn class is the main pyro class and implements all necessary interfaces
        // to interact with Office.
        // 
        // This part implents:
        //  - constructor and destructor
        //  - loading configuration
        //  - logging
        //  - Extensibility interface 
        //
        
        
        // Addin instance
        private readonly int instance_id;
        private static int instance_id_counter = 0;
        private static int finalize_counter = 0;
        private static int connected_counter = 0;
        private static int disconnected_counter = 0;
        private bool created;
        private bool broken;
        
        // Configuration
        private string config_ironpython_root;
        private string config_ipy_addin_path;
        private string config_ipy_addin_module;
        private bool   config_pydev_debug;
        private string config_pydev_codebase;
        private string config_ironpython_root_full;
        private string config_ipy_addin_path_full;
        
        
        // logging + debugging
        private TextWriterTraceListener listener;
        private FileStream logFileStream;
        private bool debug;
        
        //private static TraceSource _logger = new TraceSource("sourceName");
        
        #region Contructor and reset
        // ============================
        // = Constructor / Destructor =
        // ============================
        
        public AddIn() {
            instance_id_counter += 1;
            instance_id = instance_id_counter;
            
            ReloadConfig();
#if DEBUG
            // configure Debug logging
            string path = Path.Combine(config_ipy_addin_path_full, "pyro-debug.log");
            try
            {
                logFileStream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error creating FileStream for trace file \"{0}\":" +"\r\n{1}", path, ex.Message);
                return;
            }
            
            Debug.Listeners.Add(new TextWriterTraceListener(logFileStream));
            Debug.AutoFlush = true;
#endif
            
            Trace.AutoFlush = true;
            Trace.Flush();
            
            Trace.TraceInformation("Addin started");
            
        }
        
        
        
        ~AddIn() {
            finalize_counter += 1;
        }
        
        public void Dispose() {
            Trace.TraceInformation("Dispose");
            
            // finalize debug messages
            if (listener != null) {
                listener.Flush();
                listener.Close();
                listener.Dispose();
                listener = null;
                logFileStream.Close();
                logFileStream.Dispose();
            }
        }
        
        private void Reset() {
            Trace.TraceInformation("Reset");
            created = false;
            app = null;
            ipy = null;
            scope = null;
            python_delegate = null;
            context = null;
            broken = false;
            debug = false;
            
        }
        
        private void Message(string s) {
            MessageBox.Show(s);
        }
        
        #endregion
        
        
        #region Config
        // ==========
        // = Config =
        // ==========
        
        public void ReloadConfig() {
            Trace.TraceInformation("ReloadConfig");
            
            Configuration libConfig = ConfigurationManager.OpenExeConfiguration(Assembly.GetExecutingAssembly().Location);
            AppSettingsSection section = (libConfig.GetSection("appSettings") as AppSettingsSection);
            
            config_ironpython_root  = section.Settings["ironpython_root"].Value;
            config_ipy_addin_path   = section.Settings["ipy_addin_path"].Value;
            config_ipy_addin_module = section.Settings["ipy_addin_module"].Value;
            config_pydev_debug      = section.Settings["pydev_debug"].Value == "True";
            config_pydev_codebase   = section.Settings["pydev_codebase"].Value;
            
            section = null;
            libConfig = null;
            
            string codebase = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            config_ironpython_root_full = (config_ironpython_root == "" ? codebase : Path.GetFullPath(Path.Combine(codebase, config_ironpython_root)) );
            config_ipy_addin_path_full  = (config_ipy_addin_path == "" ? codebase : Path.GetFullPath(Path.Combine(codebase, config_ipy_addin_path)) );
            
        }
        
        #endregion
        
        
        
        
        
        
        #region Addin interface
        // ==========================
        // = Shared Addin Interface =
        // ==========================
        
        public void OnConnection(object application, ext_ConnectMode connect_mode, object addin_inst, ref Array custom)  {
            connected_counter += 1;
            OnConnection2(application);
        }
        
        public void OnConnection2(object application)
        {
            Trace.TraceInformation("OnConnection2 called");
            
            try {
                // determine host
                DetermineHostApplication(application);
                
                // reset addin and config
                Trace.TraceInformation("ReloadConfig");
                ReloadConfig();
                this.Reset();
                this.app = application;
                
                // initialize python instance
                Trace.TraceInformation("Initialize Python instance");
                LoadPython();
                
                // bootstrap addin
                BootstrapAddIn();
                created = true;
                
                // Window Handle
                IntPtr hwnd = Process.GetCurrentProcess().MainWindowHandle;
                this.python_delegate.set_window_hwnd(hwnd);
                
                
                
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        

        public void OnDisconnection(ext_DisconnectMode remove_mode, ref Array custom)
        {    
            Trace.TraceInformation("OnDisconnection: instance_id=" + instance_id);
            try {
                if(!broken) {
                    python_delegate.on_destroy();
                }
            } catch (Exception e) {
                Trace.TraceInformation(e.ToString());
            }
            try {
                if(ipy != null) {
                    ipy.Runtime.Shutdown();
                }
                Reset();
            } catch (Exception e) {
                Trace.TraceInformation(e.ToString());
            }
            disconnected_counter += 1;
            if (connected_counter == disconnected_counter) {
                Dispose();
            }
        }
        
        public void OnAddInsUpdate(ref Array custom)
        {    
        }
        
        public void OnStartupComplete(ref Array custom)
        {
        }
        
        public void OnBeginShutdown(ref Array custom)
        {    
            Trace.TraceInformation("OnBeginShutdown");
        }
        

        #endregion
        
        

        
//         #region Application events
//         // ======================
//         // = Application events =
//         // ======================
//
//         // EXCEL
//
//         private void BindExcelEvents(Excel.Application application)
//         {
//             ((Excel.AppEvents_Event)application).WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Excel_WorkbookOpen);
//             ((Excel.AppEvents_Event)application).NewWorkbook += new Excel.AppEvents_NewWorkbookEventHandler(Excel_NewWorkbook);
//         }
//
//         private void Excel_NewWorkbook(Excel.Workbook workbook)
//         {
//             Trace.TraceInformation("Excel: new workbook: " + workbook.FullName);
//             try {
// #if OFFICE2010
// #else
//                 //CreateTaskPaneForWindow(workbook.Windows[1]);
// #endif
//             } catch (Exception e) {
//                 Trace.TraceInformation(e.ToString());
//             }
//         }
//
//         private void Excel_WorkbookOpen(Excel.Workbook workbook)
//         {
//             Trace.TraceInformation("Excel: workbook opened: " + workbook.FullName);
//             try {
// #if OFFICE2010
// #else
//                 //CreateTaskPaneForWindow(workbook.Windows[1]);
// #endif
//             } catch (Exception e) {
//                 Trace.TraceInformation(e.ToString());
//             }
//         }
//
//         // POWER POINT
//
//         private void BindPowerPointEvents(PowerPoint.Application application)
//         {
//             //FIXME: On addin reload, these events cause a Microsoft.CSharp.RuntimeBinder.RuntimeBinderException as events are not unbinded!
//             ((PowerPoint.EApplication_Event)application).PresentationOpen += new PowerPoint.EApplication_PresentationOpenEventHandler(PowerPoint_PresentatonOpen);
//             ((PowerPoint.EApplication_Event)application).NewPresentation += new PowerPoint.EApplication_NewPresentationEventHandler(PowerPoint_NewPresentation);
//             ((PowerPoint.EApplication_Event)application).WindowSelectionChange += new PowerPoint.EApplication_WindowSelectionChangeEventHandler(PowerPoint_WindowSelectionChange);
//         }
//
//         private void PowerPoint_NewPresentation(PowerPoint.Presentation presentation)
//         {
//             Trace.TraceInformation("PowerPoint: new presentation: " + presentation.FullName);
//             try {
// #if OFFICE2010
// #else
//                 //CreateTaskPaneForWindow(presentation.Windows[1]);
// #endif
//             } catch (Exception e) {
//                 Trace.TraceInformation(e.ToString());
//             }
//         }
//
//         private void PowerPoint_PresentatonOpen(PowerPoint.Presentation presentation)
//         {
//             Trace.TraceInformation("PowerPoint: presentation opened: " + presentation.FullName);
//             try {
// #if OFFICE2010
// #else
//                 //CreateTaskPaneForWindow(presentation.Windows[1]);
// #endif
//             } catch (Exception e) {
//                 Trace.TraceInformation(e.ToString());
//             }
//         }
//
//         private void PowerPoint_WindowSelectionChange(PowerPoint.Selection selection)
//         {
//             Trace.TraceInformation("PowerPoint: window selection changed");
//             try {
//                 selection_type = (int)selection.Type;
//
//                 if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText) {
//                     int shape_id = 0;
//
//                     // Set values for fast enable events
//                     selection_shapes = 1;
//                     selection_containstextframe = true;
//
//                     // Store shape id
//                     if (selection.HasChildShapeRange) {
//                         shape_id = selection.ChildShapeRange[1].Id;
//                     } else {
//                         shape_id = selection.ShapeRange[1].Id;
//                     }
//
//                     // Selection is changed for each key press (e.g. while typing), therefore we introduce some thresholds:
//                     // Update timestamp after 2 seconds, when shape id changed, or (see MouseDownEvent) when mouse is clicked
//                     if ((DateTime.Now-ppt_last_selection_changed).TotalSeconds > 2 || ppt_last_selection_shape_id != shape_id) {
//                         ppt_last_selection_changed  = DateTime.Now;
//                         ppt_last_selection_shape_id = shape_id;
//                     } else {
//                         // stop method here, no python_delegate (at the bottom)
//                         return;
//                     }
//                 } else {
//                     ppt_last_selection_changed = DateTime.MinValue;
//                     // Set values for fast enable events
//                     if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) {
//                         if (selection.HasChildShapeRange) {
//                             selection_shapes = selection.ChildShapeRange.Count;
//                         } else {
//                             selection_shapes = selection.ShapeRange.Count;
//                         }
//                         selection_containstextframe = Ppt_Selection_Contains_Textframe(selection);
//                     } else {
//                         selection_shapes = 0;
//                         selection_containstextframe = false;
//                     }
//                 }
//
//                 python_delegate.ppt_selection_changed(selection);
//             } catch (Exception e) {
//                 Trace.TraceInformation(e.ToString());
//             }
//         }
//
//         private bool Ppt_Selection_Contains_Textframe(PowerPoint.Selection selection)
//         {
//             try {
//                 if (selection.HasChildShapeRange) {
//                     if (selection.ChildShapeRange.HasTextFrame == MsoTriState.msoFalse) {
//                         return false;
//                     } else {
//                         return true;
//                     }
//                 } else {
//                     if (selection.ShapeRange.HasTextFrame != MsoTriState.msoFalse || selection.ShapeRange.HasTable != MsoTriState.msoFalse) {
//                         return true;
//                     } else {
//                         foreach (PowerPoint.Shape el in selection.ShapeRange) {
//                             if (el.Type == MsoShapeType.msoGroup && el.GroupItems.Range(null).HasTextFrame != MsoTriState.msoFalse) {
//                                 return true;
//                             }
//                             if (el.Type == MsoShapeType.msoSmartArt) {
//                                 return true;
//                             }
//                         }
//                         return false;
//                     }
//                 }
//             } catch (Exception) {
//                 // Trace.TraceInformation(e.ToString());
//                 return false;
//             }
//         }
//
//         // WORD
//
//         private void BindWordEvents(Word.Application application)
//         {
//             ((Word.ApplicationEvents4_Event)application).DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(Word_DocumentOpen);
//             ((Word.ApplicationEvents4_Event)application).NewDocument += new Word.ApplicationEvents4_NewDocumentEventHandler(Word_NewDocument);
//         }
//
//         private void Word_NewDocument(Word.Document document)
//         {
//             Trace.TraceInformation("Word: new document: " + document.FullName);
//             try {
// #if OFFICE2010
// #else
//                 //CreateTaskPaneForWindow(document.Windows[1]);
// #endif
//             } catch (Exception e) {
//                 Trace.TraceInformation(e.ToString());
//             }
//         }
//
//         private void Word_DocumentOpen(Word.Document document)
//         {
//             Trace.TraceInformation("Word: document opened: " + document.FullName);
//             try {
// #if OFFICE2010
// #else
//                 //CreateTaskPaneForWindow(document.Windows[1]);
// #endif
//             } catch (Exception e) {
//                 Trace.TraceInformation(e.ToString());
//             }
//         }
//         #endregion


//         #region Window handles
//         // ===============================
//         // = Get window handle for forms =
//         // ===============================
//
//         public object GetActiveWindow()
//         {
//
//             if (host == HostApplication.Excel)
//             {
//                 return ((Excel.Application)context.app).ActiveWindow;
//             }
//             else if (host == HostApplication.PowerPoint)
//             {
//                 if ( ((PowerPoint.Application)context.app).Windows.Count == 0) {
//                     // Avoid error: System.Runtime.InteropServices.COMException (0x80048240): Application (unknown member) : Invalid request.  There is no currently active document window.
//                     Trace.TraceInformation("GetActiveWindow: no active Windows!");
//                     throw new NullReferenceException("No active windows");
//                 }
//                 return ((PowerPoint.Application)context.app).ActiveWindow;
//             }
//             else if (host == HostApplication.Word)
//             {
//                 return ((Word.Application)context.app).ActiveWindow;
//             }
//             // else if (host == HostApplication.Visio)
//             // {
//             //     return ((Visio.Application)context.app).ActiveWindow;
//             // }
//             else
//             {
//                 throw new NotSupportedException("Unknown host application");
//             }
//         }
//
//
//         public int GetWindowHandle(object window=null)
//         {
// #if OFFICE2010
//             return 0;
// #else
//             try
//             {
//                 if (window == null)
//                 {
//                     window = GetActiveWindow();
//                 }
//                 if (host == HostApplication.Excel)
//                 {
//                     int windowID = ((Excel.Window)window).Hwnd;
//                     return windowID;
//                 }
//                 else if (host == HostApplication.PowerPoint)
//                 {
//                     int windowID = ((PowerPoint.DocumentWindow)window).HWND;
//                     return windowID;
//                 }
//                 else if (host == HostApplication.Word)
//                 {
//                     int windowID = ((Word.Window)window).Hwnd;
//                     return windowID;
//                 }
//                 // else if (host == HostApplication.Visio)
//                 // {
//                 //     Int32 windowID = ((Visio.Window)window).WindowHandle32;
//                 //     return windowID;
//                 // }
//                 else
//                 {
//                     return 0;
//                 }
//             }
//             catch (Exception)
//             {
//                 return 0;
//             }
// #endif
//         }
//
//         #endregion

        
        
        
    }
}
