
using System;

// for Dictionary
using System.Collections.Generic;

// for Debug
using System.Diagnostics;

// for Path
using System.IO; 

// for python scripting hosting
using Microsoft.Scripting.Hosting;
using IronPython.Hosting;


namespace pyro
{
            
    public partial class AddIn
    {
        // The AddIn class is the main pyro class and implements all necessary interfaces
        // to interact with Office.
        // 
        // This part implents:
        //  - setup of IronPython Scripting Host
        //  - initialize and hold instance of Python addin
        //  - debugging methods for Python with pydev in Eclipse
        //

        // Python environment
        private object app;
        private ScriptEngine ipy;
        private ScriptScope scope;
        private dynamic python_delegate;
        private Context context;
        
        
        
        #region Python engine
        // ===================
        // = Python delegate =
        // ===================

        private void LoadPython() {
            Trace.TraceInformation("LoadPython Called");
    
            try {
                CreatePythonEngine();
                if(debug) {
                    TryLoadDebugger();
                }
            } catch (Exception e) {
                broken = true;
                Message(e.ToString());
            }
        }

        private void TryLoadDebugger() {
            try {
                Python.ImportModule(scope, "pydevd");
                ipy.Execute("pydevd.settrace(stdoutToServer=True, stderrToServer=True)", scope);
            } catch(Exception e) {
                debug = false;
                Message(e.ToString());
            }
        }

        public dynamic GetDelegate() {
            return python_delegate;
        }
        
        private void BootstrapAddIn() {
            Trace.TraceInformation("BootstrapAddIn called");
            Debug.Indent();
    
            if(config_ipy_addin_module == null) {
                broken = true;
                Debug.Unindent();
                return;
            }
    
            // import bootstrap module
            Trace.TraceInformation("import module");
            // Python.ImportModule(ipy, ipy_addin_module);
            ipy.ImportModule(config_ipy_addin_module);
            dynamic module = ipy.GetSysModule().GetVariable("modules").get(config_ipy_addin_module);
            Trace.TraceInformation("done. loaded bootstrap module: " + config_ipy_addin_module);
    
            // start addin on python side
            try {
                Trace.TraceInformation("create addin on python side");
                python_delegate = module.create_addin();
                Trace.TraceInformation("done.");
            } catch (Exception e) {
                Message(e.ToString());
            }

            if(python_delegate == null) {
                Debug.Unindent();
                throw new Exception("addin bootstrapper returned null");
            } else if(!created) {
                // FIXME: check for context==null instead
                Trace.TraceInformation("create context object");
                context = new Context(app, this, debug, hostAppName);
                Trace.TraceInformation("calling on_create");
                python_delegate.on_create(context);
                Trace.TraceInformation("done.");
            }
            Debug.Unindent();
        }


        private void CreatePythonEngine() {
            Trace.TraceInformation("CreatePythonEngine called");
            Debug.Indent();
            Trace.TraceInformation("get instance");
    
            // check python debugging
            var options = new Dictionary<string, object>();
            if(config_pydev_debug) {
                options["Frames"] = true;
                options["FullFrames"] = true;
            }
    
            // initialze scripting engine
            // IronPython will load from: <root>\bin\
            Trace.TraceInformation("create scripting engine");
            ipy = Python.CreateEngine(options);
            Trace.TraceInformation("create scope");
            scope = ipy.CreateScope();
    
            // Initialize system path for Python
            // Python will load modules from <root> and <root>\Lib
            // where <root> is the directory of IronPython.dll
            Trace.TraceInformation("add ironpython paths");
            ICollection<string> paths = ipy.GetSearchPaths();
            paths.Clear();
            paths.Add(config_ironpython_root_full);
            paths.Add(Path.Combine(config_ironpython_root_full, "Lib"));
    
            if(config_ipy_addin_path != null) {
                Trace.TraceInformation("addind addin_path to sys.path: " + config_ipy_addin_path_full );
                paths.Add(config_ipy_addin_path_full);
            }
            
            // add debug path
            if(config_pydev_debug) {
                Trace.TraceInformation("add pydev-codebase (debug)");
                if(config_pydev_codebase == null) {
                    Message("debugging enabled, but pydev_codebase not set");
                } else {
                    debug = true;
                    paths.Add(config_pydev_codebase);
                }
            }
    
            ipy.SetSearchPaths(paths);
            Debug.Unindent();
            Trace.TraceInformation("CreatePythonEngine done.");
        }
        #endregion

    }
}