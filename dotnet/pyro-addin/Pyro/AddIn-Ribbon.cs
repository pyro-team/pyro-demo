
using System;
using System.Diagnostics;

// Path, ...
using System.IO; 

// Assembly
using System.Reflection;


// ResourceManager
using System.Resources; 

// IRibbonExtensibility
using Microsoft.Office.Core;

// Bitmap
using System.Drawing;

// xml validation
using System.Xml;
using System.Xml.Schema;
using System.Xml.Linq;



namespace pyro
{
            
    public partial class AddIn : IRibbonExtensibility
    {
        // The AddIn class is the main pyro class and implements all necessary interfaces
        // to interact with Office.
        // 
        // This part implents:
        //  - the IRibbonExtensibility interface to handle ribbon events and pass them to the python addin
        //  - methods to load and return custom-ui (xml-string) for the ribbon
        //
        
        
        public string GetCustomUI(string ribbon_id) {
            Trace.TraceInformation("GetCustomUI called");
            return GetPythonCustomUIAndWriteToFile(ribbon_id);
        }
        
        private string GetPythonCustomUIAndWriteToFile(string ribbon_id) {
            if(broken) {
                return "";
            }
            try {
                string customUI = python_delegate.get_custom_ui(ribbon_id);
                //Trace.TraceInformation(customUI);
                string filename = GetFilenameFromRibbonId(ribbon_id);
                System.IO.StreamWriter file = new System.IO.StreamWriter(filename);
                file.WriteLine(customUI);
                file.Close();
                Trace.TraceInformation("wrote xml in: " + filename);
                return customUI;
            } catch (Exception e) {
                broken = true;
                Message(e.ToString());
                return "";
            }
        }
        
        private string GetCustomUIFromFile(string ribbon_id) {
            string filename = GetFilenameFromRibbonId(ribbon_id);
            Trace.TraceInformation("loading xml from: " + filename);
            string customUI;
            if (File.Exists(filename)) {
                StreamReader streamReader = new StreamReader(filename);
                customUI = streamReader.ReadToEnd();
                streamReader.Close();
                //Trace.TraceInformation(customUI);
            } else {
                Trace.TraceInformation("file not found.");
                customUI = "";
            }
            
            return customUI;
        }
        
        private string GetFilenameFromRibbonId(string ribbon_id) {
            // IronPythonLoader loader = IronPythonLoader.GetInstance();
            string dirname = config_ipy_addin_path_full + "\\resources\\xml\\";
            Directory.CreateDirectory(dirname);
            return dirname + ribbon_id + ".xml";            
        }
        
        
        public void VerifyCustomUI(string text)
        {
            ResourceManager resources = new ResourceManager("pyro.xml_schemata", Assembly.GetExecutingAssembly());
            byte[] xsd_data = (byte[]) resources.GetObject("customui14_xsd");
            Stream xsd_res = new MemoryStream(xsd_data);
            
            XmlSchemaSet ss = new XmlSchemaSet();
            ss.Add("http://schemas.microsoft.com/office/2009/07/customui", XmlReader.Create(xsd_res));
                // File.OpenRead(@"C:\Office 2010 Developer Resources\Schemas\customui14.xsd"))
            
            xsd_res.Close();
            
            XDocument doc = XDocument.Parse(text);
            Console.WriteLine("Validating XML");
            doc.Validate(ss, new ValidationEventHandler(ValidationCallBack));
        }

        private void ValidationCallBack(object sender, ValidationEventArgs vea)
        {    
            Console.WriteLine("\t event sender: " + sender);
            Console.WriteLine("\t validation args: " + vea);
            Console.WriteLine("");
        }
        
        
        
        
        
        
        #region Callbacks with Python delegation
        // ========================================
        // = Python Events: information callbacks =
        // ========================================
        // For callback signatures see
        //  https://msdn.microsoft.com/en-us/library/aa722523(v=office.12).aspx
        //  https://msdn.microsoft.com/en-us/library/bb736142(v=office.12).aspx
        
        public string PythonGetContent(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetContent " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_content(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public string PythonGetDescription(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetDescription " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_description(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetEnabled(IRibbonControl control)
        {
            Trace.TraceInformation("event GetEnabled " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_enabled(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
         
        public Bitmap PythonGetImage(IRibbonControl control) {
            Trace.TraceInformation("event GetImage " + control.Id);
            if (!created) return null;
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_image(control);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetKeytip(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetKeytip " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_keytip(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public string PythonGetLabel(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetLabel " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_label(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetPressed(IRibbonControl control)
        {
            Trace.TraceInformation("event GetPressed " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_pressed(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        public string PythonGetScreentip(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetScreentip " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_screentip(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetShowImage(IRibbonControl control)
        {
            Trace.TraceInformation("event GetShowImage " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_show_image(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        public bool PythonGetShowLabel(IRibbonControl control)
        {
            Trace.TraceInformation("event GetShowLabel " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_show_label(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        public int PythonGetSize(IRibbonControl control)
        {
            Trace.TraceInformation("event GetSize " + control.Id);
            if (!created) return 0;
            try {
                var result = python_delegate.get_size(control);
                if(result == "large") {
                    return 1;
                } else {
                    return 0;
                }
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        public string PythonGetSupertip(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetSupertip " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_supertip(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public string PythonGetText(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetText " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_text(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }

        public string PythonGetTitle(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetTitle " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_title(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        
        public bool PythonGetVisible(IRibbonControl control)
        {
            Trace.TraceInformation("event GetVisible " + control.Id);
            if (!created) return false;
            try {
                return (python_delegate.get_visible(control) == true);
            } catch (Exception e) {
                Message(e.ToString());
                return false;
            }
        }
        
        
        // ====================================
        // = Python Events: gallery/combo box =
        // ====================================
        
        public int PythonGetItemCount(IRibbonControl control) {
            Trace.TraceInformation("event GetItemCount " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                var v = python_delegate.get_item_count(control);
                if (v==null) {
                    return 0;
                } else {
                    return v;
                }
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }

        public int PythonGetSelectedItemIndex(IRibbonControl control) {
            Trace.TraceInformation("event GetSelectedItemIndex " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                return python_delegate.get_selected_item_index(control);
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        public string PythonGetSelectedItemID(IRibbonControl control)
        {    
            Trace.TraceInformation("event GetSelectedItemID " + control.Id);
            if (!created) return "";
            try {
                var result = python_delegate.get_selected_item_id(control);
                if(result == null) {
                    return "";
                } else {
                    return result.ToString();
                }
            } catch (Exception e) {
                Message(e.ToString());
                return "";
            }
        }
        

        // ==============================================
        // = Python Events: gallery/combo box (indexed) =
        // ==============================================
        
        public int PythonGetItemHeight(IRibbonControl control) {
            Trace.TraceInformation("event GetItemHeight " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                return python_delegate.get_item_height(control);
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        public string PythonGetItemID(IRibbonControl control, int index) {
            Trace.TraceInformation("event GetItemID " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_id(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public Bitmap PythonGetItemImage(IRibbonControl control, int index) {
        //public stdole.IPictureDisp GetItemImage(IRibbonControl oRbnCtrl, int iItemIndex)
            Trace.TraceInformation("event GetItemImage " + control.Id);
            if (!created) return null;
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_image(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetItemLabel(IRibbonControl control, int index) {
            Trace.TraceInformation("event GetItemLabel " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_label(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetItemScreentip(IRibbonControl control, int index) {
            Trace.TraceInformation("event GetItemScreentip " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_screentip(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public string PythonGetItemSupertip(IRibbonControl control, int index) {
            Trace.TraceInformation("event GetItemSupertip " + control.Id);
            if (!created) return "";
            if(broken) {
                return null;
            }
            try {
                return python_delegate.get_item_supertip(control, index);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public int PythonGetItemWidth(IRibbonControl control) {
            Trace.TraceInformation("event GetItemWidth " + control.Id);
            if (!created) return 0;
            if(broken) {
                return 0;
            }
            try {
                return python_delegate.get_item_width(control);
            } catch (Exception e) {
                Message(e.ToString());
                return 0;
            }
        }
        
        
        // ===================================
        // = Python Events: action callbacks =
        // ===================================
        
        public void PythonOnAction(IRibbonControl control)
        {    
            Trace.TraceInformation("event OnAction " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_action(control);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnActionRepurposed(IRibbonControl control, ref bool cancelDefault)
        {    
            Trace.TraceInformation("event OnActionRepurposed " + control.Id);
            if (!created) return;
            try {
                cancelDefault = Convert.ToBoolean(python_delegate.on_action_repurposed(control));
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnActionIndexed(IRibbonControl control, string selectedItem, int index)
        {    
            Trace.TraceInformation("event OnActionIndex " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_action_indexed(control, selectedItem, index);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnToggleAction(IRibbonControl control, bool pressed)
        {    
            Trace.TraceInformation("event OnToggleAction " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_toggle_action(control, pressed);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        public void PythonOnChange(IRibbonControl control, string value)
        {    
            Trace.TraceInformation("event OnChange " + control.Id);
            if (!created) return;
            try {
                python_delegate.on_change(control, value);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        
        
        // ======================================
        // = Python Events: image / ribbon load =
        // ======================================
        
        public Bitmap PythonLoadImage(string image_name) {
            Trace.TraceInformation("event LoadImage " + image_name);
            if (!created) return null;
            if(broken) {
                return null;
            }
            try {
                return python_delegate.load_image(image_name);
            } catch (Exception e) {
                Message(e.ToString());
                return null;
            }
        }
        
        public void PythonOnRibbonLoad(IRibbonUI ui)
        {
            Trace.TraceInformation("event OnRibbonLoad");
            if (!created) 
            {
                return;
            }
            try {
                context.ribbon = ui;
                python_delegate.on_ribbon_load(ui);
            } catch (Exception e) {
                Message(e.ToString());
            }
        }
        #endregion

        
        #region Fast enabled events
        // =======================
        // = Fast Enabled-Events =
        // =======================
        
        private int selection_type = 0;
        private int selection_shapes = 0;
        private bool selection_containstextframe = false;

        public bool GetEnabled_True(IRibbonControl control)
        {
            Trace.TraceInformation("event GetEnabled_True " + control.Id);
            if (!created) return false;
            return true;
        }

        public bool GetEnabled_Ppt_ShapesOrText(IRibbonControl control)
        {
            Trace.TraceInformation("event GetEnabled_Ppt_ShapesOrText " + control.Id);
            if (!created) return false;
            return selection_type == 2 || selection_type == 3;
        }

        public bool GetEnabled_Ppt_Shapes_ExactOne(IRibbonControl control)
        {
            Trace.TraceInformation("event GetEnabled_Ppt_Shapes_ExactOne " + control.Id);
            if (!created) return false;
            return selection_shapes == 1;
        }

        public bool GetEnabled_Ppt_Shapes_ExactTwo(IRibbonControl control)
        {
            Trace.TraceInformation("event GetEnabled_Ppt_Shapes_ExactTwo " + control.Id);
            if (!created) return false;
            return selection_shapes == 2;
        }

        public bool GetEnabled_Ppt_Shapes_MinTwo(IRibbonControl control)
        {
            Trace.TraceInformation("event GetEnabled_Ppt_Shapes_MinTwo " + control.Id);
            if (!created) return false;
            return selection_shapes >= 2;
        }

        public bool GetEnabled_Ppt_ContainsTextFrame(IRibbonControl control)
        {
            Trace.TraceInformation("event GetEnabled_Ppt_ContainsTextFrame " + control.Id);
            if (!created) return false;
            return selection_containstextframe;
        }
        
        /*
        Enabled based on selection
            0 = ppSelectionNone
            1 = ppSelectionSlide
            2 = ppSelectionShape
            3 = ppSelectionText
        */
        // public Boolean GetEnabled_Shapes_Selected(IRibbonControl control)
        // {
        //             Trace.TraceInformation("event GetEnabled_Selection_Available " + control.Id);
        //             if (!created) return false;
        //             return ((PowerPoint.Application)app.ActiveWindow.selection.Type == 2);
        // }
        //
        // public Boolean GetEnabled_Text_Selected(IRibbonControl control)
        // {
        //             Trace.TraceInformation("event GetEnabled_Selection_Available " + control.Id);
        //             if (!created) return false;
        //             return ((PowerPoint.Application)app.ActiveWindow.selection.Type == 3);
        // }

        #endregion
        
    }
}