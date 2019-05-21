
using Microsoft.Office.Core;

namespace pyro
{
    /// <summary>
    /// Description of Class1.
    /// </summary>
    public class Context
	{	
		public object app {get; private set; }
		public AddIn addin {get; private set; }
		public bool debug {get; private set; }
		public IRibbonUI ribbon {get; internal set; }
		public string hostAppName {get; private set; }
		
		public Context(object app, AddIn addin, bool debug, string hostAppName)
		{
			this.app = app;
			this.addin = addin;
			this.debug = debug;
			this.hostAppName = hostAppName;
		}
	}
}
