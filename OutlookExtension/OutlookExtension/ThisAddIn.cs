using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookExtension
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Make sure no exception in our addin should crash Outlook
            try {
                Core core = Core.Instance;
                core.Initialize();
            } catch (Exception) {
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }

    public class Core {
        private static Core _instance;
        public static Core Instance {
            get {
                if (_instance == null)
                    _instance = new Core();
                return _instance;
            }
        }

        private Core() {

        }

        public void Initialize() {
            // Execute implant here
            MessageBox.Show("MyOutlook addin has been loaded!");
            System.Diagnostics.Process.Start("calc.exe");
        }
    }
}
