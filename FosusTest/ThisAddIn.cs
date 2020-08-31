using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using Office = Microsoft.Office.Core;
using System.Diagnostics;

namespace FosusTest
{
    public partial class ThisAddIn
    {
        private Ribbon1 ribbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.VisioIsIdle += Application_VisioIsIdle;
        }

        void Application_VisioIsIdle(Visio.Application app)
        {
            if (ribbon != null && ribbon.shouldInvalidate)
            {
                Debug.WriteLine("INVALIDATE");
                ribbon.InvalidateControl("dropDown1");
                ribbon.shouldInvalidate = false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon1();
            return ribbon;
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
}
