using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;


namespace FosusTest
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        string[] data = { "1", "2", "3", "4", "5" };
        int selectedIndex;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("FosusTest.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public string GetText(IRibbonControl control)
        {
            if (control.Id == "editBox1")
            {
                return "1";
            }
            else
            {
                return "2";
            }
        }

        public void EditBoxTextChange(IRibbonControl control, string text)
        {
            for (var i = 0; i < data.Count(); ++i)
            {
                if (data[i] == text)
                {
                    selectedIndex = i;
                    ribbon.InvalidateControl("dropDown1");
                    break;
                }
            }
        }

        public int GetSelectedItemIndex(IRibbonControl control)
        {
            return selectedIndex;
        }

        public int GetItemCount(IRibbonControl control)
        {
            return data.Count();
        }

        public string GetItemLabel(IRibbonControl control, int index)
        {
            return data[index];
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
