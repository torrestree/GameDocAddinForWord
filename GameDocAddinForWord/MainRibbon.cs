using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord
{
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private Word.Application application;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("GameDocAddinForWord.MainRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
            application = Globals.ThisAddIn.Application;
        }

        public void BtnDesignDocDashboardTable_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.DashboardTable.Insert(application);
        }
        public void BtnDesignDocModelTable_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.ModelTable.Insert(application);
        }
        public void BtnDesignDocSaveTable_OnAction(Office.IRibbonControl control)
        {
            MessageBox.Show("功能尚待开发！");
        }
        public void BtnDesignDocEnumTable_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.EnumTable.Insert(application);
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
