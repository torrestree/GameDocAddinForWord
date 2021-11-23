using GameDocAddinForWord.Common;
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

        public void BtnDesignDashboardTable_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.DashboardTable.Insert(application);
        }
        public void BtnDesignModelTable_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.ModelTable.Insert(application);
        }
        public void BtnDesignSaveTable_OnAction(Office.IRibbonControl control)
        {
            MessageBox.Show(Helpers.MsgUnderDeveloping);
        }
        public void BtnDesignEnumTable_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.EnumTable.Insert(application);
        }

        public void BtnDesignPropertyType_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyType.TryOverwrite(application);
        }
        public int CkbDesignPropertyType_GetItemCount(Office.IRibbonControl control)
        {
            return DesignDoc.PropertyType.Items.Count;
        }
        public string CkbDesignPropertyType_GetItemID(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertyType.Items[index].Id;
        }
        public string CkbDesignPropertyType_GetItemLabel(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertyType.Items[index].Label;
        }
        public int CkbDesignPropertyType_GetSelectedItemIndex(Office.IRibbonControl control)
        {
            return DesignDoc.PropertyType.SelectedIndex;
        }
        public void CkbDesignPropertyType_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            DesignDoc.PropertyType.SelectedIndex = selectedIndex;
        }

        public void BtnDesignPropertySource_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertySource.TryOverwrite(application);
        }
        public int CkbDesignPropertySource_GetItemCount(Office.IRibbonControl control)
        {
            return DesignDoc.PropertySource.Items.Count;
        }
        public string CkbDesignPropertySource_GetItemID(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertySource.Items[index].Id;
        }
        public string CkbDesignPropertySource_GetItemLabel(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertySource.Items[index].Label;
        }
        public int CkbDesignPropertySource_GetSelectedItemIndex(Office.IRibbonControl control)
        {
            return DesignDoc.PropertySource.SelectedIndex;
        }
        public void CkbDesignPropertySource_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            DesignDoc.PropertySource.SelectedIndex = selectedIndex;
        }

        public void BtnDesignSimpleProperty_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.SimpleProperty.TryOverwrite(application);
        }
        public void BtnDesignCompoundPropertyInt_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.CompoundProperty.TryOverwrite(application, true);
        }
        public void BtnDesignCompoundPropertyFloat_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.CompoundProperty.TryOverwrite(application, false);
        }

        public void BtnDevelopPropertyTable_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.ClassTable.Insert(application, DevelopDoc.ClassTable.TableTypes.Property);
        }
        public void BtnDevelopEventTable_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.ClassTable.Insert(application, DevelopDoc.ClassTable.TableTypes.Event);
        }
        public void BtnDevelopMethodTable_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.ClassTable.Insert(application, DevelopDoc.ClassTable.TableTypes.Method);
        }
        public void BtnDevelopEnumTable_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.EnumTable.Insert(application);
        }

        public void BtnDevelopPropertyType_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.PropertyType.TryOverwrite(application);
        }
        public int CkbDevelopPropertyType_GetItemCount(Office.IRibbonControl control)
        {
            return DevelopDoc.PropertyType.Items.Count;
        }
        public string CkbDevelopPropertyType_GetItemID(Office.IRibbonControl control, int index)
        {
            return DevelopDoc.PropertyType.Items[index].Id;
        }
        public string CkbDevelopPropertyType_GetItemLabel(Office.IRibbonControl control, int index)
        {
            return DevelopDoc.PropertyType.Items[index].Label;
        }
        public int CkbDevelopPropertyType_GetSelectedItemIndex(Office.IRibbonControl control)
        {
            return DevelopDoc.PropertyType.SelectedIndex;
        }
        public void CkbDevelopPropertyType_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            DevelopDoc.PropertyType.SelectedIndex = selectedIndex;
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
