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

        public void BtnDesignTableDashboard_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.TableGenerator.Insert(application, DesignDoc.TableGenerator.TableTypes.Dashboard);
        }
        public void BtnDesignTableModel_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.TableGenerator.Insert(application, DesignDoc.TableGenerator.TableTypes.Model);
        }
        public void BtnDesignTableSave_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.TableGenerator.Insert(application, DesignDoc.TableGenerator.TableTypes.Save);
        }
        public void BtnDesignTableEnum_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.TableGenerator.Insert(application, DesignDoc.TableGenerator.TableTypes.Enum);
        }

        public void BtnDesignPropertyTypeText_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyType.TryOverwrite(application, DesignDoc.PropertyType.PropertyTypes.Text);
        }
        public void BtnDesignPropertyTypeInt_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyType.TryOverwrite(application, DesignDoc.PropertyType.PropertyTypes.Int);
        }
        public void BtnDesignPropertyTypeFloat_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyType.TryOverwrite(application, DesignDoc.PropertyType.PropertyTypes.Float);
        }
        public void BtnDesignPropertyTypeBool_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyType.TryOverwrite(application, DesignDoc.PropertyType.PropertyTypes.Bool);
        }
        public void BtnDesignPropertyTypeEnum_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyType.TryOverwrite(application, DesignDoc.PropertyType.PropertyTypes.Enum);
        }
        public void BtnDesignPropertyTypeList_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyType.TryOverwrite(application, DesignDoc.PropertyType.PropertyTypes.List);
        }

        public void BtnDesignPropertySourceSave_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertySource.TryOverwrite(application, DesignDoc.PropertySource.PropertySources.Save);
        }
        public void BtnDesignPropertySourceModel_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertySource.TryOverwrite(application, DesignDoc.PropertySource.PropertySources.Model);
        }
        public void BtnDesignPropertySourceSystem_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertySource.TryOverwrite(application, DesignDoc.PropertySource.PropertySources.System);
        }
        public void BtnDesignPropertySourceInput_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertySource.TryOverwrite(application, DesignDoc.PropertySource.PropertySources.Input);
        }

        public int CbbDesignPropertyDashboardType_GetItemCount(Office.IRibbonControl control)
        {
            return DesignDoc.PropertyDashboard.TypeItems.Count;
        }
        public string CbbDesignPropertyDashboardType_GetItemId(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertyDashboard.TypeItems[index].Id;
        }
        public string CbbDesignPropertyDashboardType_GetItemLabel(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertyDashboard.TypeItems[index].Label;
        }
        public int CbbDesignPropertyDashboardType_GetSelectedItemIndex(Office.IRibbonControl control)
        {
            return DesignDoc.PropertyDashboard.SelectedTypeIndex;
        }
        public void CbbDesignPropertyDashboardType_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            DesignDoc.PropertyDashboard.SelectedTypeIndex = selectedIndex;
        }

        public int CbbDesignPropertyDashboardSource_GetItemCount(Office.IRibbonControl control)
        {
            return DesignDoc.PropertyDashboard.SourceItems.Count;
        }
        public string CbbDesignPropertyDashboardSource_GetItemId(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertyDashboard.SourceItems[index].Id;
        }
        public string CbbDesignPropertyDashboardSource_GetItemLabel(Office.IRibbonControl control, int index)
        {
            return DesignDoc.PropertyDashboard.SourceItems[index].Label;
        }
        public int CbbDesignPropertyDashboardSource_GetSelectedItemIndex(Office.IRibbonControl control)
        {
            return DesignDoc.PropertyDashboard.SelectedSourceIndex;
        }
        public void CbbDesignPropertyDashboardSource_OnAction(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            DesignDoc.PropertyDashboard.SelectedSourceIndex = selectedIndex;
        }

        public void BtnDesignPropertyDashboardSimple_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyDashboard.TryOverwrite(application, DesignDoc.PropertyDashboard.OverwriteTypes.Simple);
        }

        public void BtnDesignPropertyDashboardCompoundInt_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyDashboard.TryOverwrite(application, DesignDoc.PropertyDashboard.OverwriteTypes.CompoundInt);
        }
        public void BtnDesignPropertyDashboardCompoundFloat_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.PropertyDashboard.TryOverwrite(application, DesignDoc.PropertyDashboard.OverwriteTypes.CompoundFloat);
        }

        public void BtnDevelopTableProperty_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.TableGenerator.Insert(application, DevelopDoc.TableGenerator.TableTypes.Property);
        }
        public void BtnDevelopTableEvent_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.TableGenerator.Insert(application, DevelopDoc.TableGenerator.TableTypes.Event);
        }
        public void BtnDevelopTableMethod_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.TableGenerator.Insert(application, DevelopDoc.TableGenerator.TableTypes.Method);
        }
        public void BtnDevelopTableEnum_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.TableGenerator.Insert(application, DevelopDoc.TableGenerator.TableTypes.Enum);
        }

        public void BtnDevelopPropertyTypeString_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.PropertyType.TryOverwrite(application, DevelopDoc.PropertyType.PropertyTypes.String);
        }
        public void BtnDevelopPropertyTypeInt_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.PropertyType.TryOverwrite(application, DevelopDoc.PropertyType.PropertyTypes.Int);
        }
        public void BtnDevelopPropertyTypeFloat_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.PropertyType.TryOverwrite(application, DevelopDoc.PropertyType.PropertyTypes.Float);
        }
        public void BtnDevelopPropertyTypeBool_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.PropertyType.TryOverwrite(application, DevelopDoc.PropertyType.PropertyTypes.Bool);
        }
        public void BtnDevelopPropertyTypeEnum_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.PropertyType.TryOverwrite(application, DevelopDoc.PropertyType.PropertyTypes.Enum);
        }
        public void BtnDevelopPropertyTypeList_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.PropertyType.TryOverwrite(application, DevelopDoc.PropertyType.PropertyTypes.List);
        }

        public void BtnDevelopEventTypeAction_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.EventType.TryOverwrite(application, DevelopDoc.EventType.EventTypes.Action);
        }
        public void BtnDevelopEventTypeFunc_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.EventType.TryOverwrite(application, DevelopDoc.EventType.EventTypes.Func);
        }

        public void BtnDevelopMethodTypeVoid_OnAction(Office.IRibbonControl control)
        {
            DevelopDoc.MethodType.TryOverwrite(application);
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
