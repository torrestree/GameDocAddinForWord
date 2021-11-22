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
            MessageBox.Show(Common.Helpers.MsgUnderDeveloping);
        }
        public void BtnDesignDocEnumTable_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.EnumTable.Insert(application);
        }

        public string BtnDesignDocType_GetLabel(Office.IRibbonControl control)
        {
            string label = "类型（type）";
            switch (DesignDoc.ValueType.SelectedType)
            {
                case DesignDoc.ValueType.ValueTypes.Text: return label.Replace("type", "文");
                case DesignDoc.ValueType.ValueTypes.Int: return label.Replace("type", "整");
                case DesignDoc.ValueType.ValueTypes.Float: return label.Replace("type", "浮");
                case DesignDoc.ValueType.ValueTypes.Bool: return label.Replace("type", "布");
                case DesignDoc.ValueType.ValueTypes.Enum: return label.Replace("type", "枚");
                case DesignDoc.ValueType.ValueTypes.List: return label.Replace("type", "集");
                default: return label.Replace("type", "");
            }
        }
        public void BtnDesignDocType_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.ValueType.TryOverwrite(application);
        }
        public bool CkbDesignDocType_GetPressed(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "CkbDesignDocTypeText": return DesignDoc.ValueType.SelectedType == DesignDoc.ValueType.ValueTypes.Text;
                case "CkbDesignDocTypeInt": return DesignDoc.ValueType.SelectedType == DesignDoc.ValueType.ValueTypes.Int;
                case "CkbDesignDocTypeFloat": return DesignDoc.ValueType.SelectedType == DesignDoc.ValueType.ValueTypes.Float;
                case "CkbDesignDocTypeBool": return DesignDoc.ValueType.SelectedType == DesignDoc.ValueType.ValueTypes.Bool;
                case "CkbDesignDocTypeEnum": return DesignDoc.ValueType.SelectedType == DesignDoc.ValueType.ValueTypes.Enum;
                case "CkbDesignDocTypeList": return DesignDoc.ValueType.SelectedType == DesignDoc.ValueType.ValueTypes.List;
                default: return false;
            }
        }
        public void CkbDesignDocType_OnAction(Office.IRibbonControl control, bool pressed)
        {
            switch (control.Id)
            {
                case "CkbDesignDocTypeText": DesignDoc.ValueType.SelectedType = DesignDoc.ValueType.ValueTypes.Text; break;
                case "CkbDesignDocTypeInt": DesignDoc.ValueType.SelectedType = DesignDoc.ValueType.ValueTypes.Int; break;
                case "CkbDesignDocTypeFloat": DesignDoc.ValueType.SelectedType = DesignDoc.ValueType.ValueTypes.Float; break;
                case "CkbDesignDocTypeBool": DesignDoc.ValueType.SelectedType = DesignDoc.ValueType.ValueTypes.Bool; break;
                case "CkbDesignDocTypeEnum": DesignDoc.ValueType.SelectedType = DesignDoc.ValueType.ValueTypes.Enum; break;
                case "CkbDesignDocTypeList": DesignDoc.ValueType.SelectedType = DesignDoc.ValueType.ValueTypes.List; break;
            }

            ribbon.InvalidateControl("CkbDesignDocTypeText");
            ribbon.InvalidateControl("CkbDesignDocTypeInt");
            ribbon.InvalidateControl("CkbDesignDocTypeFloat");
            ribbon.InvalidateControl("CkbDesignDocTypeBool");
            ribbon.InvalidateControl("CkbDesignDocTypeEnum");
            ribbon.InvalidateControl("CkbDesignDocTypeList");

            ribbon.InvalidateControl("BtnDesignDocType");

            DesignDoc.ValueType.TryOverwrite(application);
        }

        public string BtnDesignDocSource_GetLabel(Office.IRibbonControl control)
        {
            string label = "来源（type）";
            switch (DesignDoc.ValueSource.SelectedSource)
            {
                case DesignDoc.ValueSource.ValueSources.Save: return label.Replace("type", "存");
                case DesignDoc.ValueSource.ValueSources.Model: return label.Replace("type", "模");
                case DesignDoc.ValueSource.ValueSources.System: return label.Replace("type", "系");
                case DesignDoc.ValueSource.ValueSources.Input: return label.Replace("type", "入");
                default: return label.Replace("type", "");
            }
        }
        public void BtnDesignDocSource_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.ValueSource.TryOverwrite(application);
        }
        public bool CkbDesignDocSource_GetPressed(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "CkbDesignDocSourceSave": return DesignDoc.ValueSource.SelectedSource == DesignDoc.ValueSource.ValueSources.Save;
                case "CkbDesignDocSourceModel": return DesignDoc.ValueSource.SelectedSource == DesignDoc.ValueSource.ValueSources.Model;
                case "CkbDesignDocSourceSystem": return DesignDoc.ValueSource.SelectedSource == DesignDoc.ValueSource.ValueSources.System;
                case "CkbDesignDocSourceInput": return DesignDoc.ValueSource.SelectedSource == DesignDoc.ValueSource.ValueSources.Input;
                default: return false;
            }
        }
        public void CkbDesignDocSource_OnAction(Office.IRibbonControl control, bool pressed)
        {
            switch (control.Id)
            {
                case "CkbDesignDocSourceSave": DesignDoc.ValueSource.SelectedSource = DesignDoc.ValueSource.ValueSources.Save; break;
                case "CkbDesignDocSourceModel": DesignDoc.ValueSource.SelectedSource = DesignDoc.ValueSource.ValueSources.Model; break;
                case "CkbDesignDocSourceSystem": DesignDoc.ValueSource.SelectedSource = DesignDoc.ValueSource.ValueSources.System; break;
                case "CkbDesignDocSourceInput": DesignDoc.ValueSource.SelectedSource = DesignDoc.ValueSource.ValueSources.Input; break;
            }

            ribbon.InvalidateControl("CkbDesignDocSourceSave");
            ribbon.InvalidateControl("CkbDesignDocSourceModel");
            ribbon.InvalidateControl("CkbDesignDocSourceSystem");
            ribbon.InvalidateControl("CkbDesignDocSourceInput");

            ribbon.InvalidateControl("BtnDesignDocSource");

            DesignDoc.ValueSource.TryOverwrite(application);
        }

        public void BtnDesignDocSimpleProperty_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.SimpleProperty.TryOverwrite(application);
        }

        public void BtnDesignDocCompoundProperty_OnAction(Office.IRibbonControl control)
        {

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
