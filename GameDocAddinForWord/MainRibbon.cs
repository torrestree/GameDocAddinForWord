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

            DesignDoc.ValueType.Overwrite(application);
        }
        public void BtnDesignDocType_OnAction(Office.IRibbonControl control)
        {
            DesignDoc.ValueType.Overwrite(application);
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
