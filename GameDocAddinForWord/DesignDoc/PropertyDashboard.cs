using GameDocAddinForWord.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class PropertyDashboard
    {
        public static List<ComboBoxItemInfo> TypeItems { get; set; }
        public static int SelectedTypeIndex { get; set; }
        public static string SelectedTypeLabel
        {
            get { return TypeItems[SelectedTypeIndex].Label; }
        }

        public static List<ComboBoxItemInfo> SourceItems { get; set; }
        public static int SelectedSourceIndex { get; set; }
        public static string SelectedSourceLabel
        {
            get { return SourceItems[SelectedSourceIndex].Label; }
        }

        public static void Init()
        {
            TypeItems = new List<ComboBoxItemInfo>
            {
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardTypeText", Label = "文本" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardTypeInt", Label = "整型" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardTypeFloat", Label = "浮点" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardTypeBool", Label = "布尔" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardTypeEnum", Label = "枚举" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardTypeList", Label = "集合" }
            };

            SourceItems = new List<ComboBoxItemInfo>
            {
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardSourceSave", Label = "存档" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardSourceModel", Label = "模型" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardSourceSystem", Label = "系统参数" },
                new ComboBoxItemInfo { Id = "CbiDesignPropertyDashboardSourceInput", Label = "外部传入" }
            };
        }

        public static void TryOverwrite(Word.Application application, OverwriteTypes overwriteType)
        {
            if (PropertySource.CanOverwrite(application, out int rowIndex, out Word.Table table))
            {
                switch (overwriteType)
                {
                    case OverwriteTypes.Simple: OverwriteSimpleProperty(rowIndex, table); break;
                    case OverwriteTypes.CompoundInt: OverwriteCompoundProperty(rowIndex, table, true); break;
                    case OverwriteTypes.CompoundFloat: OverwriteCompoundProperty(rowIndex, table, false); break;
                }
            }
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }
        private static void OverwriteSimpleProperty(int rowIndex, Word.Table table)
        {
            WriteRow(table.Rows[rowIndex], "", SelectedTypeLabel, SelectedSourceLabel);
        }
        private static void OverwriteCompoundProperty(int rowIndex, Word.Table table, bool isInt)
        {
            string type = default;
            if (isInt)
                type = "整型";
            else
                type = "浮点";
            WriteRow(table.Rows[rowIndex], "", "复合数值", "");
            WriteRow(table.Rows.Add(), ">基础值", type, SelectedSourceLabel);
            WriteRow(table.Rows.Add(), ">额外值", type, "额外值公式");
            WriteRow(table.Rows.Add(), ">完整值", type, "完整值公式");
        }
        private static void WriteRow(Word.Row row, string name, string type, string source)
        {
            row.Cells[1].Range.Text = name;
            row.Cells[2].Range.Text = type;
            row.Cells[4].Range.Text = source;
        }

        public enum OverwriteTypes
        {
            Simple,
            CompoundInt,
            CompoundFloat
        }
    }
}
