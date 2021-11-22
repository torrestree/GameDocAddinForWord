using GameDocAddinForWord.Common;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class PropertyType
    {
        public static List<CbxItemInfo> Items { get; set; }
        public static int SelectedIndex { get; set; }
        public static string SelectedLabel
        {
            get { return Items[SelectedIndex].Label; }
        }

        public static void Init()
        {
            Items = new List<CbxItemInfo>
            {
                new CbxItemInfo(){ Id = "DesignTypeText", Label = "文本" },
                new CbxItemInfo(){ Id = "DesignTypeInt", Label = "整型" },
                new CbxItemInfo(){ Id = "DesignTypeFloat", Label = "浮点" },
                new CbxItemInfo(){ Id = "DesignTypeBool", Label = "布尔" },
                new CbxItemInfo(){ Id = "DesignTypeEnum", Label = "枚举" },
                new CbxItemInfo(){ Id = "DesignTypeList", Label = "集合" }
            };
        }

        public static void TryOverwrite(Word.Application application)
        {
            if (CanOverwrite(application, out int rowIndex, out Word.Table table))
                Overwrite(rowIndex, table);
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }
        public static bool CanOverwrite(Word.Application application, out int rowIndex, out Word.Table table)
        {
            return application.GetRowIndex(2, out rowIndex, out table);
        }
        public static void Overwrite(int rowIndex, Word.Table table)
        {
            table.Rows[rowIndex].Cells[2].Range.Text = SelectedLabel;
        }
    }
}
