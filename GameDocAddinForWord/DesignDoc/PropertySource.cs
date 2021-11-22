using GameDocAddinForWord.Common;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class PropertySource
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
                new CbxItemInfo(){ Id = "DesignSourceSave", Label = "存档" },
                new CbxItemInfo(){ Id = "DesignSourceModel", Label = "模型" },
                new CbxItemInfo(){ Id = "DesignSourceSystem", Label = "系统参数" },
                new CbxItemInfo(){ Id = "DesignSourceInput", Label = "参数传入" }
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
            return application.GetRowIndex(4, out rowIndex, out table);
        }
        public static void Overwrite(int rowIndex, Word.Table table)
        {
            table.Rows[rowIndex].Cells[4].Range.Text = SelectedLabel;
        }
    }
}
