using GameDocAddinForWord.Common;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DevelopDoc
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
                new CbxItemInfo { Id = "DevelopTypeString", Label = "String" },
                new CbxItemInfo { Id = "DevelopTypeInt", Label = "Int" },
                new CbxItemInfo { Id = "DevelopTypeFloat", Label = "Float" },
                new CbxItemInfo { Id = "DevelopTypeBool", Label = "Bool" },
                new CbxItemInfo { Id = "DevelopTypeAction", Label = "Action" },
                new CbxItemInfo { Id = "DevelopTypeFunc", Label = "Func" },
                new CbxItemInfo { Id = "DevelopTypeVoid", Label = "Void" }
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
