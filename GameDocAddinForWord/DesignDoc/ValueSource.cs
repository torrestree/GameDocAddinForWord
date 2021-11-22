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
    internal static class ValueSource
    {
        public static ValueSources SelectedSource { get; set; }

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
            Word.Range range = table.Rows[rowIndex].Cells[4].Range;
            switch (SelectedSource)
            {
                case ValueSources.Save: range.Text = "存档"; break;
                case ValueSources.Model: range.Text = "模型"; break;
                case ValueSources.System: range.Text = "系统参数"; break;
                case ValueSources.Input: range.Text = "外部输入"; break;
                default: break;
            }
        }

        public enum ValueSources
        {
            Save,
            Model,
            System,
            Input
        }
    }
}
