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
    internal static class ValueType
    {
        public static ValueTypes SelectedType { get; set; }

        public static void Overwrite(Word.Application application)
        {
            if (application.GetRowIndex(2, out int rowIndex, out Word.Table table))
            {
                Word.Range range = table.Rows[rowIndex].Cells[2].Range;
                switch (SelectedType)
                {
                    case ValueTypes.Text: range.Text = "文本"; break;
                    case ValueTypes.Int: range.Text = "整型"; break;
                    case ValueTypes.Float: range.Text = "浮点"; break;
                    case ValueTypes.Bool: range.Text = "布尔"; break;
                    case ValueTypes.Enum: range.Text = "枚举"; break;
                    case ValueTypes.List: range.Text = "集合"; break;
                    default: break;
                }
            }
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }

        public enum ValueTypes
        {
            Text,
            Int,
            Float,
            Bool,
            Enum,
            List
        }
    }
}
