using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class CompoundProperty
    {
        public static void TryOverwrite(Word.Application application, bool isIntType)
        {
            string type;
            if (isIntType)
                type = "整型";
            else
                type = "浮点";
            if (PropertySource.CanOverwrite(application, out int rowIndex, out Word.Table table))
            {
                WriteRow(table.Rows[rowIndex], "", "复合数值", "");
                WriteRow(table.Rows.Add(), ">基础值", type, PropertySource.SelectedLabel);
                WriteRow(table.Rows.Add(), ">额外值", type, "额外值公式");
                WriteRow(table.Rows.Add(), ">完整值", type, "完整值公式");
            }
        }
        private static void WriteRow(Word.Row row, string name, string type, string source)
        {
            row.Cells[1].Range.Text = name;
            row.Cells[2].Range.Text = type;
            row.Cells[4].Range.Text = source;
        }
    }
}
