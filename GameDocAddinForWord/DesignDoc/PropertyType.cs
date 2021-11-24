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
    internal static class PropertyType
    {
        public static void TryOverwrite(Word.Application application, PropertyTypes propertyType)
        {
            if (CanOverwrite(application, out int rowIndex, out Word.Table table))
            {
                string value = default;
                switch (propertyType)
                {
                    case PropertyTypes.Text: value = "文本"; break;
                    case PropertyTypes.Int: value = "整型"; break;
                    case PropertyTypes.Float: value = "浮点"; break;
                    case PropertyTypes.Bool: value = "布尔"; break;
                    case PropertyTypes.Enum: value = "枚举"; break;
                    case PropertyTypes.List: value = "集合"; break;
                }
                Overwrite(rowIndex, table, value);
            }
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }
        public static bool CanOverwrite(Word.Application application, out int rowIndex, out Word.Table table)
        {
            return application.GetRowIndex(2, out rowIndex, out table);
        }
        public static void Overwrite(int rowIndex, Word.Table table, string value)
        {
            table.Rows[rowIndex].Cells[2].Range.Text = value;
        }

        public enum PropertyTypes
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
