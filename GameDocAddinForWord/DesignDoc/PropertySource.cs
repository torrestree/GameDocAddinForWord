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
    internal static class PropertySource
    {
        public static void TryOverwrite(Word.Application application, PropertySources propertySource)
        {
            if (CanOverwrite(application, out int rowIndex, out Word.Table table))
            {
                string value = default;
                switch (propertySource)
                {
                    case PropertySources.Save: value = "存档"; break;
                    case PropertySources.Model: value = "模型"; break;
                    case PropertySources.System: value = "系统参数"; break;
                    case PropertySources.Input: value = "外部传入"; break;
                }
                Overwrite(rowIndex, table, value);
            }
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }
        public static bool CanOverwrite(Word.Application application, out int rowIndex, out Word.Table table)
        {
            return application.GetRowIndex(4, out rowIndex, out table);
        }
        public static void Overwrite(int rowIndex, Word.Table table, string value)
        {
            table.Rows[rowIndex].Cells[4].Range.Text = value;
        }

        public enum PropertySources
        {
            Save,
            Model,
            System,
            Input
        }
    }
}
