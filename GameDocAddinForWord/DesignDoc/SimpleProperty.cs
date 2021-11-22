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
    internal static class SimpleProperty
    {
        public static void TryOverwrite(Word.Application application)
        {
            bool canOverwriteType = ValueType.CanOverwrite(application, out _, out _);
            bool canOverwriteSource = ValueSource.CanOverwrite(application, out int rowIndex, out Word.Table table);
            if (canOverwriteType && canOverwriteSource)
            {
                ValueType.Overwrite(rowIndex, table);
                ValueSource.Overwrite(rowIndex, table);
            }
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }
    }
}
