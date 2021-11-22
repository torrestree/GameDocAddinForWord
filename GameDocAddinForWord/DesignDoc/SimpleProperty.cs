using GameDocAddinForWord.Common;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class SimpleProperty
    {
        public static void TryOverwrite(Word.Application application)
        {
            if (PropertySource.CanOverwrite(application, out int rowIndex, out Word.Table table))
            {
                PropertyType.Overwrite(rowIndex, table);
                PropertySource.Overwrite(rowIndex, table);
            }
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }
    }
}
