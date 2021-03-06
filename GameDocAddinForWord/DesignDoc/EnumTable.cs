using GameDocAddinForWord.Common;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class EnumTable
    {
        public static void Insert(Word.Application application)
        {
            Word.Table table = application.InsertTable(2);
            SetTitleRow(table);
            application.Selection.MoveDown();
        }
        private static void SetTitleRow(Word.Table table)
        {
            table.Cell(1, 1).Range.Text = "枚举";
            table.Cell(1, 2).Range.Text = "说明";
        }
    }
}
