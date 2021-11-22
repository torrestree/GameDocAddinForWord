using GameDocAddinForWord.Common;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class ModelTable
    {
        public static void Insert(Word.Application application)
        {
            Word.Table table = application.InsertTable(3);
            SetTitleRow(table);
            SetColumnWidth(table);
            application.Selection.MoveDown();
        }
        private static void SetTitleRow(Word.Table table)
        {
            table.Cell(1, 1).Range.Text = "属性";
            table.Cell(1, 2).Range.Text = "类型";
            table.Cell(1, 3).Range.Text = "说明";
        }
        private static void SetColumnWidth(Word.Table table)
        {
            table.Columns[1].SetTableColumnWidth(10);
            table.Columns[2].SetTableColumnWidth(10);
        }
    }
}
