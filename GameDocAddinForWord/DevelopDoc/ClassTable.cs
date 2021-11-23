using GameDocAddinForWord.Common;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DevelopDoc
{
    internal static class ClassTable
    {
        public static void Insert(Word.Application application, TableTypes type)
        {
            Word.Table table = application.InsertTable(3);
            SetTitleRow(table, type);
            SetColumnWidth(table);
            application.Selection.MoveDown();
        }
        private static void SetTitleRow(Word.Table table, TableTypes type)
        {
            switch (type)
            {
                case TableTypes.Property: table.Cell(1, 1).Range.Text = "属性"; break;
                case TableTypes.Event: table.Cell(1, 1).Range.Text = "事件"; break;
                case TableTypes.Method: table.Cell(1, 1).Range.Text = "方法"; break;
            }
            table.Cell(1, 2).Range.Text = "类型";
            table.Cell(1, 3).Range.Text = "说明";
        }
        private static void SetColumnWidth(Word.Table table)
        {
            table.Columns[1].SetTableColumnWidth(30);
            table.Columns[2].SetTableColumnWidth(30);
        }

        public enum TableTypes
        {
            Property,
            Event,
            Method
        }
    }
}
