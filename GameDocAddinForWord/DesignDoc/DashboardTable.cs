using GameDocAddinForWord.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DesignDoc
{
    internal static class DashboardTable
    {
        public static void Insert(Word.Application application)
        {
            Word.Table table = application.CreateTable(4);
            SetTitleRow(table);
            SetColumnWidth(table);
            application.Selection.MoveDown();
        }
        private static void SetTitleRow(Word.Table table)
        {
            table.Cell(1, 1).Range.Text = "属性";
            table.Cell(1, 2).Range.Text = "类型";
            table.Cell(1, 3).Range.Text = "作用";
            table.Cell(1, 4).Range.Text = "来源";
        }
        private static void SetColumnWidth(Word.Table table)
        {
            table.Columns[1].SetTableColumnWidth(15);
            table.Columns[2].SetTableColumnWidth(15);
        }
    }
}
