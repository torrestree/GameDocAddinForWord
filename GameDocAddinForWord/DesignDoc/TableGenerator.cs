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
    internal static class TableGenerator
    {
        public static void Insert(Word.Application application, TableTypes tableType)
        {
            switch (tableType)
            {
                case TableTypes.Dashboard: InsertDashboardTable(application); break;
                case TableTypes.Model: InsertModelTable(application); break;
                case TableTypes.Save: InsertSaveTable(application); break;
                case TableTypes.Enum: InsertEnumTable(application); break;
            }
            application.Selection.MoveDown();
        }
        private static void InsertDashboardTable(Word.Application application)
        {
            Word.Table table = application.InsertTable(4);

            table.Cell(1, 1).Range.Text = "属性";
            table.Cell(1, 2).Range.Text = "类型";
            table.Cell(1, 3).Range.Text = "作用";
            table.Cell(1, 4).Range.Text = "来源";

            table.Columns[1].SetTableColumnWidth(15);
            table.Columns[2].SetTableColumnWidth(15);
        }
        private static void InsertModelTable(Word.Application application)
        {
            Word.Table table = application.InsertTable(3);

            table.Cell(1, 1).Range.Text = "属性";
            table.Cell(1, 2).Range.Text = "类型";
            table.Cell(1, 3).Range.Text = "说明";

            table.Columns[1].SetTableColumnWidth(10);
            table.Columns[2].SetTableColumnWidth(10);
        }
        private static void InsertSaveTable(Word.Application application)
        {
            Word.Table table = application.InsertTable(4);

            table.Cell(1, 1).Range.Text = "属性";
            table.Cell(1, 2).Range.Text = "类型";
            table.Cell(1, 3).Range.Text = "作用";
            table.Cell(1, 4).Range.Text = "初始";

            table.Columns[1].SetTableColumnWidth(15);
            table.Columns[2].SetTableColumnWidth(15);
        }
        private static void InsertEnumTable(Word.Application application)
        {
            Word.Table table = application.InsertTable(2);

            table.Cell(1, 1).Range.Text = "枚举";
            table.Cell(1, 2).Range.Text = "说明";
        }

        public enum TableTypes
        {
            Dashboard,
            Model,
            Save,
            Enum
        }
    }
}
