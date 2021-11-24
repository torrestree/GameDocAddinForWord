using GameDocAddinForWord.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DevelopDoc
{
    internal static class TableGenerator
    {
        public static void Insert(Word.Application application, TableTypes tableType)
        {
            switch (tableType)
            {
                case TableTypes.Property: InsertPropertyTable(application); break;
                case TableTypes.Event: InsertEventTable(application); break;
                case TableTypes.Method: InsertMethodTable(application); break;
                case TableTypes.Enum: InsertEnumTable(application); break;
            }
            application.Selection.MoveDown();
        }
        private static void InsertPropertyTable(Word.Application application)
        {
            InsertClassTable(application, "属性");
        }
        private static void InsertEventTable(Word.Application application)
        {
            InsertClassTable(application, "事件");
        }
        private static void InsertMethodTable(Word.Application application)
        {
            InsertClassTable(application, "方法");
        }
        private static void InsertEnumTable(Word.Application application)
        {
            Word.Table table = application.InsertTable(2);

            table.Cell(1, 1).Range.Text = "枚举";
            table.Cell(1, 2).Range.Text = "说明";
        }

        private static void InsertClassTable(Word.Application application, string header)
        {
            Word.Table table = application.InsertTable(3);

            table.Cell(1, 1).Range.Text = header;
            table.Cell(1, 2).Range.Text = "类型";
            table.Cell(1, 3).Range.Text = "说明";

            table.Columns[1].SetTableColumnWidth(30);
            table.Columns[2].SetTableColumnWidth(30);
        }

        public enum TableTypes
        {
            Property,
            Event,
            Method,
            Enum
        }
    }
}
