﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.Common
{
    internal static class Helpers
    {
        public static Word.Table CreateTable(this Word.Application application, int columns)
        {
            Word.Range range = application.Selection.Range;
            Word.Table table = application.ActiveDocument.Tables.Add(range, 2, columns);
            string styleName = ((Word.Style)application.ActiveDocument.DefaultTableStyle).NameLocal;
            if (!string.IsNullOrEmpty(styleName))
                table.set_Style(styleName);
            table.Rows[1].HeadingFormat = (int)Word.WdConstants.wdToggle;
            table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100;
            return table;
        }
        public static void SetTableColumnWidth(this Word.Column column, float width)
        {
            column.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            column.PreferredWidth = width;
        }
        public static void InsertText(this Word.Selection selection, string text)
        {
            selection.InsertAfter(text);
            selection.MoveRight();
        }
        public static void OverwriteText(this Word.Selection selection, string text)
        {
            selection.Text = text;
            selection.MoveRight();
        }
    }
}
