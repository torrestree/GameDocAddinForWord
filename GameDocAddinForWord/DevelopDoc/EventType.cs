using GameDocAddinForWord.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GameDocAddinForWord.DevelopDoc
{
    internal static class EventType
    {
        public static void TryOverwrite(Word.Application application, EventTypes eventType)
        {
            if (application.GetRowIndex(2, out int rowIndex, out Word.Table table))
                table.Rows[rowIndex].Cells[2].Range.Text = eventType.ToString();
            else
                MessageBox.Show(Helpers.MsgUnmatchedTable);
        }

        public enum EventTypes
        {
            Action,
            Func
        }
    }
}
