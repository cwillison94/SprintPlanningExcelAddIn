using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace HelloWorld
{
    public class TeamMember
    {
        private const int PlannedHoursOffset = 6;
        public TeamMember(int row, int column, string username)
        {
            this.Row = row;
            this.Column = column;
            this.Username = username;
        }

        public int Row { get; private set; }
        public int Column { get; private set; }
        public string Username { get; private set; }
        public double PlannedHours { get; set; }

        public void UpdatePlannedHours(Excel.Worksheet worksheet)
        {
            var range = (Excel.Range)worksheet.Cells[Column + PlannedHoursOffset][Row];
            range.Value = PlannedHours;
        }

    }
}
