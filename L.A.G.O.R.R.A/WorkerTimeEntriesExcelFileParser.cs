using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace L.A.G.O.R.R.A
{
    class WorkerTimeEntriesExcelFileParser
    {
        public List<WorkerTimeEntry> parse(XLWorkbook file, int entriesCount)
        {
            var workerTimeEntriesWorksheet = file.Worksheet(1);
            List<WorkerTimeEntry> workerTimeEntries = new List<WorkerTimeEntry>();
            //Start at C4, iterate all the way up to C632            
            //for (int row = 4; row <= 632; row++)
            for (int row = 1; row <= entriesCount; row++) 
            {
                int workerID = Utilities.StringToInt( workerTimeEntriesWorksheet.Cell("A" + row).GetString() );
                DateTime timestamp = workerTimeEntriesWorksheet.Cell("B" + row).GetDateTime();

                workerTimeEntries.Add( new WorkerTimeEntry(workerID, timestamp) );
            }

            return workerTimeEntries;
        }
    }
}
