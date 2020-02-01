using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace L.A.G.O.R.R.A
{
    class Program
    {
        
        private static SortedDictionary<int, WorkerWorkingPeriod> GroupTimeEntriesInWorkingPeriodsByWorker(List<WorkerTimeEntry> timeEntries)
        {
            SortedDictionary<int, WorkerWorkingPeriod> workingPeriods = new SortedDictionary<int, WorkerWorkingPeriod>();
            foreach (var timeEntry in timeEntries)
            {
                int workerId = timeEntry.getWorkerID();
                if ( !workingPeriods.ContainsKey(workerId) )
                {
                    workingPeriods[workerId] = new WorkerWorkingPeriod(workerId);
                }
                workingPeriods[workerId].addWorkerTimeEntry(timeEntry);
            }

            return workingPeriods;
        }
        static void Main(string[] args)
        {
            string hoursWorkbookName = @"C:\Users\Juanma\Documents\L.A.G.O.R.R.A\Testing files\Test_file.xlsx";;
            int timeEntriesCount = 629;
            /*
            string hoursWorkbookName = args[0];
            int workerTimeEntriesCount = Utilities.StringToInt(args[1]);
            */

            var hoursWorkbook = new XLWorkbook(hoursWorkbookName);
            var timeEntries = new WorkerTimeEntriesExcelFileParser().parse(hoursWorkbook, timeEntriesCount);
            SortedDictionary<int, WorkerWorkingPeriod> workingPeriods = GroupTimeEntriesInWorkingPeriodsByWorker(timeEntries);
            
            var workingPeriodsExcelWriter = new WorkingPeriodsExcelWriter(hoursWorkbook.Worksheet(2));
            //TODO: A dictionary might not be needed here, use set instead
            foreach (KeyValuePair<int, WorkerWorkingPeriod> workingPeriod in workingPeriods)
            {
                int workerId = workingPeriod.Key;
                workingPeriodsExcelWriter.WritePeriodHeader(workerId);
                workingPeriodsExcelWriter.WritePeriod(workingPeriod.Value);
                workingPeriodsExcelWriter.WriteBlankRow();
            }

            hoursWorkbook.Save();
        }
        
    }
    
}
