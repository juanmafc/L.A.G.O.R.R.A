using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace L.A.G.O.R.R.A
{
    class Program
    {
        
        private static ICollection<WorkerWorkingPeriod> GroupTimeEntriesInWorkingPeriodsByWorker(ICollection<WorkerTimeEntry> timeEntries)
        {
            var workingPeriods = new SortedDictionary<int, WorkerWorkingPeriod>();
            foreach (var timeEntry in timeEntries)
            {
                int workerId = timeEntry.getWorkerID();
                if ( !workingPeriods.ContainsKey(workerId) )
                {
                    //TODO: define what to do if the workerIDs do not match, maybe only timestamps should be added to workingPeriods
                    workingPeriods[workerId] = new WorkerWorkingPeriod(workerId);
                }
                workingPeriods[workerId].addWorkerTimeEntry(timeEntry);
            }
            return workingPeriods.Values;
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
            
            var workingPeriods = GroupTimeEntriesInWorkingPeriodsByWorker(timeEntries);
            
            var workingPeriodsExcelWriter = new WorkingPeriodsExcelWriter(hoursWorkbook.Worksheet(2));
            foreach (WorkerWorkingPeriod workingPeriod in workingPeriods)
            {
                workingPeriodsExcelWriter.WritePeriodHeader(workingPeriod.WorkerID);
                workingPeriodsExcelWriter.WritePeriod(workingPeriod);
                workingPeriodsExcelWriter.WriteBlankRow();
            }
            hoursWorkbook.Save();
        }
    }
    
}
