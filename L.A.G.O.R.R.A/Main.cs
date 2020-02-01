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

            var workingPeriodsSheet = hoursWorkbook.Worksheet(2);
            //TODO: rename this variable
            int workerDateCurrentRow = 1; 
            //TODO: A dictionary might not be needed here, use set instead
            foreach (KeyValuePair<int, WorkerWorkingPeriod> workingPeriod in workingPeriods)
            {
                WriteWorkingPeriodHeader(workingPeriod.Key, workingPeriodsSheet, workerDateCurrentRow);
                workerDateCurrentRow++;

                int firstDateRow = workerDateCurrentRow;
                foreach (WorkedDay workedDay in workingPeriod.Value.getWorkedDays())
                {
                    writeWorkDay(workedDay, workingPeriodsSheet, workerDateCurrentRow);                                        
                    workerDateCurrentRow++;
                }

                int lastDateRow = workerDateCurrentRow - 1;
                workingPeriodsSheet.Cell("G" + workerDateCurrentRow).FormulaA1 = "=SUM(G" + firstDateRow + ":G" + lastDateRow + ")";
                workingPeriodsSheet.Cell("H" + workerDateCurrentRow).FormulaA1 = "=SUM(H" + firstDateRow + ":H" + lastDateRow + ")";
                workingPeriodsSheet.Cell("I" + workerDateCurrentRow).FormulaA1 = "=SUM(G" + workerDateCurrentRow + ": H" + workerDateCurrentRow + ")";

                //TODO:refactor
                workingPeriodsSheet.Cell("G" + workerDateCurrentRow).Style.NumberFormat.Format = "[h]:mm:ss";
                workingPeriodsSheet.Cell("H" + workerDateCurrentRow).Style.NumberFormat.Format = "[h]:mm:ss";
                workingPeriodsSheet.Cell("I" + workerDateCurrentRow).Style.NumberFormat.Format = "[h]:mm:ss";


                /*TODO: total hours can be calculated here by adding all hours (better) or by using Excel's function (worse)
                workersEntriesByDatesWorkSheet.Cell("G" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalMorningWorkedHours();
                workersEntriesByDatesWorkSheet.Cell("H" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalAfternooWorkedHours();
                workersEntriesByDatesWorkSheet.Cell("I" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalWorkedHours();
                */
                workerDateCurrentRow +=2;
            }


            hoursWorkbook.Save();
        }

       

        static void writeWorkDay(WorkedDay workedDay, IXLWorksheet workSheet, int row)
        {
            workSheet.Cell("A" + row).Value = workedDay.getDate();
            char lastColum = 'B';

            foreach (Time loggedTime in workedDay.getLoggedTimes())
            {
                //TODO:Search for a way to use last written column
                string loggedTimeCellIndex = lastColum.ToString() + row;
                workSheet.Cell(loggedTimeCellIndex).Value = loggedTime.getTime();
                lastColum++;
            }

            workSheet.Cell("G" + row).Value = workedDay.getMorningWorkedHours();
            workSheet.Cell("H" + row).Value = workedDay.getAfternooWorkedHours();

            if ( workedDay.hasInvalidLoggedTimes() )
            {
                workSheet.Cell("G" + row).Style.Fill.BackgroundColor = XLColor.Red;
                workSheet.Cell("H" + row).Style.Fill.BackgroundColor = XLColor.Red;
            }

        }

        static void WriteWorkingPeriodHeader(int workerId, IXLWorksheet workSheet, int row)
        {
            workSheet.Cell("A" + row).Value = "ID:";
            workSheet.Cell("B" + row).Value = workerId;
            workSheet.Cell("G" + row).Value = "Mañana";
            workSheet.Cell("H" + row).Value = "Tarde";
        }
    }
}
