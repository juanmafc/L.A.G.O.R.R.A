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
        static void Main(string[] args)
        {

            string hoursWorkbookName = @"C:\Users\Juanma\Documents\L.A.G.O.R.R.A\Testing files\Test_file.xlsx";;
            int workerTimeEntriesCount = 629;
            /*
            string hoursWorkbookName = args[0];
            int workerTimeEntriesCount = Utilities.StringToInt(args[1]);
            */

            var hoursWorkbook = new XLWorkbook(hoursWorkbookName);
            var hoursWorkSheet = hoursWorkbook.Worksheet(1);
            /*
            var workedHours = hoursWorkSheet.Cell("C4").Value;
            Console.Write(workedHours);            
            */

            var workerTimeEntriesExcelFileParser = new WorkerTimeEntriesExcelFileParser();
            var workerTimeEntries = workerTimeEntriesExcelFileParser.parse(hoursWorkbook, workerTimeEntriesCount);



            SortedDictionary<int, WorkerWorkingPeriod> workerWorkingPeriods = new SortedDictionary<int, WorkerWorkingPeriod>();
            foreach (var workerTimeEntry in workerTimeEntries)
            {
                int workerID = workerTimeEntry.getWorkerID();
                if ( !workerWorkingPeriods.ContainsKey(workerID) )
                {
                    workerWorkingPeriods[workerID] = new WorkerWorkingPeriod(workerID);
                }
                workerWorkingPeriods[workerID].addWorkerTimeEntry(workerTimeEntry);
            }


            var workersEntriesByDatesWorkSheet = hoursWorkbook.Worksheet(2);
            //TODO: rename this variable
            int workerDateCurrentRow = 1; 
            //TODO: A dictionary might not be needed here, use set instead
            foreach (KeyValuePair<int, WorkerWorkingPeriod> workingPeriod in workerWorkingPeriods)
            {
                writeWorkPeriodHeader(workingPeriod.Key, workersEntriesByDatesWorkSheet, workerDateCurrentRow);
                workerDateCurrentRow++;

                int firstDateRow = workerDateCurrentRow;
                foreach (WorkedDay workedDay in workingPeriod.Value.getWorkedDays())
                {
                    writeWorkDay(workedDay, workersEntriesByDatesWorkSheet, workerDateCurrentRow);                                        
                    workerDateCurrentRow++;
                }

                int lastDateRow = workerDateCurrentRow - 1;
                workersEntriesByDatesWorkSheet.Cell("G" + workerDateCurrentRow).FormulaA1 = "=SUM(G" + firstDateRow + ":G" + lastDateRow + ")";
                workersEntriesByDatesWorkSheet.Cell("H" + workerDateCurrentRow).FormulaA1 = "=SUM(H" + firstDateRow + ":H" + lastDateRow + ")";
                workersEntriesByDatesWorkSheet.Cell("I" + workerDateCurrentRow).FormulaA1 = "=SUM(G" + workerDateCurrentRow + ": H" + workerDateCurrentRow + ")";

                //TODO:refactor
                workersEntriesByDatesWorkSheet.Cell("G" + workerDateCurrentRow).Style.NumberFormat.Format = "[h]:mm:ss";
                workersEntriesByDatesWorkSheet.Cell("H" + workerDateCurrentRow).Style.NumberFormat.Format = "[h]:mm:ss";
                workersEntriesByDatesWorkSheet.Cell("I" + workerDateCurrentRow).Style.NumberFormat.Format = "[h]:mm:ss";


                /*TODO: total hours can be calculated here by adding all hours (better) or by using Excel's function (worse)
                workersEntriesByDatesWorkSheet.Cell("G" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalMorningWorkedHours();
                workersEntriesByDatesWorkSheet.Cell("H" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalAfternooWorkedHours();
                workersEntriesByDatesWorkSheet.Cell("I" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalWorkedHours();
                */
                workerDateCurrentRow +=2;
            }

            int row = 4;
            foreach (var workerTimeEntry in workerTimeEntries )
            {
                hoursWorkSheet.Cell("I" + row).Value = workerTimeEntry.getRoundedTime();
                row++;
            }

            hoursWorkbook.Save();
        }

        static void writeWorkDay(WorkedDay workedDay, IXLWorksheet workSheet, int row)
        {
            workSheet.Cell("A" + row).Value = workedDay.getDate();
            char lastColum = 'B';

            foreach (LoggedTime loggedTime in workedDay.getLoggedTimes())
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

        static void writeWorkPeriodHeader(int workerId, IXLWorksheet workSheet, int row)
        {
            workSheet.Cell("A" + row).Value = "ID:";
            workSheet.Cell("B" + row).Value = workerId;
            workSheet.Cell("G" + row).Value = "Mañana";
            workSheet.Cell("H" + row).Value = "Tarde";
        }
    }
}
