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
            var hoursWorkbook = new XLWorkbook("Horas-Enero - copia.xlsx");           
            var hoursWorkSheet = hoursWorkbook.Worksheet(1);
            /*
            var workedHours = hoursWorkSheet.Cell("C4").Value;
            Console.Write(workedHours);            
            */

            var workerRegisteredEntriesExcelFileParser = new WorkerRegisteredEntriesExcelFileParser();
            var workerRegisteredEntries = workerRegisteredEntriesExcelFileParser.parse(hoursWorkbook);



            SortedDictionary<int, WorkerWorkingPeriod> workerWorkingPeriods = new SortedDictionary<int, WorkerWorkingPeriod>();
            foreach (var workerRegisteredEntry in workerRegisteredEntries)
            {
                int workerID = workerRegisteredEntry.getWorkerID();
                if ( !workerWorkingPeriods.ContainsKey(workerID) )
                {
                    workerWorkingPeriods[workerID] = new WorkerWorkingPeriod(workerID);
                }
                workerWorkingPeriods[workerID].addWorkerRegisteredEntry(workerRegisteredEntry);
            }


            var workersEntriesByDatesWorkSheet = hoursWorkbook.Worksheet(2);
            //TODO: rename this variable
            int workerDateCurrentRow = 1; 
            //TODO: A dictionary might not be needed here, use set instead
            foreach (KeyValuePair<int, WorkerWorkingPeriod> workingPeriod in workerWorkingPeriods)
            {
                writeWorkPeriodHeader(workingPeriod.Key, workersEntriesByDatesWorkSheet, workerDateCurrentRow);
                workerDateCurrentRow++;

                foreach (WorkedDay workedDay in workingPeriod.Value.getWorkedDays())
                {
                    writeWorkDay(workedDay, workersEntriesByDatesWorkSheet, workerDateCurrentRow);                                        
                    workerDateCurrentRow++;
                }


                workerDateCurrentRow++;

                workersEntriesByDatesWorkSheet.Cell("G" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalMorningWorkedHours();
                workersEntriesByDatesWorkSheet.Cell("H" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalAfternooWorkedHours();
                workersEntriesByDatesWorkSheet.Cell("I" + workerDateCurrentRow).Value = workingPeriod.Value.getTotalWorkedHours();

                workerDateCurrentRow++;
            }

            int row = 4;
            foreach (var workerRegisteredEntry in workerRegisteredEntries )
            {
                hoursWorkSheet.Cell("I" + row).Value = workerRegisteredEntry.getRoundedTime();
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
