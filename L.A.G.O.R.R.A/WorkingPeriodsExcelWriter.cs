using ClosedXML.Excel;

namespace L.A.G.O.R.R.A
{
    public class WorkingPeriodsExcelWriter
    {
        private IXLWorksheet periodsSheet;
        private int currentRow;

        public WorkingPeriodsExcelWriter(IXLWorksheet periodsSheet)
        {
            this.periodsSheet = periodsSheet;
            currentRow = 1;
        }

        public void WritePeriodHeader(int workerId)
        {
            periodsSheet.Cell("A" + currentRow).Value = "ID:";
            periodsSheet.Cell("B" + currentRow).Value = workerId;
            periodsSheet.Cell("G" + currentRow).Value = "Mañana";
            periodsSheet.Cell("H" + currentRow).Value = "Tarde";
            
            currentRow++;
        }

        public void WriteBlankRow()
        {
            currentRow+=2;
        }

        public void WritePeriod(WorkerWorkingPeriod workingPeriod)
        {
            int firstDayRow = currentRow;
            
            foreach (WorkedDay workedDay in workingPeriod.getWorkedDays())
            {
                writeWorkDay(workedDay);                                        
                currentRow++;
            }
            
            int lastDayRow = currentRow - 1;
                
            periodsSheet.Cell("G" + currentRow).FormulaA1 = "=SUM(G" + firstDayRow + ":G" + lastDayRow + ")";
            periodsSheet.Cell("H" + currentRow).FormulaA1 = "=SUM(H" + firstDayRow + ":H" + lastDayRow + ")";
            periodsSheet.Cell("I" + currentRow).FormulaA1 = "=SUM(G" + currentRow + ": H" + currentRow + ")";
            
            periodsSheet.Cell("G" + currentRow).Style.NumberFormat.Format = "[h]:mm:ss";
            periodsSheet.Cell("H" + currentRow).Style.NumberFormat.Format = "[h]:mm:ss";
            periodsSheet.Cell("I" + currentRow).Style.NumberFormat.Format = "[h]:mm:ss";
        }
        
        private void writeWorkDay(WorkedDay workedDay)
        {
            periodsSheet.Cell("A" + currentRow).Value = workedDay.getDate();
            char lastColum = 'B';

            foreach (Time loggedTime in workedDay.getLoggedTimes())
            {
                //TODO:Search for a way to use last written column
                string loggedTimeCellIndex = lastColum.ToString() + currentRow;
                periodsSheet.Cell(loggedTimeCellIndex).Value = loggedTime.getTime();
                lastColum++;
            }

            periodsSheet.Cell("G" + currentRow).Value = workedDay.getMorningWorkedHours();
            periodsSheet.Cell("H" + currentRow).Value = workedDay.getAfternooWorkedHours();

            if ( workedDay.hasInvalidLoggedTimes() )
            {
                periodsSheet.Cell("G" + currentRow).Style.Fill.BackgroundColor = XLColor.Red;
                periodsSheet.Cell("H" + currentRow).Style.Fill.BackgroundColor = XLColor.Red;
            }

        }
    }
}