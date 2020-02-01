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

            MoveToNextRow();
        }

        public void WriteBlankRow()
        {
            MoveToNextRow();
        }

        public void WritePeriod(WorkerWorkingPeriod workingPeriod)
        {
            int firstDayRow = currentRow;
            int lastDayRow = currentRow;
            foreach (WorkedDay workedDay in workingPeriod.getWorkedDays())
            {
                writeWorkDay(workedDay);
                lastDayRow = currentRow;
                MoveToNextRow();
            }
            WriteWorkedTimeSum(firstDayRow, lastDayRow);
            MoveToNextRow();
        }

        private void WriteWorkedTimeSum(int firstDayRow, int lastDayRow)
        {
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
        
        private void MoveToNextRow()
        {
            currentRow++;
        }
    }
}