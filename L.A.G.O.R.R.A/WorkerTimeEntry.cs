using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace L.A.G.O.R.R.A
{
    public class WorkerTimeEntry
    {
        private int workerID;
        private DateTime entryTimestamp;

        public WorkerTimeEntry(int workerID, DateTime entryTimestamp)
        {
            this.workerID = workerID;
            this.entryTimestamp = entryTimestamp;
        }

        public string getRoundedTime()
        {
            int hours = this.entryTimestamp.Hour;
            int minutes = this.entryTimestamp.Minute;

            //Redundant ranges are explicitly written to have a clearer code
            if ((0 <= minutes) && (minutes < 15)) return hours + ":00";
            if ((15 <= minutes) && (minutes <= 45)) return hours + ":30";
            //(45 < minutes) && (minutes <= 59)
            hours++;
            return hours + ":00";
        }

        public DateTime getDay()
        {
            return entryTimestamp.Date;
        }

        internal int getWorkerID()
        {
            return this.workerID;
        }
    }
}
