using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace L.A.G.O.R.R.A
{
    class WorkerRegisteredEntry
    {
        private int workerID;
        private DateTime entryTimestamp;

        public WorkerRegisteredEntry(int workerID, DateTime entryTimestamp)
        {
            this.workerID = workerID;
            this.entryTimestamp = entryTimestamp;
        }

        public string getRoundedTime()
        {
            int hours = this.entryTimestamp.Hour;
            int minutes = this.entryTimestamp.Minute;

            //Redundant ranges are explicitely written to have a clearer code
            if ((0 <= minutes) && (minutes < 15))
            {
                minutes = 0;
            }
            else if ((15 <= minutes) && (minutes <= 45))
            {
                minutes = 30;
            }
            else if ((45 < minutes) && (minutes <= 59))
            {
                hours++;
                minutes = 0;
            }

            return hours + ":" + minutes;
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
