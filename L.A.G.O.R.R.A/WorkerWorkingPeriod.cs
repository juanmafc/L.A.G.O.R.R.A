using System;
using System.Collections.Generic;


namespace L.A.G.O.R.R.A
{
    class WorkerWorkingPeriod
    {

        private List<WorkerRegisteredEntry> workerRegisteredEntries = new List<WorkerRegisteredEntry>();
        private int workerID;
        
        //TODO: do thise a sorted set by WorkedDay
        private SortedDictionary<DateTime, WorkedDay> workedDaysMap = new SortedDictionary<DateTime, WorkedDay>();

        public WorkerWorkingPeriod(int workerID)
        {
            this.workerID = workerID;
        }

        public void addWorkerRegisteredEntry(WorkerRegisteredEntry workerRegisteredEntry)
        {
            this.workerRegisteredEntries.Add(workerRegisteredEntry);

            var day = workerRegisteredEntry.getDay();
            if (!workedDaysMap.ContainsKey(day))
            {
                workedDaysMap[day] = new WorkedDay(day);
            }

            workedDaysMap[day].addLoggedTime( new LoggedTime(workerRegisteredEntry.getRoundedTime()) );

        }



        public IEnumerable<WorkedDay> getWorkedDays()
        {
            List<WorkedDay> workedDays = new List<WorkedDay>();
            foreach (var workedDay in workedDaysMap)
            {
                workedDays.Add(workedDay.Value);
            }
            return workedDays;
        }




        public override bool Equals(object obj)
        {
            WorkerWorkingPeriod other = obj as WorkerWorkingPeriod;
            if (other == null) return false;

            return this.workerID == other.workerID;
        }

        public override int GetHashCode()
        {
            return this.workerID;
        }

        public int getTotalMorningWorkedHours()
        {
            Time totalMorningWorkedHours;
            foreach (var workedDay in this.getWorkedDays())
            {
                totalMorningWorkedHours.addTime(workedDay.getMorningWorkedHours());
            }
            return totalMorningWorkedHours.toString();
        }
    }
}