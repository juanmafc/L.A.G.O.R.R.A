using System;
using System.Collections.Generic;


namespace L.A.G.O.R.R.A
{
    public class WorkerWorkingPeriod
    {

        private List<WorkerTimeEntry> workerTimeEntries = new List<WorkerTimeEntry>();
        public int WorkerID { get; }
        
        //TODO: do thise a sorted set by WorkedDay
        private SortedDictionary<DateTime, WorkedDay> workedDaysMap = new SortedDictionary<DateTime, WorkedDay>();

        public WorkerWorkingPeriod(int workerID)
        {
            this.WorkerID = workerID;
        }

        public void addWorkerTimeEntry(WorkerTimeEntry workerTimeEntry)
        {
            this.workerTimeEntries.Add(workerTimeEntry);

            var day = workerTimeEntry.getDay();
            if (!workedDaysMap.ContainsKey(day))
            {
                workedDaysMap[day] = new WorkedDay(day);
            }

            workedDaysMap[day].addLoggedTime( new Time(workerTimeEntry.getRoundedTime()) );

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

            return this.WorkerID == other.WorkerID;
        }

        public override int GetHashCode()
        {
            return this.WorkerID;
        }

        /*
        public int getTotalMorningWorkedHours()
        {            
            Time totalMorningWorkedHours;
            foreach (var workedDay in this.getWorkedDays())
            {
                totalMorningWorkedHours.addTime(workedDay.getMorningWorkedHours());
            }
            return totalMorningWorkedHours.toString();
        }
        */
    }
}