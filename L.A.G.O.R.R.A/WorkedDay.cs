using System;
using System.Collections.Generic;

namespace L.A.G.O.R.R.A
{
    public class WorkedDay
    {
        private DateTime day;
        private List<Time> loggedTimes = new List<Time>();

        public WorkedDay(DateTime workedDay)
        {
            this.day = workedDay;
        }

        public string getDate()
        {
            return String.Format("{0:dd/MM/yyyy}", this.day);
        }

        public IEnumerable<Time> getLoggedTimes()
        {
            return this.loggedTimes;
        }

        internal void addLoggedTime(Time time)
        {
            this.loggedTimes.Add(time);
        }

        public string getMorningWorkedHours()
        {
            if ( this.loggedTimes.Count < 2)
            {
                return "?????";
            }            

            return this.loggedTimes[1].getTimeDifference(this.loggedTimes[0]);
        }

        public string getAfternooWorkedHours()
        {
            if (this.loggedTimes.Count < 4)
            {
                return "?????";
            }

            return this.loggedTimes[3].getTimeDifference(this.loggedTimes[2]);
        }

        public bool hasInvalidLoggedTimes()
        {
            return this.loggedTimes.Count != 4;
        }
    }
}