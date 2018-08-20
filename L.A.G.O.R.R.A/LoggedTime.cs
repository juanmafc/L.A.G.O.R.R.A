using System;

namespace L.A.G.O.R.R.A
{
    class LoggedTime
    {
        private string time;

        public LoggedTime(string time)
        {
            this.time = time;
        }

        public string getTime()
        {
            return this.time;
        }

        public string getTimeDifference(LoggedTime other)
        {
            TimeSpan myLoggedTimeSpan = TimeSpan.Parse(this.time);
            TimeSpan otherLoggedTimeSpan = TimeSpan.Parse(other.time);

            TimeSpan timeDifference =  myLoggedTimeSpan.Subtract(otherLoggedTimeSpan);
            return timeDifference.Hours + ":" + timeDifference.Minutes;
        }
    }
}