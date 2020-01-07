using System;

namespace L.A.G.O.R.R.A
{
    public class Time
    {
        private string time;

        public Time(string time)
        {
            this.time = time;
        }

        public string getTime()
        {
            return this.time;
        }

        public string getTimeDifference(Time other)
        {
            TimeSpan myLoggedTimeSpan = TimeSpan.Parse(this.time);
            TimeSpan otherLoggedTimeSpan = TimeSpan.Parse(other.time);

            TimeSpan timeDifference =  myLoggedTimeSpan.Subtract(otherLoggedTimeSpan);
            return timeDifference.ToString(@"hh\:mm");
        }
    }
}