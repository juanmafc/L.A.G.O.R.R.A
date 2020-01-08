using System;
using NUnit.Framework;

namespace L.A.G.O.R.R.A.Test
{
    [TestFixture]
    public class WorkerTimeEntryTests
    {

        private DateTime CreateEntryTimestamp(int hour, int minutes)
        {
            return new DateTime(2020, 12, 25, hour, minutes, 0);
        }
        
        private WorkerTimeEntry CreateWorkerTimeEntry(DateTime entryTimestamp)
        {
            const int workerId = 0;
            return new WorkerTimeEntry(workerId, entryTimestamp);
        }
        
        public WorkerTimeEntryTestingClass GivenAWorkerTimeEntryAt(int hour, int minutes)
        {
            var entryTimestamp = CreateEntryTimestamp(hour, minutes);
            var entry = CreateWorkerTimeEntry(entryTimestamp);
            return new WorkerTimeEntryTestingClass(entry);
        }

        public class WorkerTimeEntryTestingClass
        {
            private WorkerTimeEntry entry;

            public WorkerTimeEntryTestingClass(WorkerTimeEntry entry)
            {
                this.entry = entry;
            }


            public void TheRoundedTimeShouldBe(string roundedTime)
            {
                Assert.AreEqual(roundedTime, entry.getRoundedTime());
            }
        }
        
        
        [Test]
        public void GivenATimeEntry_WhenRoundingTime_ThenTheExpectedRoundedTimeShouldBeReturned()
        {
            GivenAWorkerTimeEntryAt(7, 0).TheRoundedTimeShouldBe("7:00");
            GivenAWorkerTimeEntryAt(7, 1).TheRoundedTimeShouldBe("7:00");
            GivenAWorkerTimeEntryAt(7, 14).TheRoundedTimeShouldBe("7:00");
            
            GivenAWorkerTimeEntryAt(7, 15).TheRoundedTimeShouldBe("7:30");
            GivenAWorkerTimeEntryAt(7, 44).TheRoundedTimeShouldBe("7:30");
            GivenAWorkerTimeEntryAt(7, 45).TheRoundedTimeShouldBe("7:30");
            
            GivenAWorkerTimeEntryAt(7, 46).TheRoundedTimeShouldBe("8:00");
            GivenAWorkerTimeEntryAt(7, 59).TheRoundedTimeShouldBe("8:00");
            GivenAWorkerTimeEntryAt(8, 0).TheRoundedTimeShouldBe("8:00");
        }
    }
}