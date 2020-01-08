using NUnit.Framework;

namespace L.A.G.O.R.R.A.Test
{
    [TestFixture]
    public class TimeTests
    {
        [Test]
        public void GivenTwoTimes_WhenGettingTimeDifference_TheAbsoluteDifferenceInHoursAndMinutesIsReturned()
        {
            Time time1 = new Time("19:00:00");
            Time time2 = new Time("7:00:00");
            
            Assert.AreEqual("12:00", time1.getTimeDifference(time2));
            Assert.AreEqual("12:00", time2.getTimeDifference(time1));
        }
    }
}