using DevExpress.XtraScheduler;
using System;
using System.Linq;


namespace Сессия.Classes
{
    #region #mytimescaleday
    public class MyTimeScaleDay : TimeScaleFixedInterval
    {
        TimeSpan start = TimeSpan.FromHours(8);
        TimeSpan end = TimeSpan.FromHours(22);

        public MyTimeScaleDay(TimeSpan span) : base(span)
        {
            this.Width = 65;
        }

        public override DateTime Floor(DateTime date)
        {
            DateTime result = date.Date;
            if (result.DayOfWeek == DayOfWeek.Monday && date.TimeOfDay < start)
                result = result.AddDays(-3);
            else if (date.TimeOfDay < start)
                result = result.AddDays(-1);
            return result + start;
        }
        public override string FormatCaption(DateTime start, DateTime end)
        {
            return start.ToShortDateString();
        }
        protected override TimeSpan SortingWeight
        {
            get { return TimeSpan.FromDays(1); }
        }
    }
    #endregion #mytimescaleday

    #region #MyTimeScaleLessThanDay
    public class MyTimeScaleLessThanDay : TimeScaleFixedInterval
    {
        readonly TimeSpan StartTimeLimitation = TimeSpan.FromHours(8);
        readonly TimeSpan EndTimeLimitation = TimeSpan.FromHours(22);
        TimeSpan start { get { return StartTimeLimitation; } }
        TimeSpan end { get { return EndTimeLimitation - Value; } }
        TimeSpan duration;

        public MyTimeScaleLessThanDay(TimeSpan span) : base(span)
        {
            duration = span;
        }

        protected override string DefaultMenuCaption { get { return "8:00-22:00"; } }

        protected override TimeSpan SortingWeight
        {
            get { return duration; }
        }

        public override DateTime Floor(DateTime date)
        {
            if (date == DateTime.MinValue || date == DateTime.MaxValue)
                return date;

            if (date.TimeOfDay < start)
                // Round down to the end of the previous working day.
                return RoundToVisibleIntervalEdge(date.AddDays(-1), end);

            if (date.TimeOfDay > end)
                // Round down to the end of the current working day.
                return RoundToVisibleIntervalEdge(date, end);

            return base.Floor(date);
        }

        protected DateTime RoundToVisibleIntervalEdge(DateTime dateTime, TimeSpan time)
        {
            return dateTime.Date + time;
        }

        protected override bool HasNextDate(DateTime date)
        {
            return date <= RoundToVisibleIntervalEdge(date, end);
        }

        public override DateTime GetNextDate(DateTime date)
        {
            return (date.TimeOfDay >= end) ? RoundToVisibleIntervalEdge(date.AddDays(1), start) : base.GetNextDate(date);
        }
    }
    #endregion #M
}
