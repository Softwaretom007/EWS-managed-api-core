using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.Data.Misc
{
    public struct TransitionTime
    {
        public int Day { get;  }
        public DayOfWeek DayOfWeek { get;  }
        public bool IsFixedDateRule { get;  }
        public int Month { get;  }
        public DateTime TimeOfDay { get;  }
        public int Week { get;  }

        public TimeZoneInfo.TransitionTime Origin { get; }

        public TransitionTime(TimeZoneInfo.TransitionTime time)
        {
            Origin = time;
            Day = time.Day;
            DayOfWeek = time.DayOfWeek;
            IsFixedDateRule = time.IsFixedDateRule;
            Month = time.Month;
            TimeOfDay = time.TimeOfDay;
            Week = time.Week;
        }

        internal static TransitionTime CreateFixedDateRule(DateTime dateTime, int month, int dayOrder)
        {
            return new TransitionTime(TimeZoneInfo.TransitionTime.CreateFixedDateRule(dateTime, month, dayOrder));
        }

        internal static TransitionTime CreateFloatingDateRule(DateTime dateTime, int month, int dayOrder, DayOfWeek dayOfWeek)
        {
            return new TransitionTime(TimeZoneInfo.TransitionTime.CreateFloatingDateRule(dateTime, month, dayOrder, dayOfWeek));
        }
    }
}
