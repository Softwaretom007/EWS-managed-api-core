using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.Data.Misc
{
    public class AdjustmentRule
    {
        public DateTime DateStart { get;  }
        public DateTime DateEnd { get;  }
        public TransitionTime DaylightTransitionStart { get;  }
        public TransitionTime DaylightTransitionEnd { get;  }
        public TimeSpan DaylightDelta { get; }

        public TimeZoneInfo.AdjustmentRule Origin { get; }

        public AdjustmentRule(TimeZoneInfo.AdjustmentRule rule)
        {
            Origin = rule;
            DateStart = rule.DateStart;
            DateEnd = rule.DateEnd;
            DaylightTransitionStart = new TransitionTime(rule.DaylightTransitionStart);
            DaylightTransitionEnd = new TransitionTime(rule.DaylightTransitionEnd);
            DaylightDelta = rule.DaylightDelta;
        }

        internal static AdjustmentRule CreateAdjustmentRule(DateTime date1, DateTime date2, TimeSpan timeSpan, TransitionTime transitionTime1, TransitionTime transitionTime2)
        {
            return new AdjustmentRule(TimeZoneInfo.AdjustmentRule.CreateAdjustmentRule(date1, date2, timeSpan, transitionTime1.Origin, transitionTime2.Origin));
        }
    }
}
