using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.Data.Misc
{
    public static class TimeZoneExtensions
    {
        public static TimeZoneInfo CreateCustomTimeZone(string id, TimeSpan baseOffsetToUtc, string name, string standardDisplayName)
        {
            return TimeZoneInfo.CreateCustomTimeZone(id, baseOffsetToUtc, name, standardDisplayName);
        }

        public static TimeZoneInfo CreateCustomTimeZone(string id, TimeSpan baseOffsetToUtc, string name, string standardDisplayName, 
            string daylightDisplayName, AdjustmentRule[] adjustmentRule)
        {
            return TimeZoneInfo.CreateCustomTimeZone(id, baseOffsetToUtc, name, standardDisplayName,
                daylightDisplayName, adjustmentRule.Select(y => y.Origin).ToArray());
        }

        public static AdjustmentRule[] GetAdjustmentRulesEx(this TimeZoneInfo tz)
        {
            return tz.GetAdjustmentRules().Select( y => new AdjustmentRule(y)).ToArray();
        }
    }
}
