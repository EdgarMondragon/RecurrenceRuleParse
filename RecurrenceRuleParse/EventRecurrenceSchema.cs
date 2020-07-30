using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RecurrenceRuleParse
{
	public class EventRecurrenceSchema
	{
		public EventRecurrenceSchema()
		{
			Hour = new HourSchema();
			Day = new DaySchema();
			Week = new WeekSchema();
			Month = new MonthlySchema();
			End = new EndSchema();

		}
		public HourSchema Hour { get; set; }

		public DaySchema Day { get; set; }

		public WeekSchema Week { get; set; }

		public MonthlySchema Month { get; set; }

		public EndSchema End { get; set; }

		public SchemaType SchemaType { get; set; }

		public string ExcludeDates { get; set; }
	}

	public class HourSchema
	{
		public int EveryHour { get; set; }
	}

	public class DaySchema
	{
		public int EveryDay { get; set; }

		public bool IsEveryWeekDay { get; set; }
	}

	public class WeekSchema
	{
		public int EveryWeek { get; set; }

		public List<int> DaysOfTheWeek { get; set; }

		public WeekSchema()
		{
			DaysOfTheWeek = new List<int>();
		}
	}

	public class MonthlySchema
	{
		public bool IsEveryMonth { get; set; }

		public int DayMonth { get; set; }

		public int EveryMonth { get; set; }

		public int MonthlyFrequency { get; set; }

		public int MonthlyPeriod { get; set; }

	}

	public class EndSchema
	{
		public int Ocurrences { get; set; }

		public bool IsEndDate { get; set; }

		public DateTime? EndDate { get; set; }
	}

	public enum SchemaType
	{
		Hour,
		Day,
		Week,
		Month
	}
}
