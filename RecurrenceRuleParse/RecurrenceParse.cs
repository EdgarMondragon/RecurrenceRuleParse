using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace RecurrenceRuleParse
{
	public static class RecurrenceParse
	{
		#region Internal Fields

		internal static List<string> valueList = new List<string>();
		internal static string COUNT;
		internal static string RECCOUNT;
		internal static string HOURLY;
		internal static string DAILY;
		internal static string WEEKLY;
		internal static string MONTHLY;
		internal static string YEARLY;
		internal static string INTERVAL;
		internal static string INTERVALCOUNT;
		internal static string BYSETPOS;
		internal static string BYSETPOSCOUNT;
		internal static string BYDAY;
		internal static string BYDAYVALUE;
		internal static string BYMONTHDAY;
		internal static string BYMONTHDAYCOUNT;
		internal static string BYMONTH;
		internal static string BYMONTHCOUNT;
		internal static int BYDAYPOSITION;
		internal static int BYMONTHDAYPOSITION;
		internal static int WEEKLYBYDAYPOS;
		internal static string WEEKLYBYDAY;
		internal static string EXDATE;
		internal static List<DateTime> exDateList = new List<DateTime>();
		internal static Nullable<DateTime> UNTIL;
		#endregion

		#region Methods
		/// <summary>
		/// Pass the recurrencerule string and the start date of the appointment to this method that will return the dates collection based on the recurrence rule
		/// </summary>
		/// <param name="RRule"> string value</param>
		/// <param name="RecStartDate"> DateTime value</param>
		/// <returns></returns>
		public static IEnumerable<DateTime> GetRecurrenceDateTimeCollection(string RRule, DateTime RecStartDate)
		{
			RRule = RRule.Replace("\n", ";");
			var RecDateCollection = new List<DateTime>();
			DateTime startDate = RecStartDate;
			var ruleSeperator = new[] { '=', ';', ',' };
			var weeklySeperator = new[] { ';' };
			string[] ruleArray = RRule.Split(ruleSeperator);
			FindKeyIndex(ruleArray);
			string[] weeklyHourlyRule = RRule.Split(weeklySeperator);
			FindWeeklyRule(weeklyHourlyRule);
			FindExdateList(weeklyHourlyRule);
			if (ruleArray.Length != 0 && RRule != "")
			{
				DateTime addDate = startDate;
				int recCount;
				int.TryParse(RECCOUNT, out recCount);

				#region HOURLY

				if (HOURLY == "HOURLY")
				{
					int HyDayGap = ruleArray.Length == 4 ? 1 : int.Parse(INTERVALCOUNT);
					bool IsUntilDateReached = false;
					if (recCount == 0 && UNTIL == null)
					{
						recCount = 999;
					}
					if (UNTIL == null)
					{
						for (int i = 0; i < recCount; i++)
						{
							RecDateCollection.Add(addDate);
							addDate = addDate.AddHours(HyDayGap);
						}
					}
					else
					{
						while (!IsUntilDateReached)
						{
							RecDateCollection.Add(addDate);
							addDate = addDate.AddHours(HyDayGap);
							int statusValue = DateTime.Compare(addDate, Convert.ToDateTime(UNTIL));
							if (statusValue != -1)
							{
								IsUntilDateReached = true;
							}
						}
					}
				}
				#endregion
				#region DAILY
				else if (DAILY == "DAILY")
				{

					if ((ruleArray.Length > 4 && INTERVAL == "INTERVAL") || ruleArray.Length == 4)
					{
						int DyDayGap = ruleArray.Length == 4 ? 1 : int.Parse(INTERVALCOUNT);
						if (recCount == 0 && UNTIL == null)
						{
							recCount = 730;
						}
						if (recCount > 0)
						{
							for (int i = 0; i < recCount; i++)
							{
								RecDateCollection.Add(addDate);
								addDate = addDate.AddDays(DyDayGap);
							}
						}
						else if (UNTIL != null)
						{
							bool IsUntilDateReached = false;
							while (!IsUntilDateReached)
							{
								RecDateCollection.Add(addDate);
								addDate = addDate.AddDays(DyDayGap);
								int statusValue = DateTime.Compare(addDate, Convert.ToDateTime(UNTIL));
								if (statusValue != -1)
								{
									IsUntilDateReached = true;
								}
							}
						}
					}
					else if (ruleArray.Length > 4 && BYDAY == "BYDAY")
					{
						while (RecDateCollection.Count < recCount)
						{
							if (addDate.DayOfWeek != DayOfWeek.Sunday && addDate.DayOfWeek != DayOfWeek.Saturday)
							{

								RecDateCollection.Add(addDate);
							}
							addDate = addDate.AddDays(1);
						}
					}
				}
				#endregion

				#region WEEKLY
				else if (WEEKLY == "WEEKLY")
				{
					int WyWeekGap = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? int.Parse(INTERVALCOUNT) : 1;
					bool isweeklyselected = weeklyHourlyRule[WEEKLYBYDAYPOS].Length > 6;
					if (recCount == 0 && UNTIL == null)
					{
						recCount = 1;
					}
					if (recCount > 0)
					{
						while (RecDateCollection.Count < recCount && isweeklyselected)
						{
							GetWeeklyDateCollection(addDate, weeklyHourlyRule, RecDateCollection);
							addDate = addDate.DayOfWeek == DayOfWeek.Saturday ? addDate.AddDays(((WyWeekGap - 1) * 7) + 1) : addDate.AddDays(1);
						}
					}
					else if (UNTIL != null)
					{
						bool IsUntilDateReached = false;
						while (!IsUntilDateReached && isweeklyselected)
						{
							GetWeeklyDateCollection(addDate, weeklyHourlyRule, RecDateCollection);
							addDate = addDate.DayOfWeek == DayOfWeek.Saturday ? addDate.AddDays(((WyWeekGap - 1) * 7) + 1) : addDate.AddDays(1);
							int statusValue = DateTime.Compare(addDate, Convert.ToDateTime(UNTIL));
							if (statusValue != -1)
							{
								IsUntilDateReached = true;
							}
						}
					}
				}
				#endregion

				#region MONTHLY
				else if (MONTHLY == "MONTHLY")
				{
					int MyMonthGap = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? int.Parse(INTERVALCOUNT) : 1;
					if (BYMONTHDAY == "BYMONTHDAY")
					{
						int monthDate = int.Parse(BYMONTHDAYCOUNT);
						if (monthDate <= 31)
						{
							int currDate = int.Parse(startDate.Day.ToString());
							var temp = new DateTime(addDate.Year, addDate.Month, monthDate, addDate.Hour, addDate.Minute, addDate.Second);
							addDate = monthDate < currDate ? temp.AddMonths(1) : temp;
							if (recCount == 0 && UNTIL == null)
							{
								recCount = 24;
							}
							if (recCount > 0)
							{
								for (int i = 0; i < recCount; i++)
								{
									addDate = GetByMonthDayDateCollection(addDate, RecDateCollection, monthDate, MyMonthGap);
								}
							}
							else if (UNTIL != null)
							{
								bool IsUntilDateReached = false;
								while (!IsUntilDateReached)
								{
									addDate = GetByMonthDayDateCollection(addDate, RecDateCollection, monthDate, MyMonthGap);
									int statusValue = DateTime.Compare(addDate, Convert.ToDateTime(UNTIL));
									if (statusValue != -1)
									{
										IsUntilDateReached = true;
									}
								}
							}
						}
						else
						{
							if (recCount == 0 && UNTIL == null)
							{
								recCount = 24;
							}
							if (recCount > 0)
							{
								for (int i = 0; i < recCount; i++)
								{
									if (addDate.Day == startDate.Day)
									{
										RecDateCollection.Add(addDate);
									}
									else
									{
										i = i - 1;
									}
									addDate = addDate.AddMonths(MyMonthGap);
									addDate = new DateTime(addDate.Year, addDate.Month, DateTime.DaysInMonth(addDate.Year, addDate.Month), addDate.Hour, addDate.Minute, addDate.Second);
								}
							}
							else if (UNTIL != null)
							{
								bool IsUntilDateReached = false;
								while (!IsUntilDateReached)
								{
									if (addDate.Day == startDate.Day)
									{
										RecDateCollection.Add(addDate);
									}
									addDate = addDate.AddMonths(MyMonthGap);
									addDate = new DateTime(addDate.Year, addDate.Month, DateTime.DaysInMonth(addDate.Year, addDate.Month), addDate.Hour, addDate.Minute, addDate.Second);
									int statusValue = DateTime.Compare(addDate, Convert.ToDateTime(UNTIL));
									if (statusValue != -1)
									{
										IsUntilDateReached = true;
									}
								}
							}
						}
					}
					else if (BYDAY == "BYDAY")
					{
						if (recCount == 0 && UNTIL == null)
						{
							recCount = 24;
						}
						if (recCount > 0)
						{
							while (RecDateCollection.Count < recCount)
							{
								var weekCount = MondaysInMonth(addDate);
								var monthStart = new DateTime(addDate.Year, addDate.Month, 1, addDate.Hour, addDate.Minute, addDate.Second);
								DateTime weekStartDate = monthStart.AddDays(-(int)(monthStart.DayOfWeek));
								var monthStartWeekday = (int)(monthStart.DayOfWeek);
								int nthweekDay = GetWeekDay(BYDAYVALUE) - 1;
								int nthWeek;
								int bySetPos = 0;
								int setPosCount;
								int.TryParse(BYSETPOSCOUNT, out setPosCount);
								if (monthStartWeekday <= nthweekDay)
								{
									if (setPosCount < 1)
									{
										bySetPos = weekCount + setPosCount;
									}
									else
									{
										bySetPos = setPosCount;
									}
									if (setPosCount < 0)
									{
										nthWeek = bySetPos;
									}
									else
									{
										nthWeek = bySetPos - 1;
									}
								}
								else
								{
									if (setPosCount < 0)
									{
										bySetPos = weekCount + setPosCount;
									}
									else
									{
										bySetPos = setPosCount;
									}
									nthWeek = bySetPos;
								}
								addDate = weekStartDate.AddDays((nthWeek) * 7);
								addDate = addDate.AddDays(nthweekDay);
								if (addDate.CompareTo(startDate.Date) < 0)
								{
									addDate = addDate.AddMonths(1);
									continue;
								}
								if (weekCount == 6 && addDate.Day == 23)
								{
									int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
									bool flag = true;
									if (addDate.Month == 2)
									{
										flag = false;
									}
									if (flag)
									{
										addDate = addDate.AddDays(7);
										RecDateCollection.Add(addDate);
									}
									addDate = addDate.AddMonths(MyMonthGap);
								}
								else if (weekCount == 6 && addDate.Day == 24)
								{
									int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
									bool flag = true;
									if (addDate.AddDays(7).Day != days)
									{
										flag = false;
									}
									if (flag)
									{
										addDate = addDate.AddDays(7);
										RecDateCollection.Add(addDate);
									}
									addDate = addDate.AddMonths(MyMonthGap);
								}
								else if (!(addDate.Day <= 23 && int.Parse(BYSETPOSCOUNT) == -1))
								{
									RecDateCollection.Add(addDate);
									addDate = addDate.AddMonths(MyMonthGap);
								}
							}
						}
						else if (UNTIL != null)
						{
							bool IsUntilDateReached = false;
							while (!IsUntilDateReached)
							{
								var weekCount = MondaysInMonth(addDate);
								var monthStart = new DateTime(addDate.Year, addDate.Month, 1, addDate.Hour, addDate.Minute, addDate.Second);
								DateTime weekStartDate = monthStart.AddDays(-(int)(monthStart.DayOfWeek));
								var monthStartWeekday = (int)(monthStart.DayOfWeek);
								int nthweekDay = GetWeekDay(BYDAYVALUE) - 1;
								int nthWeek;
								int bySetPos = 0;
								int setPosCount;
								int.TryParse(BYSETPOSCOUNT, out setPosCount);
								if (monthStartWeekday <= nthweekDay)
								{
									if (setPosCount < 1)
									{
										bySetPos = weekCount + setPosCount;
									}
									else
									{
										bySetPos = setPosCount;
									}
									if (setPosCount < 0)
									{
										nthWeek = bySetPos;
									}
									else
									{
										nthWeek = bySetPos - 1;
									}
								}
								else
								{
									if (setPosCount < 0)
									{
										bySetPos = weekCount + setPosCount;
									}
									else
									{
										bySetPos = setPosCount;
									}
									nthWeek = bySetPos;
								}
								addDate = weekStartDate.AddDays((nthWeek) * 7);
								addDate = addDate.AddDays(nthweekDay);
								if (addDate.CompareTo(startDate) < 0)
								{
									addDate = addDate.AddMonths(1);
									continue;
								}
								if (weekCount == 6 && addDate.Day == 23)
								{
									int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
									bool flag = true;
									if (addDate.Month == 2)
									{
										flag = false;
									}
									if (flag)
									{
										addDate = addDate.AddDays(7);
										RecDateCollection.Add(addDate);
									}
									addDate = addDate.AddMonths(MyMonthGap);
								}
								else if (weekCount == 6 && addDate.Day == 24)
								{
									int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
									bool flag = true;
									if (addDate.AddDays(7).Day != days)
									{
										flag = false;
									}
									if (flag)
									{
										addDate = addDate.AddDays(7);
										RecDateCollection.Add(addDate);
									}
									addDate = addDate.AddMonths(MyMonthGap);
								}
								else if (!(addDate.Day <= 23 && int.Parse(BYSETPOSCOUNT) == -1))
								{
									RecDateCollection.Add(addDate);
									addDate = addDate.AddMonths(MyMonthGap);
								}
								int statusValue = DateTime.Compare(addDate, Convert.ToDateTime(UNTIL));
								if (statusValue != -1)
								{
									IsUntilDateReached = true;
								}
							}
						}
					}
				}
				#endregion
			}
			return RecDateCollection.Except(exDateList).ToList();
		}

		public static EventRecurrenceSchema GetRecurrenceSchema(string RRule)
		{
			EventRecurrenceSchema eventRecurrenceSchema = new EventRecurrenceSchema();
			RRule = RRule.Replace("\n", ";");
			var ruleSeperator = new[] { '=', ';', ',' };
			var weeklySeperator = new[] { ';' };
			string[] ruleArray = RRule.Split(ruleSeperator);
			FindKeyIndex(ruleArray);
			string[] weeklyHourlyRule = RRule.Split(weeklySeperator);
			FindWeeklyRule(weeklyHourlyRule);
			FindExdateList(weeklyHourlyRule);
			eventRecurrenceSchema.ExcludeDates = string.Join(",", exDateList);
			if (ruleArray.Length != 0 && RRule != "")
			{
				int recCount = 0;
				int.TryParse(RECCOUNT, out recCount);
				eventRecurrenceSchema.End.Ocurrences = recCount;
				eventRecurrenceSchema.End.EndDate = UNTIL;
				#region HOURLY

				if (HOURLY == "HOURLY")
				{
					int HyDayGap = ruleArray.Length == 4 ? 1 : int.Parse(INTERVALCOUNT);
					eventRecurrenceSchema.SchemaType = SchemaType.Hour;
					eventRecurrenceSchema.Hour.EveryHour = HyDayGap;
				}
				#endregion
				#region DAILY
				else if (DAILY == "DAILY")
				{
					eventRecurrenceSchema.SchemaType = SchemaType.Day;
					if ((ruleArray.Length > 4 && INTERVAL == "INTERVAL") || ruleArray.Length == 4)
					{
						int DyDayGap = ruleArray.Length == 4 ? 1 : int.Parse(INTERVALCOUNT);
						eventRecurrenceSchema.Day.EveryDay = DyDayGap;
						eventRecurrenceSchema.Day.IsEveryWeekDay = false;
					}
					if (ruleArray.Length > 4 && BYDAY == "BYDAY")
					{
						eventRecurrenceSchema.Day.IsEveryWeekDay = true;
					}
				}
				#endregion

				#region WEEKLY
				else if (WEEKLY == "WEEKLY")
				{
					eventRecurrenceSchema.SchemaType = SchemaType.Week;
					int WyWeekGap = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? int.Parse(INTERVALCOUNT) : 1;
					bool isweeklyselected = weeklyHourlyRule[WEEKLYBYDAYPOS].Length > 6;
					eventRecurrenceSchema.Week.EveryWeek = WyWeekGap;
					if (isweeklyselected)
					{
						if (weeklyHourlyRule[WEEKLYBYDAYPOS].Contains("SU"))
						{
							eventRecurrenceSchema.Week.DaysOfTheWeek.Add((int)DayOfWeek.Sunday);
						}
						if (weeklyHourlyRule[WEEKLYBYDAYPOS].Contains("MO"))
						{
							eventRecurrenceSchema.Week.DaysOfTheWeek.Add((int)DayOfWeek.Monday);
						}
						if (weeklyHourlyRule[WEEKLYBYDAYPOS].Contains("TU"))
						{
							eventRecurrenceSchema.Week.DaysOfTheWeek.Add((int)DayOfWeek.Tuesday);
						}
						if (weeklyHourlyRule[WEEKLYBYDAYPOS].Contains("WE"))
						{
							eventRecurrenceSchema.Week.DaysOfTheWeek.Add((int)DayOfWeek.Wednesday);
						}
						if (weeklyHourlyRule[WEEKLYBYDAYPOS].Contains("TH"))
						{
							eventRecurrenceSchema.Week.DaysOfTheWeek.Add((int)DayOfWeek.Thursday);
						}
						if (weeklyHourlyRule[WEEKLYBYDAYPOS].Contains("FR"))
						{
							eventRecurrenceSchema.Week.DaysOfTheWeek.Add((int)DayOfWeek.Friday);
						}
						if (weeklyHourlyRule[WEEKLYBYDAYPOS].Contains("SA"))
						{
							eventRecurrenceSchema.Week.DaysOfTheWeek.Add((int)DayOfWeek.Saturday);
						}
					}
				}
				#endregion

				#region MONTHLY
				else if (MONTHLY == "MONTHLY")
				{
					eventRecurrenceSchema.SchemaType = SchemaType.Month;
					int MyMonthGap = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? int.Parse(INTERVALCOUNT) : 1;

					if (BYMONTHDAY == "BYMONTHDAY")
					{
						eventRecurrenceSchema.Month.DayMonth = int.Parse(BYMONTHDAYCOUNT);
						eventRecurrenceSchema.Month.EveryMonth = MyMonthGap;
						eventRecurrenceSchema.Month.IsEveryMonth = true;
					}
					else if (BYDAY == "BYDAY")
					{
						eventRecurrenceSchema.Month.IsEveryMonth = false;
						eventRecurrenceSchema.Month.EveryMonth = MyMonthGap;
						eventRecurrenceSchema.Month.MonthlyFrequency = int.Parse(BYSETPOSCOUNT) - 1;

						var byDay = WEEKLYBYDAY.Split('=');
						for (int i = 0; i < byDay.Length; i++)
						{
							if (byDay[i] == "BYDAY")
							{
								var days = byDay[i + 1];
								if (days == "MO,TU,WE,TH,FR,SA,SU")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 0;
								}
								if (days == "MO,TU,WE,TH,FR")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 1;
								}
								if (days == "SA,SU")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 2;
								}
								if (days == "SU")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 3;
								}
								if (days == "MO")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 4;
								}
								if (days == "TU")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 5;
								}
								if (days == "WE")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 6;
								}
								if (days == "TH")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 7;
								}
								if (days == "FR")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 8;
								}
								if (days == "SA")
								{
									eventRecurrenceSchema.Month.MonthlyPeriod = 9;
								}
							}
						}
					}
				}
				#endregion
			}
			return eventRecurrenceSchema;
		}

		private static void FindExdateList(string[] weeklyRule)
		{
			for (int i = 0; i < weeklyRule.Length; i++)
			{
				if (weeklyRule[i].Contains("EXDATE"))
				{
					EXDATE = weeklyRule[i];
					var _rule = weeklyRule[i].Substring(weeklyRule[i].IndexOf("EXDATE")).Split(new[] { '=', ':' });
					if (_rule[0] == "EXDATE")
					{
						var exDates = _rule[1].Split(',');
						for (var j = 0; j < exDates.Length; j++)
						{
							StringBuilder sb = new StringBuilder(exDates[j].Trim());
							sb.Insert(4, '-'); sb.Insert(7, '-'); sb.Insert(13, ':'); sb.Insert(16, ':');
							DateTimeOffset value = DateTimeOffset.ParseExact(sb.ToString(), "yyyy-MM-dd'T'HH:mm:ss'Z'",
																   CultureInfo.InvariantCulture);
							exDateList.Add(value.DateTime);
						}
					}
					break;
				}
			}
		}

		private static int GetWeekDay(string weekDay)
		{
			switch (weekDay)
			{
				case "SU":
					{
						return 1;
					}
				case "MO":
					{
						return 2;
					}
				case "TU":
					{
						return 3;
					}
				case "WE":
					{
						return 4;
					}
				case "TH":
					{
						return 5;
					}
				case "FR":
					{
						return 6;
					}
				case "SA":
					{
						return 7;
					}
			}
			return 8;
		}

		private static void FindWeeklyRule(string[] weeklyRule)
		{
			for (int i = 0; i < weeklyRule.Length; i++)
			{
				if (weeklyRule[i].Contains("BYDAY"))
				{
					WEEKLYBYDAY = weeklyRule[i];
					WEEKLYBYDAYPOS = i;
					break;
				}
			}
		}

		private static void FindKeyIndex(string[] ruleArray)
		{
			RECCOUNT = "";
			DAILY = "";
			WEEKLY = "";
			MONTHLY = "";
			YEARLY = "";
			BYSETPOS = "";
			BYSETPOSCOUNT = "";
			INTERVAL = "";
			INTERVALCOUNT = "";
			COUNT = "";
			BYDAY = "";
			BYDAYVALUE = "";
			BYMONTHDAY = "";
			BYMONTHDAYCOUNT = "";
			BYMONTH = "";
			BYMONTHCOUNT = "";
			WEEKLYBYDAY = "";
			UNTIL = null;

			for (int i = 0; i < ruleArray.Length; i++)
			{
				if (ruleArray[i].Contains("COUNT"))
				{
					COUNT = ruleArray[i];
					RECCOUNT = ruleArray[i + 1];

				}

				if (ruleArray[i].Contains("UNTIL"))
				{
					StringBuilder sb = new StringBuilder(ruleArray[i + 1]);
					sb.Insert(4, '-'); sb.Insert(7, '-'); sb.Insert(13, ':'); sb.Insert(16, ':');
					DateTimeOffset value = DateTimeOffset.ParseExact(sb.ToString(), "yyyy-MM-dd'T'HH:mm:ss'Z'",
														   CultureInfo.InvariantCulture);
					UNTIL = value.DateTime.Date;
				}

				if (ruleArray[i].Contains("HOURLY"))
				{
					HOURLY = ruleArray[i].Trim();
				}

				if (ruleArray[i].Contains("DAILY"))
				{
					DAILY = ruleArray[i].Trim();
				}

				if (ruleArray[i].Contains("WEEKLY"))
				{
					WEEKLY = ruleArray[i].Trim();
				}
				if (ruleArray[i].Contains("INTERVAL"))
				{
					INTERVAL = ruleArray[i].Trim();
					INTERVALCOUNT = ruleArray[i + 1];
				}
				if (ruleArray[i].Contains("MONTHLY"))
				{
					MONTHLY = ruleArray[i].Trim();
				}
				if (ruleArray[i].Contains("YEARLY"))
				{
					YEARLY = ruleArray[i].Trim();
				}
				if (ruleArray[i].Contains("BYSETPOS"))
				{
					BYSETPOS = ruleArray[i].Trim();
					BYSETPOSCOUNT = ruleArray[i + 1];
				}
				if (ruleArray[i].Contains("BYDAY"))
				{
					BYDAYPOSITION = i;
					BYDAY = ruleArray[i].Trim();
					BYDAYVALUE = ruleArray[i + 1].Trim();
				}
				if (ruleArray[i].Contains("BYMONTHDAY"))
				{
					BYMONTHDAYPOSITION = i;
					BYMONTHDAY = ruleArray[i].Trim();
					BYMONTHDAYCOUNT = ruleArray[i + 1].Trim();
				}
				if (ruleArray[i].Contains("BYMONTH"))
				{
					BYMONTH = ruleArray[i].Trim();
					BYMONTHCOUNT = ruleArray[i + 1].Trim();
				}
			}
		}

		private static void GetWeeklyDateCollection(DateTime addDate, string[] weeklyRule, List<DateTime> RecDateCollection)
		{
			switch (addDate.DayOfWeek)
			{
				case DayOfWeek.Sunday:
					{
						if (weeklyRule[WEEKLYBYDAYPOS].Contains("SU"))
						{
							RecDateCollection.Add(addDate);
						}
						break;
					}
				case DayOfWeek.Monday:
					{
						if (weeklyRule[WEEKLYBYDAYPOS].Contains("MO"))
						{
							RecDateCollection.Add(addDate);
						}
						break;
					}
				case DayOfWeek.Tuesday:
					{
						if (weeklyRule[WEEKLYBYDAYPOS].Contains("TU"))
						{
							RecDateCollection.Add(addDate);
						}
						break;
					}
				case DayOfWeek.Wednesday:
					{
						if (weeklyRule[WEEKLYBYDAYPOS].Contains("WE"))
						{
							RecDateCollection.Add(addDate);
						}
						break;
					}
				case DayOfWeek.Thursday:
					{
						if (weeklyRule[WEEKLYBYDAYPOS].Contains("TH"))
						{
							RecDateCollection.Add(addDate);
						}
						break;
					}
				case DayOfWeek.Friday:
					{
						if (weeklyRule[WEEKLYBYDAYPOS].Contains("FR"))
						{
							RecDateCollection.Add(addDate);
						}
						break;
					}
				case DayOfWeek.Saturday:
					{
						if (weeklyRule[WEEKLYBYDAYPOS].Contains("SA"))
						{
							RecDateCollection.Add(addDate);
						}
						break;
					}
			}
		}

		private static DateTime GetByMonthDayDateCollection(DateTime addDate, List<DateTime> RecDateCollection, int monthDate, int MyMonthGap)
		{
			if (addDate.Month == 2 && monthDate > 28)
			{
				addDate = new DateTime(addDate.Year, addDate.Month, DateTime.DaysInMonth(addDate.Year, 2));
				addDate = addDate.AddMonths(MyMonthGap);
				addDate = new DateTime(addDate.Year, addDate.Month, monthDate, addDate.Hour, addDate.Minute, addDate.Second);
			}
			else
			{
				RecDateCollection.Add(addDate);
				addDate = addDate.AddMonths(MyMonthGap);
			}
			return addDate;
		}

		private static int MondaysInMonth(DateTime thisMonth)
		{
			DateTime today = thisMonth;
			//extract the month
			int daysInMonth = DateTime.DaysInMonth(today.Year, today.Month);
			DateTime firstOfMonth = new DateTime(today.Year, today.Month, 1, today.Hour, today.Minute, today.Second);
			//days of week starts by default as Sunday = 0
			int firstDayOfMonth = (int)firstOfMonth.DayOfWeek;
			int weeksInMonth = (int)Math.Ceiling((firstDayOfMonth + daysInMonth) / 7.0);
			return weeksInMonth;
		}
		#endregion
	}
}
