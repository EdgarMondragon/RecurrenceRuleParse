# RecurrenceRuleParse

Pass the recurrencerule string and the start date of the appointment to this method that will return the dates collection based on the recurrence rule

Rules based on Telerik Rule and FullCalendar Rule

example of rrule:
RRULE:FREQ=MONTHLY;COUNT=24;INTERVAL=1;BYSETPOS=1;BYDAY=MO,TU,WE,TH,FR,SA,SU\nEXDATE:20200401T000000Z,20200501T000000Z

RRULE:FREQ=HOURLY;UNTIL=20170430T000000Z;INTERVAL=1 EXDATE:20170413T170000Z,20170417T200000Z,20170425T000000Z

RRULE:FREQ=DAILY;COUNT=7;INTERVAL=1;BYDAY=MO,TU,WE,TH,FR,SA,SU\nEXDATE:20110128T152000Z

DTSTART:20200220T160000Z DTEND:20200220T163000Z RRULE:FREQ=WEEKLY;UNTIL=20200401T000000Z;INTERVAL=1;BYDAY=MO,TU,WE,TH,FR

Return recurrence date list

Implementation

List<DateTime> dates = new List<DateTime>();

dates = RecurrenceHelper.GetRecurrenceDateTimeCollection(RecurrenceRule, DateTime.Now).ToList();

Parse Hourly, Daily, Weekly and Month recurrece rule