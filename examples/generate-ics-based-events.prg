* example: generate a series of occurences from an event definition

* load the libraries
DO LOCFILE("icalloader.prg")

* prepare a cursor to hold the generated series
CREATE CURSOR CalEvents (eventName varchar(64), startTime datetime, endTime datetime, startTimeUTC datetime, endTimeUTC datetime)
INDEX ON startTime TAG start

LOCAL ICS AS String
LOCAL IC AS iCalendar
LOCAL ICProc AS ICSProcessor
LOCAL TZ AS iCalCompVTIMEZONE
LOCAL EV AS iCalCompVEVENT
LOCAL DTSTART AS iCalPropDTSTART
LOCAL DTEND AS iCalPropDTEND
LOCAL RRULE AS iCalPropRRULE
LOCAL EventIndex AS Integer
LOCAL SeriesCursor AS String

LOCAL Start AS Datetime, End AS Datetime
LOCAL StartUTC AS Datetime, EndUTC AS Datetime
LOCAL IsLocalTime AS Logical
LOCAL Duration AS Integer

#DEFINE ICSFILE	ADDBS(SYS(2023)) + "vfp-iCalendar.ics"

* set the calendar record, consisting (basically) of a time zone + events information

* in this example, in the first event a schedule is set for the working days of the month of October 2024 from 14:00 to 17:00, Brussels time
* an EXDATE excludes the schedule of the 11th, and a combination os EXDATE and RDATE sets the schedule on the 14th from 16:00 to 19:00

* a second set of events occurs on the working days of the same month but during the morning, every two days

* UTC times are also registered, and after the last Sunday of the month the UTC offset decreases to just one hour,
* as the CEST time zone enters standard time

* you may easily create other ICS records online at https://ical.marudot.com/

TEXT TO m.ICS NOSHOW
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//ical.marudot.com//iCal Event Maker
CALSCALE:GREGORIAN
BEGIN:VTIMEZONE
TZID:Europe/Brussels
LAST-MODIFIED:20240422T053450Z
TZURL:https://www.tzurl.org/zoneinfo-outlook/Europe/Brussels
X-LIC-LOCATION:Europe/Brussels
BEGIN:DAYLIGHT
TZNAME:CEST
TZOFFSETFROM:+0100
TZOFFSETTO:+0200
DTSTART:19700329T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU
END:DAYLIGHT
BEGIN:STANDARD
TZNAME:CET
TZOFFSETFROM:+0200
TZOFFSETTO:+0100
DTSTART:19701025T030000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU
END:STANDARD
END:VTIMEZONE
BEGIN:VEVENT
DTSTAMP:20241002T055225Z
UID:1727848262290-76110-1@ical.marudot.com
DTSTART;TZID=Europe/Brussels:20241001T140000
RRULE:FREQ=DAILY;BYDAY=MO,TU,WE,TH,FR;UNTIL=20241101T000000Z
DTEND;TZID=Europe/Brussels:20241001T170000
EXDATE:20241011140000
EXDATE:20241014140000
RDATE:20241014160000
SUMMARY:October's schedule (afternoon)
END:VEVENT
BEGIN:VEVENT
DTSTAMP:20241002T055225Z
UID:1727848262290-76110-2@ical.marudot.com
DTSTART;TZID=Europe/Brussels:20241001T100000
RRULE:FREQ=DAILY;BYDAY=MO,TU,WE,TH,FR;INTERVAL=2;UNTIL=20241101T000000Z
DTEND;TZID=Europe/Brussels:20241001T120000
SUMMARY:October's schedule (morning)
END:VEVENT
END:VCALENDAR
ENDTEXT

* instantiate the ICS processor
m.ICProc = CREATEOBJECT("ICSProcessor")

* and read an iCalendar from memory
m.IC = m.ICProc.Read(m.ICS)

* fetch the time zone (the example uses only one, but several could be referenced in the event iCalendar definition)
m.TZ = m.IC.GetICComponent("VTIMEZONE")

* go through all events in the iCalendar definition
FOR m.EventIndex = 1 TO m.IC.GetICComponentsCount("VEVENT")

	* get the event (as above, several events could be referenced in the same definition)
	m.EV = m.IC.GetICComponent("VEVENT", m.EventIndex)

	* we need a starting date for the first event in the series
	m.DTSTART = m.EV.GetICProperty("DTSTART")
	m.Start = m.DTSTART.GetValue()
	* get the UTC time for the starting date
	IF m.DTSTART.GetICParameterValue("TZID") == m.TZ.GetICPropertyValue("TZID")
		m.StartUTC = m.TZ.ToUTC(m.Start)
		m.IsLocalTime = .T.
	ELSE
		m.StartUTC = m.Start
		m.IsLocalTime = .F.
	ENDIF

	* we may have, or not, an ending date to establish a duration
	m.Duration = 0
	m.DTEND = m.EV.GetICProperty("DTEND")
	IF ! ISNULL(m.DTEND)
		m.End = m.DTEND.GetValue()
		* as for the starting date
		IF m.DTEND.GetICParameterValue("TZID") == m.TZ.GetICPropertyValue("TZID")
			m.EndUTC = m.TZ.ToUTC(m.End)
		ELSE
			m.EndUTC = m.End
		ENDIF
		m.Duration = MAX(m.EndUTC - m.StartUTC, 0)
	ELSE
		m.End = m.Start
		m.EndUTC = m.StartUTC
	ENDIF

	* we may have, or not, a recurrence rule to generate a series of events
	m.RRULE = m.EV.GetICProperty("RRULE")
	IF ! ISNULL(m.RRULE)

		* arrange collections of dates that should be excluded from or included in the generated series
		LOCAL IncludeDates AS Collection, ExcludeDates AS Collection
		LOCAL DatesCount AS Integer

		m.IncludeDates = CREATEOBJECT("Collection")
		m.ExcludeDates = CREATEOBJECT("Collection")

		FOR m.DatesCount = 1 TO m.EV.GetICPropertiesCount("RDATE")
			m.IncludeDates.Add(m.EV.GetICProperty("RDATE", m.DatesCount))
		ENDFOR
		FOR m.DatesCount = 1 TO m.EV.GetICPropertiesCount("EXDATE")
			m.ExcludeDates.Add(m.EV.GetICProperty("EXDATE", m.DatesCount))
		ENDFOR

		* calculate the series as UTC or local (system) time
		IF m.IsLocalTime
			m.SeriesCursor = m.RRULE.CalculateAll(m.Start, {^9999-12-31}, m.TZ, m.IncludeDates, m.ExcludeDates)
		ELSE
			m.SeriesCursor = m.RRULE.CalculateAll(m.StartUTC, .NULL., .NULL., m.IncludeDates, m.ExcludeDates)
		ENDIF

		* clear dangling object references in the dates collection
		m.IncludeDates.Remove(-1)
		m.ExcludeDates.Remove(-1)

		IF ! ISNULL(m.SeriesCursor)

			INSERT INTO CalEvents (;
					eventname, ;
					startTime, ;
					endTime, ;
					startTimeUTC, ;
					endTimeUTC) ;
				SELECT ;
						m.EV.GetICPropertyValue("SUMMARY"), ;
						localtime, ;
						IIF(m.IsLocalTime, m.TZ.ToLocalTime(m.TZ.ToUTC(localtime) + m.Duration), localtime + m.Duration), ;
						IIF(m.IsLocalTime, m.TZ.ToUTC(localtime), localtime), ;
						IIF(m.IsLocalTime, m.TZ.ToUTC(localtime) + m.Duration, localtime + m.Duration) ;
					FROM (m.SeriesCursor)

		ENDIF

	ELSE

		* no recurrence, ignore any RDATEs and assume just a one-time event
		INSERT INTO CalEvents (;
				eventname, ;
				startTime, ;
				endTime, ;
				startTimeUTC, ;
				endTimeUTC) ;
			VALUES (;
				m.EV.GetICPropertyValue("SUMMARY"), ;
				m.Start, ;
				m.End, ;
				m.StartUTC, ;
				m.EndUTC)

	ENDIF

	* discard the rules cursor
	USE IN SELECT(m.SeriesCursor)

ENDFOR

* view the series
SELECT CalEvents
GO TOP
BROWSE
