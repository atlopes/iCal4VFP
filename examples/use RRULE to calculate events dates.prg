* uses the Recurrence Rule processor to identify the previous and next dates of an event, as well as all of event dates in a year

DO LOCFILE("icalloader.prg")

SET DATE ANSI
SET CENTURY ON
SET HOURS TO 24

LOCAL ICSSource AS String
LOCAL ICS AS ICSProcessor
LOCAL iCal AS iCalendar
LOCAL iEvent AS iCompVEVENT
LOCAL Timezone AS iCalCompTIMEZONE
LOCAL Rule AS iCalPropRRULE

* example from CALConnect
* http://www.calconnect.org/tests/PayDay.ics

TEXT TO m.ICSSource NOSHOW
BEGIN:VCALENDAR
CALSCALE:GREGORIAN
PRODID:-//Cyrusoft International\, Inc.//Mulberry v4.0//EN
VERSION:2.0
X-WR-CALNAME:PayDay
BEGIN:VTIMEZONE
LAST-MODIFIED:20040110T032845Z
TZID:US/Eastern
BEGIN:DAYLIGHT
DTSTART:20000404T020000
RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=4
TZNAME:EDT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
END:DAYLIGHT
BEGIN:STANDARD
DTSTART:20001026T020000
RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10
TZNAME:EST
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
END:STANDARD
END:VTIMEZONE
BEGIN:VEVENT
DTSTAMP:20050211T173501Z
DTSTART;VALUE=DATE:20040227
RRULE:FREQ=MONTHLY;BYDAY=-1MO,-1TU,-1WE,-1TH,-1FR;BYSETPOS=-1
SUMMARY:PAY DAY
UID:DC3D0301C7790B38631F1FBB@ninevah.local
END:VEVENT
END:VCALENDAR
ENDTEXT

m.ICS = CREATEOBJECT("ICSProcessor")
m.iCal = m.ICS.Read(m.ICSSource)
m.iEvent = m.iCal.GetICComponent("VEVENT")
m.Rule = m.iEvent.GetICproperty("RRULE")
m.Timezone = m.iCal.GetICComponent("VTIMEZONE")

* pay day is set to the last weekday (Monday to Friday) of each month

MESSAGEBOX(TEXTMERGE("Previous <<m.iEvent.GetICPropertyValue('SUMMARY')>>: " + ;
							"<<m.Rule.CalculatePrevious(m.iEvent.GetICPropertyValue('DTSTART'), DATETIME(), m.Timezone, .NULL., .NULL.)>> " + ;
							"<<m.Rule.PreviousTzName>>."), ;
							64, "RRULE previous and next dates")

MESSAGEBOX(TEXTMERGE("Next <<m.iEvent.GetICPropertyValue('SUMMARY')>>: " + ;
							"<<m.Rule.CalculateNext(m.iEvent.GetICPropertyValue('DTSTART'), DATETIME(), m.Timezone, .NULL., .NULL.)>> " + ;
							"<<m.Rule.NextTzName>>."), ;
							64, "RRULE previous and next dates")

LOCAL PayDayCursor AS String
LOCAL PayDays AS String

m.PayDayCursor = m.Rule.CalculateAll(m.iEvent.GetICPropertyValue('DTSTART'), DTOT(GOMONTH(DATE(), 12)), m.Timezone, .NULL., .NULL.)

SELECT LocalTime FROM (m.PayDayCursor) WHERE YEAR(LocalTime) = YEAR(DATE()) INTO ARRAY PayDaysInTheYear

m.PayDays = TEXTMERGE("List of pay days for the year <<YEAR(DATE())>>:<<CHR(13)+CHR(10)>>")
SCAN FOR YEAR(LocalTime) = YEAR(DATE())
	m.PayDays = m.PayDays + TEXTMERGE("- <<CMONTH(LocalTime)>>: <<TTOD(LocalTime)>> <<TzName>><<CHR(13)+CHR(10)>>")
ENDSCAN

MESSAGEBOX(m.PayDays, 64, "RRULE all dates")

USE IN (m.PayDayCursor)
