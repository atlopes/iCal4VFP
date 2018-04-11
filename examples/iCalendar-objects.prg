* example: create an iCalendar object, element by element

* load the libraries
DO LOCFILE("icalloader.prg")

#INCLUDE "icalendar.h"

LOCAL Tst AS iCalendar
LOCAL TstEvent AS iCalCompVEVENT

#DEFINE ICSFILE	ADDBS(SYS(2023)) + "vfp-iCalendar.ics"

* create an iCalendar core object
m.Tst = CREATEOBJECT("iCalendar")
* add the required properties
m.Tst.AddICProperty("VERSION")
m.Tst.AddICProperty("PRODID")

* add a VEVENT component to the created iCalendar core object
m.TstEvent = m.Tst.AddICcomponent(CREATEOBJECT("iCalCompVEVENT"))

* add properties to the VEVENT component (values come from VFP expressions, where possible)
m.TstEvent.AddICproperty("DTSTAMP", {^1996-07-04 12:00:00}, ICAL_DATE_IS_UTC)
m.TstEvent.AddICproperty("UID", "uid1@example.com")
m.TstEvent.AddICproperty("ORGANIZER", "mailto:jsmith@example.com")
m.TstEvent.AddICproperty("DTSTART", {^1996-09-18 14:30:00}, ICAL_DATE_IS_UTC)
m.TstEvent.AddICproperty("DTEND", {^1996-09-20 22:00:00}, ICAL_DATE_IS_UTC)
m.TstEvent.AddICproperty("STATUS", "CONFIRMED")
m.TstEvent.AddICproperty("CATEGORIES", "CONFERENCE")
m.TstEvent.AddICproperty("SUMMARY", "Networld+Interop Conference")
m.TstEvent.AddICproperty("DESCRIPTION", "Networld+Interop Conference and Exhibit" + CHR(13) + CHR(10) + ;
																					"Atlanta World Congress Center" + CHR(13) + CHR(10) + ;
																					"Atlanta, Georgia")

* serialize in iCalendar format and see the result
IF STRTOFILE(m.Tst.Serialize(), ICSFILE, 0) > 0
	MODIFY FILE ICSFILE NOEDIT
ENDIF

* clean up
m.TstEvent = .NULL.
m.Tst = .NULL.
