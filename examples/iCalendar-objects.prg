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
m.Tst.AddICProperty(CREATEOBJECT("iCalPropVERSION"))
m.Tst.AddICProperty(CREATEOBJECT("iCalPropPRODID"))

* add a VEVENT component to the created iCalendar core object
m.TstEvent = m.Tst.AddICcomponent(CREATEOBJECT("iCalCompVEVENT"))

* add properties to the VEVENT component (values come from VFP expressions, where possible)
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropDTSTAMP", {^1996-07-04 12:00:00}, ICAL_DATE_IS_UTC))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropUID", "uid1@example.com"))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropORGANIZER", "mailto:jsmith@example.com"))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropDTSTART", {^1996-09-18 14:30:00}, ICAL_DATE_IS_UTC))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropDTEND", {^1996-09-20 22:00:00}, ICAL_DATE_IS_UTC))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropSTATUS", "CONFIRMED"))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropCATEGORIES", "CONFERENCE"))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropSUMMARY", "Networld+Interop Conference"))
m.TstEvent.AddICproperty(CREATEOBJECT("iCalPropDESCRIPTION", "Networld+Interop Conference and Exhibit" + CHR(13) + CHR(10) + ;
																					"Atlanta World Congress Center" + CHR(13) + CHR(10) + ;
																					"Atlanta, Georgia"))


* serialize in iCalendar format and see the result
IF STRTOFILE(m.Tst.Serialize(), ICSFILE, 0) > 0
	MODIFY FILE ICSFILE NOEDIT
ENDIF

* clean up
m.TstEvent = .NULL.
m.Tst = .NULL.
