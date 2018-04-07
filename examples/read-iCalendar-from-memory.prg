* example: import an iCalendar from a memory string

* load the libraries
DO LOCFILE("icalloader.prg")

LOCAL ICS AS String
LOCAL IC AS iCalendar
LOCAL ICProc AS ICSProcessor

#DEFINE ICSFILE	ADDBS(SYS(2023)) + "vfp-iCalendar.ics"

* set the contents of a variable
TEXT TO m.ICS NOSHOW
BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//ABC Corporation//NONSGML My Product//EN
BEGIN:VTODO
DTSTAMP:19980130T134500Z
SEQUENCE:2
UID:uid4@example.com
ORGANIZER:mailto:unclesam@example.com
ATTENDEE;PARTSTAT=ACCEPTED:mailto:jqpublic@example.com
DUE:19980415T000000
STATUS:NEEDS-ACTION
SUMMARY:Submit Income Taxes
BEGIN:VALARM
ACTION:AUDIO
TRIGGER:19980403T120000Z
ATTACH;FMTTYPE=audio/basic:http://example.com/pub/audio-
 files/ssbanner.aud
REPEAT:4
DURATION:PT1H
END:VALARM
END:VTODO
END:VCALENDAR
ENDTEXT

* instantiate the ICS processor
m.ICProc = CREATEOBJECT("ICSProcessor")

* and read an iCalendar from memory
m.IC = m.ICProc.Read(m.ICS)

* serialize in iCalendar format and see the result
IF STRTOFILE(m.IC.Serialize(), ICSFILE, 0) > 0
	MODIFY FILE ICSFILE NOEDIT
ENDIF

m.IC = .NULL.
m.ICProc = .NULL.
