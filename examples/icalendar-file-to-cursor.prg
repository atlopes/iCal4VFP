* example: import an iCalendar from an ICS file and move a simplified version of its events into a cursor

* load the libraries
DO LOCFILE("icalloader.prg")

LOCAL IC AS iCalendar
LOCAL ICProc AS ICSProcessor

#DEFINE ICSFILE	ADDBS(SYS(2023)) + "vfp-iCalendar.ics"
#DEFINE ICSCURSOR	"tmpICS"

m.ICProc = CREATEOBJECT("ICSProcessor")

* read the ICS file from disk
m.IC = m.ICProc.ReadFile(GETFILE("ics"))

* is everything ok?
IF !ISNULL(m.IC)

	* just to control what we just imported...
	IF STRTOFILE(m.IC.Serialize(), ICSFILE, 0) > 0
		MODIFY FILE ICSFILE NOEDIT NOWAIT
		SET STEP ON
	ENDIF

	* send it to a cursor and display it
	m.ICProc.ICSToCursor(m.IC, "EVENTS", ICSCURSOR)
	BROWSE LAST

ENDIF

m.IC = .NULL.
m.ICProc = .NULL.

