* tests the Recurrence Rule processor by running all RFC 5545 RRULE examples

DO LOCFILE("icalloader.prg")

* test all 42 examples, or just the ones that are marked
#DEFINE TESTALL		.T.
* set verbosity (in each example, display the expected vs. calculated dates or report just the final summary)
#DEFINE VERBOSE		.T.
* set fail verbosity (when fail, browse the verification cursor)
#DEFINE VERBOSEFAIL	.T.

SET DATE ANSI
SET CENTURY ON
SET HOURS TO 24

LOCAL StartDate AS Datetime
LOCAL LoopIndex AS Integer
LOCAL LoopDate AS Datetime

LOCAL Description AS String
LOCAL NumberOfTests AS Integer
LOCAL NumberOfFails AS Integer

LOCAL ICSSource AS String
LOCAL ICS AS ICSProcessor
LOCAL iCal AS iCalendar
LOCAL Timezone AS iCalCompTIMEZONE

TEXT TO m.ICSSource NOSHOW
BEGIN:VCALENDAR
VERSION:2.0
PRODID://RFC5545 example
BEGIN:VTIMEZONE
TZID:America/New_York
LAST-MODIFIED:20050809T050000Z
BEGIN:DAYLIGHT
DTSTART:19670430T020000
RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=-1SU;UNTIL=19730429T070000Z
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
END:DAYLIGHT
BEGIN:STANDARD
DTSTART:19671029T020000
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU;UNTIL=20061029T060000Z
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:19740106T020000
RDATE:19750223T020000
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
END:DAYLIGHT
BEGIN:DAYLIGHT
DTSTART:19760425T020000
RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=-1SU;UNTIL=19860427T070000Z
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
END:DAYLIGHT
BEGIN:DAYLIGHT
DTSTART:19870405T020000
RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU;UNTIL=20060402T070000Z
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
END:DAYLIGHT
BEGIN:DAYLIGHT
DTSTART:20070311T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
END:DAYLIGHT
BEGIN:STANDARD
DTSTART:20071104T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
END:STANDARD
END:VTIMEZONE
END:VCALENDAR
ENDTEXT

m.ICS = CREATEOBJECT("ICSProcessor")
m.iCal = m.ICS.Read(m.ICSSource)
m.Timezone = m.iCal.GetTimezone("America/New_York")

m.NumberOfTests = 0
m.NumberOfFails = 0

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
Daily for 10 occurrences:

DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=DAILY;COUNT=10

       ==> (1997 9:00 AM EDT) September 2-11
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
FOR m.LoopIndex = 1 TO 10
	INSERT INTO Verification (Expected) VALUES (m.StartDate + (m.LoopIndex - 1) * 86400)
ENDFOR

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=DAILY;COUNT=10", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1
#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
Daily until December 24, 1997:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=DAILY;UNTIL=19971224T000000Z

       ==> (1997 9:00 AM EDT) September 2-30;October 1-25
           (1997 9:00 AM EST) October 26-31;November 1-30;December 1-23
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
FOR m.LoopIndex = 1 TO 113
	INSERT INTO Verification (Expected) VALUES (m.StartDate + (m.LoopIndex - 1) * 86400)
ENDFOR

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=DAILY;UNTIL=19971224T000000Z", m.StartDate, 113)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
Every other day - forever:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=DAILY;INTERVAL=2

       ==> (1997 9:00 AM EDT) September 2,4,6,8...24,26,28,30;
                              October 2,4,6...20,22,24
           (1997 9:00 AM EST) October 26,28,30;
                              November 1,3,5,7...25,27,29;
                              December 1,3,...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
m.LoopDate = m.StartDate
DO WHILE m.LoopDate <= {^1997-12-03 09:00:00}
	INSERT INTO Verification (Expected) VALUES (m.LoopDate)
	m.LoopDate = m.LoopDate + 86400 * 2
ENDDO

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=DAILY;INTERVAL=2", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every 10 days, 5 occurrences:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=DAILY;INTERVAL=10;COUNT=5

       ==> (1997 9:00 AM EDT) September 2,12,22;
                              October 2,12
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
FOR m.LoopIndex = 1 TO 5
	INSERT INTO Verification (Expected) VALUES (m.StartDate + ((m.LoopIndex - 1) * 10) * 86400)
ENDFOR

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=DAILY;INTERVAL=10;COUNT=5", m.StartDate, 5)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every day in January, for 3 years:

       DTSTART;TZID=America/New_York:19980101T090000

       RRULE:FREQ=YEARLY;UNTIL=20000131T140000Z;
        BYMONTH=1;BYDAY=SU,MO,TU,WE,TH,FR,SA
       or
       RRULE:FREQ=DAILY;UNTIL=20000131T140000Z;BYMONTH=1

       ==> (1998 9:00 AM EST)January 1-31
           (1999 9:00 AM EST)January 1-31
           (2000 9:00 AM EST)January 1-31
ENDTEXT

ResetVerification()

m.StartDate = {^1998-01-01 09:00:00}
m.LoopDate = m.StartDate
DO WHILE YEAR(m.LoopDate) <= 2000
	FOR m.LoopIndex = 1 TO 31
		INSERT INTO Verification (Expected) VALUES (m.LoopDate + (m.LoopIndex - 1) * 86400)
	ENDFOR
	m.LoopDate = DATETIME(YEAR(m.LoopDate) + 1, MONTH(m.LoopDate), DAY(m.LoopDate), HOUR(m.LoopDate), MINUTE(m.LoopDate), SEC(m.LoopDate))
ENDDO

IF !VerifyCalculation(m.Timezone, m.Description + CHR(13) + CHR(10) + "(option 1)", "RRULE:FREQ=YEARLY;UNTIL=20000131T140000Z;BYMONTH=1;BYDAY=SU,MO,TU,WE,TH,FR,SA", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

ClearVerification()

IF !VerifyCalculation(m.Timezone, m.Description + CHR(13) + CHR(10) +  "(option 2)", "RRULE:FREQ=DAILY;UNTIL=20000131T140000Z;BYMONTH=1", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
 Weekly for 10 occurrences:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=WEEKLY;COUNT=10

       ==> (1997 9:00 AM EDT) September 2,9,16,23,30;October 7,14,21
           (1997 9:00 AM EST) October 28;November 4
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
m.LoopDate = m.StartDate
FOR m.LoopIndex = 1 TO 10
	INSERT INTO Verification (Expected) VALUES (m.LoopDate + ((m.LoopIndex - 1) * 7) * 86400)
ENDFOR

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=WEEKLY;COUNT=10", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
Weekly until December 24, 1997:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=WEEKLY;UNTIL=19971224T000000Z

       ==> (1997 9:00 AM EDT) September 2,9,16,23,30;
                              October 7,14,21
           (1997 9:00 AM EST) October 28;
                              November 4,11,18,25;
                              December 2,9,16,23
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
m.LoopDate = m.StartDate
DO WHILE m.LoopDate <= {^1997-12-24 00:00:00}
	INSERT INTO Verification (Expected) VALUES (m.LoopDate)
	m.LoopDate = m.LoopDate + 7 * 86400
ENDDO

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=WEEKLY;UNTIL=19971224T000000Z", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every other week - forever:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=WEEKLY;INTERVAL=2;WKST=SU

       ==> (1997 9:00 AM EDT) September 2,16,30;
                              October 14
           (1997 9:00 AM EST) October 28;
                              November 11,25;
                              December 9,23
           (1998 9:00 AM EST) January 6,20;
                              February 3, 17
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
m.LoopDate = m.StartDate
DO WHILE m.LoopDate <= {^1998-02-18 00:00:00}
	INSERT INTO Verification (Expected) VALUES (m.LoopDate)
	m.LoopDate = m.LoopDate + 7 * 86400 * 2
ENDDO

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=WEEKLY;INTERVAL=2;WKST=SU", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Weekly on Tuesday and Thursday for five weeks:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=WEEKLY;UNTIL=19971007T000000Z;WKST=SU;BYDAY=TU,TH

       or

       RRULE:FREQ=WEEKLY;COUNT=10;WKST=SU;BYDAY=TU,TH

       ==> (1997 9:00 AM EDT) September 2,4,9,11,16,18,23,25,30;
                              October 2
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-04 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-09 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-16 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-18 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-23 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-25 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-02 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description + CHR(13) + CHR(10) + "(option 1)", "RRULE:FREQ=WEEKLY;UNTIL=19971007T000000Z;WKST=SU;BYDAY=TU,TH", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

ClearVerification()

IF !VerifyCalculation(m.Timezone, m.Description + CHR(13) + CHR(10) + "(option 2)", "RRULE:FREQ=WEEKLY;COUNT=10;WKST=SU;BYDAY=TU,TH", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every other week on Monday, Wednesday, and Friday until December
      24, 1997, starting on Monday, September 1, 1997:

       DTSTART;TZID=America/New_York:19970901T090000
       RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL=19971224T000000Z;WKST=SU;
        BYDAY=MO,WE,FR

       ==> (1997 9:00 AM EDT) September 1,3,5,15,17,19,29;
                              October 1,3,13,15,17
           (1997 9:00 AM EST) October 27,29,31;
                              November 10,12,14,24,26,28;
                              December 8,10,12,22
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-01 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-17 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-17 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-27 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-31 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-14 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-24 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-26 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-28 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-08 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-22 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=WEEKLY;INTERVAL=2;UNTIL=19971224T000000Z;WKST=SU;BYDAY=MO,WE,FR", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every other week on Tuesday and Thursday, for 8 occurrences:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=8;WKST=SU;BYDAY=TU,TH

       ==> (1997 9:00 AM EDT) September 2,4,16,18,30;
                              October 2,14,16
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-04 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-16 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-18 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-14 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-16 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=8;WKST=SU;BYDAY=TU,TH", m.StartDate, 8)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .T. OR TESTALL

TEXT TO m.Description NOSHOW
      Monthly on the first Friday for 10 occurrences:

       DTSTART;TZID=America/New_York:19970905T090000
       RRULE:FREQ=MONTHLY;COUNT=10;BYDAY=1FR

       ==> (1997 9:00 AM EDT) September 5;October 3
           (1997 9:00 AM EST) November 7;December 5
           (1998 9:00 AM EST) January 2;February 6;March 6;April 3
           (1998 9:00 AM EDT) May 1;June 5
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-05 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-02-06 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-06 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-04-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-05-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-06-05 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;COUNT=10;BYDAY=1FR", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Monthly on the first Friday until December 24, 1997:

       DTSTART;TZID=America/New_York:19970905T090000
       RRULE:FREQ=MONTHLY;UNTIL=19971224T000000Z;BYDAY=1FR

       ==> (1997 9:00 AM EDT) September 5; October 3
           (1997 9:00 AM EST) November 7; December 5
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-05 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-05 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;UNTIL=19971224T000000Z;BYDAY=1FR", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every other month on the first and last Sunday of the month for 10
      occurrences:

       DTSTART;TZID=America/New_York:19970907T090000
       RRULE:FREQ=MONTHLY;INTERVAL=2;COUNT=10;BYDAY=1SU,-1SU

       ==> (1997 9:00 AM EDT) September 7,28
           (1997 9:00 AM EST) November 2,30
           (1998 9:00 AM EST) January 4,25;March 1,29
           (1998 9:00 AM EDT) May 3,31
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-07 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-28 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-04 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-25 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-05-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-05-31 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;INTERVAL=2;COUNT=10;BYDAY=1SU,-1SU", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Monthly on the second-to-last Monday of the month for 6 months:

       DTSTART;TZID=America/New_York:19970922T090000
       RRULE:FREQ=MONTHLY;COUNT=6;BYDAY=-2MO

       ==> (1997 9:00 AM EDT) September 22;October 20
           (1997 9:00 AM EST) November 17;December 22
           (1998 9:00 AM EST) January 19;February 16
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-22 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-22 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-20 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-17 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-22 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-02-16 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;COUNT=6;BYDAY=-2MO", m.StartDate, 6)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Monthly on the third-to-the-last day of the month, forever:

       DTSTART;TZID=America/New_York:19970928T090000
       RRULE:FREQ=MONTHLY;BYMONTHDAY=-3

       ==> (1997 9:00 AM EDT) September 28
           (1997 9:00 AM EST) October 29;November 28;December 29
           (1998 9:00 AM EST) January 29;February 26
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-28 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-28 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-28 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-02-26 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;BYMONTHDAY=-3", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Monthly on the 2nd and 15th of the month for 10 occurrences:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=MONTHLY;COUNT=10;BYMONTHDAY=2,15

       ==> (1997 9:00 AM EDT) September 2,15;October 2,15
           (1997 9:00 AM EST) November 2,15;December 2,15
           (1998 9:00 AM EST) January 2,15
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-15 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;COUNT=10;BYMONTHDAY=2,15", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
     Monthly on the first and last day of the month for 10 occurrences:

       DTSTART;TZID=America/New_York:19970930T090000
       RRULE:FREQ=MONTHLY;COUNT=10;BYMONTHDAY=1,-1

       ==> (1997 9:00 AM EDT) September 30;October 1
           (1997 9:00 AM EST) October 31;November 1,30;December 1,31
           (1998 9:00 AM EST) January 1,31;February 1
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-30 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-31 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-31 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-31 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-02-01 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;COUNT=10;BYMONTHDAY=1,-1", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every 18 months on the 10th thru 15th of the month for 10
      occurrences:

       DTSTART;TZID=America/New_York:19970910T090000
       RRULE:FREQ=MONTHLY;INTERVAL=18;COUNT=10;BYMONTHDAY=10,11,12,
        13,14,15

       ==> (1997 9:00 AM EDT) September 10,11,12,13,14,15
           (1999 9:00 AM EST) March 10,11,12,13
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-10 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-14 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-13 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;INTERVAL=18;COUNT=10;BYMONTHDAY=10,11,12,13,14,15", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every Tuesday, every other month:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=MONTHLY;INTERVAL=2;BYDAY=TU

       ==> (1997 9:00 AM EDT) September 2,9,16,23,30
           (1997 9:00 AM EST) November 4,11,18,25
           (1998 9:00 AM EST) January 6,13,20,27;March 3,10,17,24,31
          ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-09 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-16 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-23 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-04 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-18 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-25 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-06 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-20 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-27 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-17 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-24 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-31 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;INTERVAL=2;BYDAY=TU", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Yearly in June and July for 10 occurrences:

       DTSTART;TZID=America/New_York:19970610T090000
       RRULE:FREQ=YEARLY;COUNT=10;BYMONTH=6,7

       ==> (1997 9:00 AM EDT) June 10;July 10
           (1998 9:00 AM EDT) June 10;July 10
           (1999 9:00 AM EDT) June 10;July 10
           (2000 9:00 AM EDT) June 10;July 10
           (2001 9:00 AM EDT) June 10;July 10

         Note: Since none of the BYDAY, BYMONTHDAY, or BYYEARDAY
         components are specified, the day is gotten from "DTSTART".
ENDTEXT

ResetVerification()

m.StartDate = {^1997-06-10 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-06-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-07-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-06-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-07-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-06-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-07-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2000-06-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2000-07-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2001-06-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2001-07-10 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;COUNT=10;BYMONTH=6,7", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every other year on January, February, and March for 10
      occurrences:

       DTSTART;TZID=America/New_York:19970310T090000
       RRULE:FREQ=YEARLY;INTERVAL=2;COUNT=10;BYMONTH=1,2,3

       ==> (1997 9:00 AM EST) March 10
           (1999 9:00 AM EST) January 10;February 10;March 10
           (2001 9:00 AM EST) January 10;February 10;March 10
           (2003 9:00 AM EST) January 10;February 10;March 10
ENDTEXT

ResetVerification()

m.StartDate = {^1997-03-10 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-03-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-01-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-02-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2001-01-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2001-02-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2001-03-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2003-01-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2003-02-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2003-03-10 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;INTERVAL=2;COUNT=10;BYMONTH=1,2,3", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every third year on the 1st, 100th, and 200th day for 10
      occurrences:

       DTSTART;TZID=America/New_York:19970101T090000
       RRULE:FREQ=YEARLY;INTERVAL=3;COUNT=10;BYYEARDAY=1,100,200

       ==> (1997 9:00 AM EST) January 1
           (1997 9:00 AM EDT) April 10;July 19
           (2000 9:00 AM EST) January 1
           (2000 9:00 AM EDT) April 9;July 18
           (2003 9:00 AM EST) January 1
           (2003 9:00 AM EDT) April 10;July 19
           (2006 9:00 AM EST) January 1
ENDTEXT

ResetVerification()

m.StartDate = {^1997-01-01 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-01-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-04-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-07-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2000-01-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2000-04-09 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2000-07-18 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2003-01-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2003-04-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2003-07-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2006-01-01 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;INTERVAL=3;COUNT=10;BYYEARDAY=1,100,200", m.StartDate, 10)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
     Every 20th Monday of the year, forever:

       DTSTART;TZID=America/New_York:19970519T090000
       RRULE:FREQ=YEARLY;BYDAY=20MO

       ==> (1997 9:00 AM EDT) May 19
           (1998 9:00 AM EDT) May 18
           (1999 9:00 AM EDT) May 17
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-05-19 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-05-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-05-18 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-05-17 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;BYDAY=20MO", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
 Monday of week number 20 (where the default start of the week is
      Monday), forever:

       DTSTART;TZID=America/New_York:19970512T090000
       RRULE:FREQ=YEARLY;BYWEEKNO=20;BYDAY=MO

       ==> (1997 9:00 AM EDT) May 12
           (1998 9:00 AM EDT) May 11
           (1999 9:00 AM EDT) May 17
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-05-12 09:00:00}

INSERT INTO Verification (Expected) VALUES ({^1997-05-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-05-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-05-17 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;BYWEEKNO=20;BYDAY=MO", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every Thursday in March, forever:

       DTSTART;TZID=America/New_York:19970313T090000
       RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=TH

       ==> (1997 9:00 AM EST) March 13,20,27
           (1998 9:00 AM EST) March 5,12,19,26
           (1999 9:00 AM EST) March 4,11,18,25
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-03-13 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-03-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-03-20 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-03-27 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-26 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-04 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-18 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-03-25 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=TH", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every Thursday, but only during June, July, and August, forever:

       DTSTART;TZID=America/New_York:19970605T090000
       RRULE:FREQ=YEARLY;BYDAY=TH;BYMONTH=6,7,8

       ==> (1997 9:00 AM EDT) June 5,12,19,26;July 3,10,17,24,31;
                              August 7,14,21,28
           (1998 9:00 AM EDT) June 4,11,18,25;July 2,9,16,23,30;
                              August 6,13,20,27
           (1999 9:00 AM EDT) June 3,10,17,24;July 1,8,15,22,29;
                              August 5,12,19,26
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-06-05 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-06-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-06-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-06-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-06-26 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-07-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-07-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-07-17 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-07-24 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-07-31 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-14 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-21 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-28 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-06-04 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-06-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-06-18 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-06-25 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-07-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-07-09 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-07-16 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-07-23 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-07-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-08-06 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-08-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-08-20 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-08-27 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-06-03 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-06-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-06-17 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-06-24 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-07-01 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-07-08 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-07-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-07-22 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-07-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-08-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-08-12 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-08-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-08-26 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;BYDAY=TH;BYMONTH=6,7,8", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every Friday the 13th, forever:

       DTSTART;TZID=America/New_York:19970902T090000
       EXDATE;TZID=America/New_York:19970902T090000
       RRULE:FREQ=MONTHLY;BYDAY=FR;BYMONTHDAY=13

       ==> (1998 9:00 AM EST) February 13;March 13;November 13
           (1999 9:00 AM EDT) August 13
           (2000 9:00 AM EDT) October 13
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1998-02-13 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1998-02-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-11-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1999-08-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2000-10-13 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;BYDAY=FR;BYMONTHDAY=13", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      The first Saturday that follows the first Sunday of the month,
      forever:

       DTSTART;TZID=America/New_York:19970913T090000
       RRULE:FREQ=MONTHLY;BYDAY=SA;BYMONTHDAY=7,8,9,10,11,12,13

       ==> (1997 9:00 AM EDT) September 13;October 11
           (1997 9:00 AM EST) November 8;December 13
           (1998 9:00 AM EST) January 10;February 7;March 7
           (1998 9:00 AM EDT) April 11;May 9;June 13...
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-13 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-08 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-13 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-02-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-04-11 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-05-09 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-06-13 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;BYDAY=SA;BYMONTHDAY=7,8,9,10,11,12,13", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every 4 years, the first Tuesday after a Monday in November,
      forever (U.S. Presidential Election day):

       DTSTART;TZID=America/New_York:19961105T090000
       RRULE:FREQ=YEARLY;INTERVAL=4;BYMONTH=11;BYDAY=TU;
        BYMONTHDAY=2,3,4,5,6,7,8

        ==> (1996 9:00 AM EST) November 5
            (2000 9:00 AM EST) November 7
            (2004 9:00 AM EST) November 2
            ...
ENDTEXT

ResetVerification()

m.StartDate = {^1996-11-05 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1996-11-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2000-11-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2004-11-02 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=YEARLY;INTERVAL=4;BYMONTH=11;BYDAY=TU;BYMONTHDAY=2,3,4,5,6,7,8", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      The third instance into the month of one of Tuesday, Wednesday, or
      Thursday, for the next 3 months:

       DTSTART;TZID=America/New_York:19970904T090000
       RRULE:FREQ=MONTHLY;COUNT=3;BYDAY=TU,WE,TH;BYSETPOS=3

       ==> (1997 9:00 AM EDT) September 4;October 7
           (1997 9:00 AM EST) November 6
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-04 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-04 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-07 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-06 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;COUNT=3;BYDAY=TU,WE,TH;BYSETPOS=3", m.StartDate, 3)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      The second-to-last weekday of the month:

       DTSTART;TZID=America/New_York:19970929T090000
       RRULE:FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-2

       ==> (1997 9:00 AM EDT) September 29
           (1997 9:00 AM EST) October 30;November 27;December 30
           (1998 9:00 AM EST) January 29;February 26;March 30
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-29 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-10-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-11-27 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-12-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-01-29 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-02-26 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1998-03-30 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;BYDAY=MO,TU,WE,TH,FR;BYSETPOS=-2", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every 3 hours from 9:00 AM to 5:00 PM on a specific day:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=HOURLY;INTERVAL=3;UNTIL=19970902T170000Z

       ==> (September 2, 1997 EDT) 09:00,12:00,15:00
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 12:00:00})

* the next insert, although listed in the example results, shouldn't occur:
* the UNTIL part rule is expressed in UTC time, and 19970902T150000 EDT is in fact > 19970902T170000Z
*!*	INSERT INTO Verification (Expected) VALUES ({^1997-09-02 15:00:00})

IF !VerifyCalculation(m.Timezone, m.Description + CHR(13) + CHR(10) + "(see comments in code)", "RRULE:FREQ=HOURLY;INTERVAL=3;UNTIL=19970902T170000Z", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every 15 minutes for 6 occurrences:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=MINUTELY;INTERVAL=15;COUNT=6

       ==> (September 2, 1997 EDT) 09:00,09:15,09:30,09:45,10:00,10:15
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:15:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:30:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:45:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 10:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 10:15:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MINUTELY;INTERVAL=15;COUNT=6", m.StartDate, 6)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every hour and a half for 4 occurrences:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=MINUTELY;INTERVAL=90;COUNT=4

       ==> (September 2, 1997 EDT) 09:00,10:30;12:00;13:30
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 10:30:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 12:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-09-02 13:30:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MINUTELY;INTERVAL=90;COUNT=4", m.StartDate, 4)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      Every 20 minutes from 9:00 AM to 4:40 PM every day:

       DTSTART;TZID=America/New_York:19970902T090000
       RRULE:FREQ=DAILY;BYHOUR=9,10,11,12,13,14,15,16;BYMINUTE=0,20,40
       or
       RRULE:FREQ=MINUTELY;INTERVAL=20;BYHOUR=9,10,11,12,13,14,15,16

       ==> (September 2, 1997 EDT) 9:00,9:20,9:40,10:00,10:20,
                                   ... 16:00,16:20,16:40
           (September 3, 1997 EDT) 9:00,9:20,9:40,10:00,10:20,
                                   ...16:00,16:20,16:40
           ...
ENDTEXT

ResetVerification()

m.StartDate = {^1997-09-02 09:00:00}
m.LoopDate = m.StartDate
DO WHILE m.LoopDate < {^1997-09-02 17:00:00}
	INSERT INTO Verification (Expected) VALUES (m.LoopDate)
	m.LoopDate = m.LoopDate + 20 * 60
ENDDO
m.LoopDate = m.StartDate + 86400
DO WHILE m.LoopDate < {^1997-09-03 17:00:00}
	INSERT INTO Verification (Expected) VALUES (m.LoopDate)
	m.LoopDate = m.LoopDate + 20 * 60
ENDDO

IF !VerifyCalculation(m.Timezone, m.Description + CHR(13) + CHR(10) + "(option 1)", "RRULE:FREQ=DAILY;BYHOUR=9,10,11,12,13,14,15,16;BYMINUTE=0,20,40", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

ClearVerification()

IF !VerifyCalculation(m.Timezone, m.Description + CHR(13) + CHR(10) + "(option 2)", "RRULE:FREQ=MINUTELY;INTERVAL=20;BYHOUR=9,10,11,12,13,14,15,16", m.StartDate)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      An example where the days generated makes a difference because of
      WKST:

       DTSTART;TZID=America/New_York:19970805T090000
       RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=4;BYDAY=TU,SU;WKST=MO

       ==> (1997 EDT) August 5,10,19,24
ENDTEXT

ResetVerification()

m.StartDate = {^1997-08-05 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-08-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-10 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-24 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=4;BYDAY=TU,SU;WKST=MO", m.StartDate, 4)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      changing only WKST from MO to SU, yields different results...

       DTSTART;TZID=America/New_York:19970805T090000
       RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=4;BYDAY=TU,SU;WKST=SU

       ==> (1997 EDT) August 5,17,19,31
ENDTEXT

ResetVerification()

m.StartDate = {^1997-08-05 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^1997-08-05 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-17 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-19 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^1997-08-31 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=WEEKLY;INTERVAL=2;COUNT=4;BYDAY=TU,SU;WKST=SU", m.StartDate, 4)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

#IF .F. OR TESTALL

TEXT TO m.Description NOSHOW
      An example where an invalid date (i.e., February 30) is ignored.

       DTSTART;TZID=America/New_York:20070115T090000
       RRULE:FREQ=MONTHLY;BYMONTHDAY=15,30;COUNT=5

       ==> (2007 EST) January 15,30
           (2007 EST) February 15
           (2007 EDT) March 15,30
ENDTEXT

ResetVerification()

m.StartDate = {^2007-01-15 09:00:00}
INSERT INTO Verification (Expected) VALUES ({^2007-01-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2007-01-30 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2007-02-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2007-03-15 09:00:00})
INSERT INTO Verification (Expected) VALUES ({^2007-03-30 09:00:00})

IF !VerifyCalculation(m.Timezone, m.Description, "RRULE:FREQ=MONTHLY;BYMONTHDAY=15,30;COUNT=5", m.StartDate, 5)
	m.NumberOfFails = m.NumberOfFails + 1
ENDIF
m.NumberOfTests = m.NumberOfTests + 1

#ENDIF

*****************************************************************************************************************************************************************************
MESSAGEBOX(TEXTMERGE("RRULE processor tests concluded<<CHR(13)+CHR(10)>>Tests ran: <<m.NumberOfTests>><<CHR(13)+CHR(10)>>Fails: <<m.NumberOfFails>>"), 64, "RRULE processor")
*****************************************************************************************************************************************************************************

FUNCTION ResetVerification ()

	USE IN SELECT("Verification")
	CREATE CURSOR Verification (Expected Datetime, Calculated Datetime, TzName Varchar(32))

ENDFUNC

FUNCTION ClearVerification ()

	UPDATE Verification SET Calculated = {}

ENDFUNC

FUNCTION VerifyCalculation (TZ AS iCalCompVTIMEZONE, Test AS String, RuleDefinition AS String, StartDate AS Datetime, MaxCount AS Integer) AS Logical

	LOCAL Fail AS Logical
	LOCAL Rule AS iCalPropRRULE
	LOCAL RecurCursor AS String

	m.Rule = CREATEOBJECT("iCalPropRRULE")
	m.Rule.Parse(m.RuleDefinition)

	m.RecurCursor = m.Rule.CalculateAll(m.StartDate, DATETIME(), m.TZ, .NULL., .NULL.)

	GO TOP IN "Verification"

	SELECT (m.RecurCursor)
	SCAN FOR RECNO() <= RECCOUNT("Verification")
		REPLACE Verification.Calculated WITH LocalTime, Verification.TzName WITH NVL(TzName, "")
		SKIP IN Verification
	ENDSCAN

	SELECT Verification
	LOCATE FOR Expected != Calculated
	m.Fail = FOUND() OR (!EMPTY(m.MaxCount) AND m.MaxCount != RECCOUNT(m.RecurCursor))

	IF VERBOSE OR (m.Fail AND VERBOSEFAIL)
		GO TOP
		BROWSE LAST NOWAIT

		IF MESSAGEBOX(TEXTMERGE("<<m.Test>><<REPLICATE(CHR(13)+CHR(10),2)>>Test <<IIF(m.Fail, 'failed', 'succeeded')>>."), IIF(m.Fail, 49, 65), "RRULE processor") = 2
			CANCEL
		ENDIF
	ENDIF

	IF !m.Fail
 		USE IN (m.RecurCursor)
 	ENDIF
 
 	RETURN !m.Fail
 
ENDFUNC
