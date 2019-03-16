*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
*!*	|                                                                    |
*!*	|    Source and docs: https://bitbucket.org/atlopes/ical4vfp         |
*!*	|                                                                    |
*!*	+--------------------------------------------------------------------+

*!*	iCalendar parameters sub-classes

* install dependencies
IF _VFP.StartMode = 0
	SET PATH TO (JUSTPATH(SYS(16))) ADDITIVE
ENDIF
DO "icalendar.prg"

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

DEFINE CLASS iCalParmALTREP AS _iCalParameter

	ICName = "ALTREP"
	xICName = "altrep"

	Value_DataType = "URI"

ENDDEFINE

DEFINE CLASS iCalParmCN AS _iCalParameter

	ICName = "CN"
	xICName = "cn"

ENDDEFINE

DEFINE CLASS iCalParmCUTYPE AS _iCalParameter

	ICName = "CUTYPE"
	xICName = "cutype"

	DefaultValue = "INDIVIDUAL"

	Enumeration = "INDIVIDUAL,GROUP,RESOURCE,ROOM,UNKNOWN"

ENDDEFINE

DEFINE CLASS iCalParmDELEGATED_FROM AS _iCalParameter

	ICName = "DELEGATED-FROM"
	xICName = "delegated-from"

	Value_DataType = "CAL-ADDRESS"
	Value_IsList = .T.

ENDDEFINE

DEFINE CLASS iCalParmDELEGATED_TO AS _iCalParameter

	ICName = "DELEGATED-TO"
	xICName = "delegated-to"

	Value_DataType = "CAL-ADDRESS"
	Value_IsList = .T.

ENDDEFINE

DEFINE CLASS iCalParmDIR AS _iCalParameter

	ICName = "DIR"
	xICName = "dir"

	Value_DataType = "URI"

ENDDEFINE

DEFINE CLASS iCalParmENCODING AS _iCalParameter

	ICName = "ENCODING"
	xICName = "encoding"

	Enumeration = "BIT8,BASE64"
	Extensions = .F.

ENDDEFINE

DEFINE CLASS iCalParmFMTTYPE AS _iCalParameter

	ICName = "FMTTYPE"
	xICName = "fmttype"

ENDDEFINE

DEFINE CLASS iCalParmFBTYPE AS _iCalParameter

	ICName = "FBTYPE"
	xICName = "fbtype"

	Enumeration = "FREE,BUSY,BUSY-UNAVAILABLE,BUSY-TENTATIVE"

ENDDEFINE

DEFINE CLASS iCalParmLANGUAGE AS _iCalParameter

	ICName = "LANGUAGE"
	xICName = "language"

	DefaultValue = "EN"

ENDDEFINE

DEFINE CLASS iCalParmMEMBER AS _iCalParameter

	ICName = "MEMBER"
	xICName = "member"

	Value_DataType = "CAL-ADDRESS"
	Value_IsList = .T.

ENDDEFINE

DEFINE CLASS iCalParmPARTSTAT AS _iCalParameter

	ICName = "PARTSTAT"
	xICName = "partstat"
	AlternativeClasses = "VEVENT,VJOURNAL,VTODO"

	Enumeration = "NEEDS-ACTION,ACCEPTED,DECLINED,TENTATIVE,DELEGATED,COMPLETED,IN-PROCESS"

ENDDEFINE

DEFINE CLASS iCalParmPARTSTAT_VEvent AS iCalParmPARTSTAT

	Enumeration = "NEEDS-ACTION,ACCEPTED,DECLINED,TENTATIVE,DELEGATED"

ENDDEFINE

DEFINE CLASS iCalParmPARTSTAT_VJournal AS iCalParmPARTSTAT

	Enumeration = "NEEDS-ACTION,ACCEPTED,DECLINED"

ENDDEFINE

DEFINE CLASS iCalParmPARTSTAT_VToDo AS iCalParmPARTSTAT

ENDDEFINE

DEFINE CLASS iCalParmRANGE AS _iCalParameter

	ICName = "RANGE"
	xICName = "range"

	DefaultValue = "THISANDFUTURE"

	ReadOnly = .T.

ENDDEFINE

DEFINE CLASS iCalParmRELATED AS _iCalParameter

	ICName = "RELATED"
	xICName = "related"

	Enumeration = "START,END"
	Extensions = .F.

ENDDEFINE

DEFINE CLASS iCalParmRELTYPE AS _iCalParameter

	ICName = "RELTYPE"
	xICName = "reltype"

	Enumeration = "PARENT,CHILD,SIBLING"

ENDDEFINE

DEFINE CLASS iCalParmROLE AS _iCalParameter

	ICName = "ROLE"
	xICName = "role"

	Enumeration = "CHAIR,REQ-PARTICIPANT,OPT-PARTICIPANT,NON-PARTICIPANT"

ENDDEFINE

DEFINE CLASS iCalParmRSVP AS _iCalParameter

	ICName = "RSVP"
	xICName = "rsvp"

	DefaultValue = .F.
	Value_DataType = "BOOLEAN"

ENDDEFINE

DEFINE CLASS iCalParmSENT_BY AS _iCalParameter

	ICName = "SENT-BY"
	xICName = "sent-by"

	Value_DataType = "CAL-ADDRESS"

ENDDEFINE

DEFINE CLASS iCalParmTZID AS _iCalParameter

	ICName = "TZID"
	xICName = "tzid"

ENDDEFINE

DEFINE CLASS iCalParmVALUE AS _iCalParameter

	ICName = "VALUE"
	xICName = "value"

	DefaultValue = "TEXT"
	Enumeration = "BINARY,BOOLEAN,CAL-ADDRESS,DATE,DATE-TIME,DURATION,FLOAT,INTEGER,PERIOD,RECUR,TEXT,TIME,URI,UTC-OFFSET"

ENDDEFINE
