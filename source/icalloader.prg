*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
*!*	|                                                                    |
*!*	+--------------------------------------------------------------------+

*!*	VFP iCalendar classes loader
IF _VFP.StartMode = 0
	SET PATH TO (JUSTPATH(SYS(16))) ADDITIVE
ENDIF

DO "icsprocessor.prg"
DO "icalcomponents.prg"
DO "icalproperties.prg"
DO "icalparameters.prg"
