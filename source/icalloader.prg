*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
*!*	|                                                                    |
*!*	|    Source and docs: https://bitbucket.org/atlopes/ical4vfp         |
*!*	|                                                                    |
*!*	+--------------------------------------------------------------------+

*!*	VFP iCalendar classes loader

IF _VFP.StartMode = 0
	DO LOCFILE("icsprocessor.prg")
ELSE
	DO "icsprocessor.prg"
ENDIF

DO "icalcomponents.prg"
DO "icalproperties.prg"
DO "icalparameters.prg"
