*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
*!*	|                                                                    |
*!*	|    Source and docs: https://bitbucket.org/atlopes/ical4vfp         |
*!*	|                                                                    |
*!*	+--------------------------------------------------------------------+

*!*	iCalendar components sub-classes

* install dependencies
IF _VFP.StartMode = 0
	SET PATH TO (JUSTPATH(SYS(16))) ADDITIVE
ENDIF
DO "icalendar.prg"

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

#DEFINE	SAFETHIS		ASSERT !USED("This") AND TYPE("This") == "O"

* the iCalendar Core Object (top-level container)

DEFINE CLASS iCalendar AS _iCalComponent

	IncludeComponents = "*ANY"

	ICName = "VCALENDAR"
	xICName = "vcalendar"

	Level = "CORE"

	HIDDEN Started
	Started = .F.

	_MemberData =	'<VFPData>' + ;
							'<memberdata name="gettimezone" type="method" display="GetTimezone"/>' + ;
						'</VFPData>'

	FUNCTION Parse (Serialized AS String)

		SAFETHIS

		ASSERT VARTYPE(m.Serialized) == "C"

		IF !This.Started
			This.Started = UPPER(TRIM(m.Serialized)) == "BEGIN:VCALENDAR"
			IF !This.Started
				RETURN .F.
			ELSE
				This.Reset()
			ENDIF
		ELSE
			RETURN DODEFAULT(m.Serialized)
		ENDIF

	ENDFUNC

	* get a timezone defined in the iCalendar core object
	FUNCTION GetTimezone (TzID AS String)

		ASSERT PCOUNT() = 0 OR VARTYPE(m.TzID) == "C"

		LOCAL TzCount AS Integer
		LOCAL TzIndex AS Integer
		LOCAL Timezone AS iCalCompVTIMEZONE

		* get the first (= usually, the only one)
		IF PCOUNT() = 0

			RETURN This.GetICComponent("VTIMEZONE")

		ELSE

			* or look for a specific TZID
			FOR m.TzIndex = 1 TO This.GetICComponentsCount("VTIMEZONE")

				m.Timezone = This.GetICComponent("VTIMEZONE", m.TzIndex)
				IF m.Timezone.GetICPropertyValue("TZID") == m.TzID
					RETURN m.Timezone
				ENDIF
				m.Timezone = .NULL.
			ENDFOR

		ENDIF

		* none found...
		RETURN .NULL.

	ENDFUNC

ENDDEFINE

DEFINE CLASS iCalCompVEVENT AS _iCalComponent

	ICName = "VEVENT"
	xICName = "vevent"

	IncludeComponents = "VALARM"

ENDDEFINE

DEFINE CLASS iCalCompVTODO AS _iCalComponent

	ICName = "VTODO"
	xICName = "vtodo"

	IncludeComponents = "VALARM"

ENDDEFINE

DEFINE CLASS iCalCompVJOURNAL AS _iCalComponent

	ICName = "VJOURNAL"
	xICName = "vjournal"

ENDDEFINE

DEFINE CLASS iCalCompVFREEBUSY AS _iCalComponent

	ICName = "VFREEBUSY"
	xICName = "vfreebusy"

ENDDEFINE

DEFINE CLASS iCalCompVTIMEZONE AS _iCalComponent

	ICName = "VTIMEZONE"
	xICName = "vtimezone"

	IncludeComponents = "STANDARD,DAYLIGHT"

	TzName = ""
	TzEntry = 0

	HIDDEN StartST, EndST, BiasST, _StartST(1), _EndST(1), _BiasST(1), _TzName(1), _TzEntry(1), StackST
	StartST = .NULL.
	EndST = .NULL.
	BiasST = 0
	DIMENSION _StartST(1)
	DIMENSION _EndST(1)
	DIMENSION _BiasST(1)
	DIMENSION _TzName(1)
	DIMENSION _TzEntry(1)

	StackST = 0

	_MemberData =	'<VFPData>' + ;
							'<memberdata name="tzname" type="property" display="TzName"/>' + ;
							'<memberdata name="nextsavingtimechange" type="method" display="NextSavingTimeChange"/>' + ;
							'<memberdata name="popsavingtime" type="method" display="PopSavingTime"/>' + ;
							'<memberdata name="pushsavingtime" type="method" display="PushSavingTime"/>' + ;
							'<memberdata name="tolocaltime" type="method" display="ToLocalTime"/>' + ;
							'<memberdata name="toutc" type="method" display="ToUTC"/>' + ;
							'<memberdata name="utcoffset" type="method" display="UTCOffset"/>' + ;
					'</VFPData>'

	* save or restore saving time current settings
	FUNCTION PushSavingTime ()
		This.StackST = This.StackST + 1
		DIMENSION This._StartST(This.StackST)
		This._StartST(This.StackST) = This.StartST
		DIMENSION This._EndST(This.StackST)
		This._EndST(This.StackST) = This.EndST
		DIMENSION This._BiasST(This.StackST)
		This._BiasST(This.StackST) = This.BiasST
		DIMENSION This._TzName(This.StackST)
		This._TzName(This.StackST) = This.TzName
		DIMENSION This._TzEntry(This.StackST)
		This._TzEntry(This.StackST) = This.TzEntry
	ENDFUNC
	FUNCTION PopSavingTime ()
		IF This.StackST > 0
			This.StartST = This._StartST(This.StackST)
			This.EndST = This._EndST(This.StackST)
			This.BiasST = This._BiasST(This.StackST)
			This.TzName = This._TzName(This.StackST)
			This.TzEntry = This._TzEntry(This.StackST)
			This.StackST = This.StackST - 1
			IF This.StackST != 0
				DIMENSION This._StartST(This.StackST)
				DIMENSION This._EndST(This.StackST)
				DIMENSION This._BiasST(This.StackST)
				DIMENSION This._TzName(This.StackST)
				DIMENSION This._TzEntry(This.StackST)
			ENDIF
		 ENDIF
	ENDFUNC

	* return the UTC offset (in seconds) for a given time
	FUNCTION UTCOffset (RefDate AS Datetime) AS Integer

		ASSERT VARTYPE(m.RefDate) == "T"

		RETURN m.RefDate - This.ToUTC(m.RefDate)

	ENDFUNC
	
	* take a local time, and make it UTC
	FUNCTION ToUTC (LocalTime AS Datetime)

		ASSERT VARTYPE(m.LocalTime) == "T"

		RETURN This._UTC(m.LocalTime, .T.)

	ENDFUNC

	* take a UTC time, and make it local
	FUNCTION ToLocalTime (UTCTime AS Datetime)

		ASSERT VARTYPE(m.UTCTime) == "T"

		RETURN This._UTC(m.UTCTime, .F.)

	ENDFUNC

	HIDDEN FUNCTION _UTC (RefTime AS Datetime, ToUTC AS Logical)

		SAFETHIS

		LOCAL OffsetTo AS Number
		LOCAL ClosestDate AS Datetime
		LOCAL RefDate AS Datetime
		LOCAL UntilDate AS Datetime
		LOCAL TzComp AS _iCalComponent
		LOCAL CompIndex AS Integer
		LOCAL RRule AS iCalPropRRULE
		LOCAL VRule AS iCalTypeRECUR
		LOCAL Period AS Logical

		* make sure we are working with datetimes
		IF VARTYPE(m.RefTime) == "D"
			RETURN DTOT(m.RefTime)
		ENDIF

		* the source time fits in the ST period set in a previous calculation?
		IF !ISNULL(This.StartST) ;
				AND ((m.ToUTC AND BETWEEN(m.RefTime, This.StartST, NVL(This.EndST, m.RefTime))) OR ;
					(!m.ToUTC AND BETWEEN(m.RefTime + This.BiasST * 60, This.StartST, NVL(This.EndST, m.RefTime + This.BiasST * 60))))

			* use the stored bias, no need to go through the calculation, again
			m.OffsetTo = This.BiasST

		ELSE

			* calculate the offset to UTC
			m.OffsetTo =  0
			m.ClosestDate = {^0001-01-01}
			This.TzName = ""

			FOR m.CompIndex = 1 TO This.GetICComponentsCount()

				m.TzComp = This.GetICComponent(m.CompIndex)

				* look for all STANDARD and DAYLIGHT definitions
				IF m.TzComp.ICName == "STANDARD" OR m.TzComp.ICName == "DAYLIGHT"

					* get the date when this tz definition went in effect
					m.RefDate = m.TzComp.GetICPropertyValue("DTSTART")
					* most probably, there will a recurrence rule for it
					m.RRule = m.TzComp.GetICProperty("RRULE")
					m.Period = .F.
					IF !ISNULL(m.RRule)
						m.VRule = m.RRule.Getvalue()
						* a quick check on Until rule part
						* if it has already passed, no need to check for it (it expired)
						m.UntilDate = m.VRule.Until
						* but if it didn't calculate the previous and next ST enforcement
						IF ISNULL(m.UntilDate) OR YEAR(m.UntilDate) > YEAR(m.RefTime)
							m.RRule.CalculatePeriod(m.RefDate, m.RefTime, .NULL., m.TzComp.GetICProperty("RDATE", -1), m.TzComp.GetICProperty("EXDATE", -1))
							IF !ISNULL(m.RRule.PreviousDate)
								m.RefDate = m.RRule.PreviousDate
								m.Period = .T.
							ELSE
								m.RefDate = .NULL.	&& this shouldn't happen... but something may be wrong with the rule, or the processor...
							ENDIF
						ELSE
							m.RefDate = .NULL.
						ENDIF
					* no RRULE? look for RDATEs
					ELSE
						m.RefDate = This.GetClosestRDATE(m.RefTime, m.RefDate, m.TzComp.GetICProperty("RDATE", -1))
					ENDIF

					* we have a date, and it covers our time, and it's the closest to our time that we found so far?
					IF !ISNULL(m.RefDate) AND !EMPTY(m.RefDate) AND m.RefDate < m.RefTime AND m.RefDate > m.ClosestDate
						m.ClosestDate = m.RefDate
						m.OffsetTo = m.TzComp.GetICPropertyValue("TZOFFSETTO")
						This.TzName = NVL(m.TzComp.GetICPropertyValue("TZNAME"), "")
						* store the results of our calculation for the next calls
						IF m.Period
							This.BiasST = m.OffsetTo
							This.StartST = m.RefDate
						ELSE
							This.StartST = .NULL.
							This.EndST = .NULL.
							This.BiasST = 0
						ENDIF
					ENDIF
					* if we used a period in the last calculation
					IF m.Period
						* if the next occurence is before our current ST ending date, then move our ST ending period to that point
						IF ISNULL(This.EndST) OR (m.RRule.NextDate > This.StartST AND This.EndST > m.RRule.NextDate)
							This.EndST = m.RRule.NextDate
						ENDIF
					ENDIF
				ENDIF
			ENDFOR

		ENDIF
 
 		* we have a local time and want the UTC? subtract the bias
 		IF m.ToUTC	
 			m.CalcTime = m.RefTime - m.OffsetTo * 60
 		* we have a UTC and want the local time? add the bias
 		ELSE
 			m.CalcTime = m.RefTime + m.OffsetTo * 60
 		ENDIF

		* done
 		RETURN m.CalcTime

	ENDFUNC

	* calculate the datetime of the next saving time change
	FUNCTION NextSavingTimeChange (RefTime AS Datetime, Chained AS Logical)

		SAFETHIS

		ASSERT VARTYPE(m.RefTime) + VARTYPE(m.Chained) == "TL"

		LOCAL RefDate AS Datetime
		LOCAL NextDate AS Datetime
		LOCAL NextCalc AS Datetime
		LOCAL Closer AS Integer
		LOCAL CloserName AS String
		LOCAL UntilDate AS Datetime
		LOCAL TzComp AS _iCalComponent
		LOCAL CompIndex AS Integer
		LOCAL RRule AS iCalPropRRULE
		LOCAL VRule AS iCalTypeRECUR
		LOCAL RDates AS Collection
		LOCAL Period AS Logical

		* no ST change in the future
		m.NextDate = .NULL.
		m.Closer = .NULL.
		m.CloserName = .NULL.

		FOR m.CompIndex = 1 TO This.GetICComponentsCount()

			m.TzComp = This.GetICComponent(m.CompIndex)

			* look for all STANDARD and DAYLIGHT definitions
			IF m.TzComp.ICName == "STANDARD" OR m.TzComp.ICName == "DAYLIGHT"

				m.Period = .F.
				m.CalcNext = .NULL.
				m.RefDate = .NULL.

				* get the date when this tz definition went last in effect
				IF m.Chained AND !ISNULL(m.TzComp.LastChained)
					* but only if it wasn't calculated previously in the chain
					IF m.TzComp.LastChained > m.RefTime
						m.CalcNext = m.TzComp.LastChained
					ELSE
						m.RefDate = m.TzComp.LastChained
					ENDIF
				ELSE
					m.RefDate = m.TzComp.GetICPropertyValue("DTSTART")
					m.TzComp.LastChained = .NULL.
				ENDIF

				* not calculated, yet?
				IF ISNULL(m.RefDate) OR m.RefDate <= m.RefTime
					IF !m.Chained OR ISNULL(m.CalcNext)
						* most probably, there will a recurrence rule for it
						m.RRule = m.TzComp.GetICProperty("RRULE")
						IF !ISNULL(m.RRule)
							m.VRule = m.RRule.Getvalue()
							* a quick check on Until rule part
							* if it has already passed, no need to execute the rule (it has expired)
							m.UntilDate = m.VRule.Until
							* but if it didn't, calculate the previous and next ST enforcement
							IF ISNULL(m.UntilDate) OR YEAR(m.UntilDate) >= YEAR(m.RefTime)
								m.RRule.CalculatePeriod(m.RefDate, m.RefTime, .NULL., m.TzComp.GetICProperty("RDATE", -1), m.TzComp.GetICProperty("EXDATE", -1))
								IF !ISNULL(m.RRule.PreviousDate)
									m.RefDate = m.RRule.PreviousDate
									m.TzComp.LastChained = m.RRule.NextDate
									m.CalcNext = m.RRule.NextDate
									m.Period = .T.
								ELSE
									m.RefDate = .NULL.	&& this shouldn't happen... but something may be wrong with the rule, or the processor...
								ENDIF
							ELSE
								m.RefDate = .NULL.
							ENDIF
						* no RRULE? look for RDATEs
						ELSE
							m.RDates = This.GetSurroundingRDATEs(m.RefTime, m.RefDate, m.TzComp.GetICProperty("RDATE", -1))
							IF !ISNULL(m.RDates)
								m.RefDate = m.RDates.Item(1)
								m.CalcNext = m.RDates.Item(2)
								m.TzComp.LastChained = m.CalcNext
								m.Period = .T.
								m.RDates = .NULL.
							ENDIF
						ENDIF
					ELSE
						m.Period = .T.		&& a period was calculated previously in the chain
					ENDIF
				ENDIF

				* if we used a period in the last calculation
				IF m.Period
					* if the next occurence is before our current ST ending date, then move our ST ending period to that point
					IF ISNULL(m.NextDate) OR (m.CalcNext > NVL(m.RefDate, {}) AND m.NextDate > m.CalcNext)
						m.NextDate = m.CalcNext
					ENDIF
					* set the name of the TZ name the date is in (determined from the narrowest distance to each TZ beginning)
					IF !ISNULL(m.RefDate) AND (ISNULL(m.Closer) OR TTOD(m.RefTime) - TTOD(m.RefDate) < m.Closer)
						m.CloserName = NVL(m.TzComp.GetICPropertyValue("TZNAME"), "")
						m.Closer = TTOD(m.RefTime) - TTOD(m.RefDate)
					ENDIF
				ENDIF

			ENDIF
		ENDFOR

		This.TzName = m.CloserName

		* done, we found the nearest next saving time change
 		RETURN m.NextDate

	ENDFUNC

	* get the closest date from a list of RDATE
	HIDDEN FUNCTION GetClosestRDATE (RefTime AS Datetime, Start AS Datetime, RDates AS Collection) AS Datetime

		LOCAL RDate AS iCalPropRDATE
		LOCAL RDValue
		LOCAL RDatetime AS Datetime
		LOCAL RDClosest AS Datetime
		LOCAL LoopIndex AS Integer

		m.RDClosest = m.Start

		* if there is a list
		IF !ISNULL(m.RDates)
			* in each entry,
			FOR EACH m.RDate IN m.RDates
				*  get all date values
				FOR m.LoopIndex = 1 TO m.RDate.GetValueCount()
					* in case it's a duration, get the date part
					m.RDValue = m.RDate.GetValue(m.LoopIndex)
					IF VARTYPE(m.RDValue) == "O"
						m.RDatetime = m.RDValue.Datestart
					ELSE
						m.RDatetime = m.RDValue
					ENDIF
					* if it gets closer to the referred time, set it
					IF m.RDatetime > m.Start AND m.RDatetime < m.RefTime AND m.RDatetime > NVL(m.RDClosest, {})
						m.RDClosest = m.RDatetime
					ENDIF
				ENDFOR
			ENDFOR
		ENDIF

		* the closest in the list
		RETURN m.RDClosest

	ENDFUNC

	* get surrounding RDATE
	HIDDEN FUNCTION GetSurroundingRDATEs (RefTime AS Datetime, Start AS Datetime, RDates AS Collection) AS Collection

		LOCAL RDate AS iCalPropRDATE
		LOCAL RDValue
		LOCAL RDatetime AS Datetime
		LOCAL RDClosestPrevious AS Datetime
		LOCAL RDClosestNext AS Datetime
		LOCAL LoopIndex AS Integer
		LOCAL Result AS Collection

		m.RDClosestPrevious = m.Start
		m.RDClosestNext = .NULL.

		* if there is a list
		IF !ISNULL(m.RDates)
			* in each entry,
			FOR EACH m.RDate IN m.RDates
				*  get all date values
				FOR m.LoopIndex = 1 TO m.RDate.GetValueCount()
					* in case it's a duration, get the date part
					m.RDValue = m.RDate.GetValue(m.LoopIndex)
					IF VARTYPE(m.RDValue) == "O"
						m.RDatetime = m.RDValue.Datestart
					ELSE
						m.RDatetime = m.RDValue
					ENDIF
					* if it gets closer to the referred time below, set it
					IF m.RDatetime > m.Start AND m.RDatetime < m.RefTime AND m.RDatetime > m.RDClosestPrevious
						m.RDClosestPrevious = m.RDatetime
					ENDIF
					* now, the same check for a time above
					IF m.RDatetime > m.Start AND m.RDatetime >= m.RefTime AND m.RDatetime <= NVL(m.RDClosestNext, m.RDatetime)
						m.RDClosestNext = m.RDatetime
					ENDIF
				ENDFOR
			ENDFOR
		ENDIF

		IF !ISNULL(m.RDClosestNext)
			m.Result = CREATEOBJECT("Collection")
			m.Result.Add(m.RDClosestPrevious)
			m.Result.Add(m.RDClosestNext)
		ELSE
			m.Result = .NULL.
		ENDIF

		* the closest range in the list
		RETURN m.Result

	ENDFUNC

ENDDEFINE

DEFINE CLASS iCalCompVALARM AS _iCalComponent

	ICName = "VALARM"
	xICName = "valarm"

ENDDEFINE

DEFINE CLASS iCalCompSTANDARD AS _iCalComponent

	ICName = "STANDARD"
	xICName = "standard"

	LastChained = .NULL.

ENDDEFINE

DEFINE CLASS iCalCompDAYLIGHT AS _iCalComponent

	ICName = "DAYLIGHT"
	xICName = "daylight"

	LastChained = .NULL.

ENDDEFINE
