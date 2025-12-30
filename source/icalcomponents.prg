*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
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
	Ambiguous = .F.
	AmbiguityResolution = 2

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
							'<memberdata name="ambiguityresolution" type="property" display="AmbiguityResolution"/>' + ;
							'<memberdata name="ambiguous" type="property" display="Ambiguous"/>' + ;
							'<memberdata name="tzname" type="property" display="TzName"/>' + ;
							'<memberdata name="dst" type="method" display="DST"/>' + ;
							'<memberdata name="nextsavingtimechange" type="method" display="NextSavingTimeChange"/>' + ;
							'<memberdata name="popsavingtime" type="method" display="PopSavingTime"/>' + ;
							'<memberdata name="pushsavingtime" type="method" display="PushSavingTime"/>' + ;
							'<memberdata name="tolocaltime" type="method" display="ToLocalTime"/>' + ;
							'<memberdata name="toutc" type="method" display="ToUTC"/>' + ;
							'<memberdata name="utcoffset" type="method" display="UTCOffset"/>' + ;
					'</VFPData>'

	PROCEDURE AmbiguityResolution_assign (ARValue AS Integer)
		This.AmbiguityResolution = IIF(VARTYPE(m.ARValue) != "N" OR m.ARValue != 1, 2, 1)
	ENDPROC

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
	FUNCTION UTCOffset (RefDate AS Datetime, AmbiguityResolution AS Integer) AS Integer

		ASSERT VARTYPE(m.RefDate) $ "DT" AND (PCOUNT() == 1 OR VARTYPE(m.AmbiguityResolution) == "N")

		RETURN m.RefDate - IIF(PCOUNT() == 1, This.ToUTC(m.RefDate), This.ToUTC(m.RefDate, m.AmbiguityResolution))

	ENDFUNC
	
	* take a local time, and make it UTC
	FUNCTION ToUTC (LocalTime AS Datetime, AmbiguityResolution AS Integer)

		ASSERT VARTYPE(m.LocalTime) $ "DT" AND (PCOUNT() == 1 OR VARTYPE(m.AmbiguityResolution) == "N")

		RETURN IIF(PCOUNT() == 1, This._UTC(m.LocalTime, .T.), This._UTC(m.LocalTime, .T., m.AmbiguityResolution))

	ENDFUNC

	* take a UTC time, and make it local
	FUNCTION ToLocalTime (UTCTime AS Datetime)

		ASSERT VARTYPE(m.UTCTime) $ "DT"

		RETURN This._UTC(m.UTCTime, .F.)

	ENDFUNC

	HIDDEN FUNCTION _UTC (RefTime AS Datetime, ToUTC AS Logical, AmbiguityResolution AS Integer)

		SAFETHIS

		LOCAL OffsetTo AS Number, OffsetFrom AS Number, OffsetChange AS Number
		LOCAL ClosestOffsetTo AS Number, ClosestOffsetFrom AS Number
		LOCAL ClosestDate AS Datetime
		LOCAL RefDate AS Datetime, RefLocalTime AS Datetime
		LOCAL UntilDate AS Datetime, NextChange AS Datetime, ChangeTime AS Datetime
		LOCAL TzComp AS _iCalComponent
		LOCAL CompIndex AS Integer
		LOCAL RRule AS iCalPropRRULE
		LOCAL VRule AS iCalTypeRECUR
		LOCAL Period AS Logical
		LOCAL DST AS Logical
		LOCAL IsDST AS Logical

		* make sure we are working with datetimes
		IF VARTYPE(m.RefTime) == "D"
			RETURN DTOT(m.RefTime)
		ENDIF

		IF m.ToUTC
			m.RefLocalTime = m.RefTime
		ENDIF
		m.IsDST = IIF(PCOUNT() == 2, This.AmbiguityResolution, m.AmbiguityResolution) == 1

		m.NextChange = {/:}

		* the source time fits in the ST period set in a previous calculation?
		IF !This.Ambiguous AND !ISNULL(This.StartST) AND !ISNULL(This.EndST) ;
				AND ((m.ToUTC AND m.RefTime >= This.StartST AND m.RefTime < This.EndST) OR ;
					(!m.ToUTC AND m.RefTime + This.BiasST * 60 >= This.StartST AND m.RefTime + This.BiasST * 60 < This.EndST))

			* use the stored bias, no need to go through the calculation, again
			m.ClosestOffsetTo = This.BiasST

		ELSE

			This.Ambiguous = .F.

			* calculate the offset to UTC
			STORE 0 TO m.ClosestOffsetTo, m.OffsetTo
			STORE {^0001-01-01} TO m.ClosestDate, m.ClosestStandardDate
			This.TzName = ""
			STORE .NULL. TO This.StartST, This.EndST
			
			* mark if there is a DST definition for the reference date
			m.DST = This.DST(m.RefTime, !m.ToUTC)

			FOR m.CompIndex = 1 TO This.GetICComponentsCount()

				m.TzComp = This.GetICComponent(m.CompIndex)

				* look for all STANDARD, and also DAYLIGHT (if the reference date is covered by DST)
				IF m.TzComp.ICName == "STANDARD" OR (m.DST AND m.TzComp.ICName == "DAYLIGHT")

					* offset to UTC
					m.OffsetTo = NVL(m.TzComp.GetICPropertyValue("TZOFFSETTO"), 0)
					m.OffsetFrom = NVL(m.TzComp.GetICPropertyValue("TZOFFSETFROM"), 0)

					* get the local time to test against time in the timezone definition
					IF ! m.ToUTC
						m.RefLocalTime = m.RefTime + m.OffsetFrom * 60
					ENDIF

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
						IF ISNULL(m.UntilDate) OR YEAR(m.UntilDate) >= YEAR(m.RefLocalTime)
							m.RRule.CalculatePeriod(m.RefDate, m.RefLocalTime, .NULL., m.TzComp.GetICProperty("RDATE", -1), m.TzComp.GetICProperty("EXDATE", -1))
							IF !ISNULL(m.RRule.PreviousDate)
								m.RefDate = m.RRule.PreviousDate
								* check when next change on UTC offsset will occur
								IF ! ISNULL(m.RRule.NextDate)
									IF EMPTY(m.NextChange) OR m.NextChange > m.RRule.NextDate
										m.NextChange = m.RRule.NextDate
										m.OffsetChange = m.OffsetFrom - m.OffsetTo
									ENDIF
								ENDIF
								m.Period = .T.
							ELSE
								m.RefDate = .NULL.	&& this shouldn't happen... but something may be wrong with the rule, or the processor...
							ENDIF
						ELSE
							m.RefDate = .NULL.
						ENDIF
					* no RRULE? look for RDATEs
					ELSE
						m.RefDate = This.GetClosestRDATE(m.RefLocalTime, m.RefDate, m.TzComp.GetICProperty("RDATE", -1))
					ENDIF

					* we have a date, and it covers our time, and it's the closest to our time that we found so far?
					IF !ISNULL(m.RefDate) AND !EMPTY(m.RefDate) AND m.RefDate > m.ClosestDate

						IF m.RefDate <= m.RefLocalTime
							m.ClosestDate = m.RefDate
							m.ClosestOffsetTo = m.OffsetTo
							m.ClosestOffsetFrom = m.OffsetFrom
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
					ENDIF
	
					* if we used a period in the last calculation
					IF m.Period
						* if the next occurence is before our current ST ending date, then move our ST ending period to that point
						IF ISNULL(This.EndST) OR (m.RRule.NextDate > This.StartST AND m.RRule.NextDate < This.EndST)
							This.EndST = m.RRule.NextDate
						ELSE
							This.EndST = .NULL.
						ENDIF
					ENDIF
				ENDIF
			ENDFOR

		ENDIF
 
 		* we have a local time and want the UTC? subtract the bias
		IF m.ToUTC
			* check if it is an ambiguous time that occurs when UTC offset moves back
			IF ! EMPTY(m.NextChange)
				m.ChangeTime = m.NextChange
				m.NextChange = m.NextChange - m.OffsetChange * 60
				IF m.NextChange <= m.RefLocalTime
	 				This.Ambiguous = .T.
	 			ENDIF
			ENDIF
			DO CASE
			CASE ! This.Ambiguous
				m.CalcTime = m.RefLocalTime - m.ClosestOffsetTo * 60
			CASE m.RefLocalTime == m.NextChange
				IF m.IsDST
					m.CalcTime = m.RefLocalTime - m.ClosestOffsetTo * 60
				ELSE
					m.CalcTime = m.RefLocalTime - m.ClosestOffsetFrom * 60
				ENDIF
			CASE m.RefLocalTime == m.ChangeTime
				IF m.IsDST
					m.CalcTime = m.RefLocalTime - m.ClosestOffsetFrom * 60
				ELSE
					m.CalcTime = m.RefLocalTime - m.ClosestOffsetTo * 60
				ENDIF
			CASE m.IsDST
				m.CalcTime = m.RefLocalTime - m.ClosestOffsetTo * 60
			OTHERWISE
	 			m.CalcTime = m.RefLocalTime - m.ClosestOffsetFrom * 60
	 		ENDCASE
		* we have an UTC and want the local time? add the bias
		ELSE
			m.CalcTime = m.RefTime + m.ClosestOffsetTo * 60
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

	* check if there is a DST definition for a given date
	FUNCTION DST (RefTime AS Datetime, IsUTC AS Logical) AS Logical

		SAFETHIS

		ASSERT VARTYPE(m.RefTime) + VARTYPE(m.IsUTC) == "TL"

		LOCAL CompIndex AS Integer
		LOCAL TzComp AS iCalCompDAYLIGHT
		LOCAL DaylightStart AS Datetime
		LOCAL Datestart AS Datetime
		LOCAL SetDates AS Integer
		LOCAL AdditionalDate AS Datetime
		LOCAL RRule AS iCalPropRRULE
		LOCAL VRule AS iCalTypeRECUR
		LOCAL UntilDate AS Datetime
		LOCAL OffsetTo AS Integer

		m.DaylightStart = {:}

		* go through all DAYLIGHT components defined for the time zone
		FOR m.CompIndex = 1 TO This.GetICComponentsCount("DAYLIGHT")

			m.TzComp = This.GetICComponent("DAYLIGHT", m.CompIndex)
			* when it started
			m.DateStart = m.TzComp.GetICPropertyValue("DTSTART")
			* get any additional date that moved the start date forward
			FOR m.SetDates = 1 TO m.TzComp.GetICPropertiesCount("RDATE")
				m.AdditionalDate = m.TzComp.GetICPropertyValue("RDATE", m.SetDates)
				IF m.AdditionalDate > m.DateStart
					m.DateStart = m.AdditionalDate
				ENDIF
			ENDFOR
			* get an expiration date, if there is one
			m.RRule = m.TzComp.GetICProperty("RRULE")
			IF !ISNULL(m.RRule)
				m.VRule = m.RRule.GetValue()
				m.UntilDate = m.VRule.Until
			ELSE
				m.UntilDate = .NULL.
			ENDIF
			* adjust to the UTC offset
			m.OffsetTo = m.TzComp.GetICPropertyValue("TZOFFSETTO") * IIF(m.IsUTC, 0, 60)
			* if the start of the DST is before the reference date
			IF m.DateStart > m.DaylightStart AND m.DateStart <= (m.RefTime + m.OffsetTo)
				* and there is no expiration date
				IF m.RefTime <= NVL(m.UntilDate, m.RefTime)
					* mark this as the start
					m.DaylightStart = m.DateStart
				ELSE
					* otherwise, there is no DST covering the reference date (so far)
					m.DaylightStart = {:}
				ENDIF
			ENDIF
		ENDFOR

		RETURN !EMPTY(m.DaylightStart)

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
