*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
*!*	|                                                                    |
*!*	|    Source and docs: https://bitbucket.org/atlopes/ical4vfp         |
*!*	|                                                                    |
*!*	+--------------------------------------------------------------------+

*!*	iCalendar properties sub-classes

* install dependencies
DO LOCFILE("icalendar.prg")

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

#DEFINE	SAFETHIS		ASSERT !USED("This") AND TYPE("This") == "O"

* constants to support the Recurrence Rule processor

#DEFINE ICAL_ETERNITY		5000

#DEFINE MINUTE_IN_SECONDS	60
#DEFINE HOUR_IN_SECONDS		3600
#DEFINE DAY_IN_SECONDS		86400
#DEFINE WEEK_IN_SECONDS		604800

DEFINE CLASS iCalPropACTION AS _iCalProperty

	ICName = "ACTION"
	xICName = "action"

	Enumeration = "AUDIO,DISPLAY,EMAIL"

ENDDEFINE

DEFINE CLASS iCalPropATTACH AS _iCalProperty

	ICName = "ATTACH"
	xICName = "attach"

	Value_DataType = "URI"
	Value_AlternativeDataTypes = "BINARY"

ENDDEFINE

DEFINE CLASS iCalPropATTENDEE AS _iCalProperty

	ICName = "ATTENDEE"
	xICName = "attendee"

	Value_DataType = "CAL-ADDRESS"

ENDDEFINE

DEFINE CLASS iCalPropCALSCALE AS _iCalProperty

	ICName = "CALSCALE"
	xICName = "calscale"

	DefaultValue = "GREGORIAN"
	ReadOnly = .T.

ENDDEFINE

DEFINE CLASS iCalPropCATEGORIES AS _iCalProperty

	ICName = "CATEGORIES"
	xICName = "categories"

	Value_IsList = .T.

ENDDEFINE

DEFINE CLASS iCalPropCLASS AS _iCalProperty

	ICName = "CLASS"
	xICName = "class"

	Enumeration = "PUBLIC,PRIVATE,CONFIDENTIAL"

ENDDEFINE

DEFINE CLASS iCalPropCOMMENT AS _iCalProperty

	ICName = "COMMENT"
	xICName = "comment"

ENDDEFINE

DEFINE CLASS iCalPropCOMPLETED AS _iCalProperty

	ICName = "COMPLETED"
	xICName = "completed"

	Value_DataType = "DATE-TIME"
	Value_IsUTC = .T.

ENDDEFINE

DEFINE CLASS iCalPropCONTACT AS _iCalProperty

	ICName = "CONTACT"
	xICName = "contact"

ENDDEFINE

DEFINE CLASS iCalPropCREATED AS _iCalProperty

	ICName = "CREATED"
	xICName = "created"

	Value_DataType = "DATE-TIME"
	Value_IsUTC = .T.

ENDDEFINE

DEFINE CLASS iCalPropDESCRIPTION AS _iCalProperty

	ICName = "DESCRIPTION"
	xICName = "description"

ENDDEFINE

DEFINE CLASS iCalPropDTEND AS _iCalProperty

	ICName = "DTEND"
	xICName = "dtend"

	Value_DataType = "DATE-TIME"
	Value_AlternativeDataTypes = "DATE"

ENDDEFINE

DEFINE CLASS iCalPropDTSTAMP AS _iCalProperty

	ICName = "DTSTAMP"
	xICName = "dtstamp"

	Value_DataType = "DATE-TIME"
	Value_IsUTC = .T.

ENDDEFINE

DEFINE CLASS iCalPropDTSTART AS _iCalProperty

	ICName = "DTSTART"
	xICName = "dtstart"

	Value_DataType = "DATE-TIME"
	Value_AlternativeDataTypes = "DATE"

ENDDEFINE

DEFINE CLASS iCalPropDUE AS _iCalProperty

	ICName = "DUE"
	xICName = "due"

	Value_DataType = "DATE-TIME"
	Value_AlternativeDataTypes = "DATE"

ENDDEFINE

DEFINE CLASS iCalPropDURATION AS _iCalProperty

	ICName = "DURATION"
	xICName = "duration"

	Value_DataType = "DURATION"

ENDDEFINE

DEFINE CLASS iCalPropEXDATE AS _iCalProperty

	ICName = "EXDATE"
	xICName = "exdate"

	Value_DataType = "DATE-TIME"
	Value_AlternativeDataTypes = "DATE"
	Value_IsList = .T.

ENDDEFINE

DEFINE CLASS iCalPropFREEBUSY AS _iCalProperty

	ICName = "FREEBUSY"
	xICName = "freebusy"

	Value_DataType = "PERIOD"

ENDDEFINE

DEFINE CLASS iCalPropGEO AS _iCalProperty

	ICName = "GEO"
	xICName = "geo"

	Value_DataType = "FLOAT"
	Value_IsComposite = .T.

ENDDEFINE

DEFINE CLASS iCalPropLAST_MODIFIED AS _iCalProperty

	ICName = "LAST-MODIFIED"
	xICName = "last-modified"

	Value_DataType = "DATE-TIME"
	Value_IsUTC = .T.

ENDDEFINE

DEFINE CLASS iCalPropLOCATION AS _iCalProperty

	ICName = "LOCATION"
	xICName = "location"

ENDDEFINE

DEFINE CLASS iCalPropMETHOD AS _iCalProperty

	ICName = "METHOD"
	xICName = "method"

ENDDEFINE

DEFINE CLASS iCalPropORGANIZER AS _iCalProperty

	ICName = "ORGANIZER"
	xICName = "organizer"

	Value_DataType = "CAL-ADDRESS"

ENDDEFINE

DEFINE CLASS iCalPropPERCENTCOMPLETE AS _iCalProperty

	ICName = "PERCENT-COMPLETE"
	xICName = "percent-complete"

	Value_DataType = "INTEGER"

ENDDEFINE

DEFINE CLASS iCalPropPRIORITY AS _iCalProperty

	ICName = "PRIORITY"
	xICName = "priority"

	Value = 0
	Value_DataType = "INTEGER"

ENDDEFINE

DEFINE CLASS iCalPropPRODID AS _iCalProperty

	ICName = "PRODID"
	xICName = "prodid"

	DefaultValue = "//BKM/iCal4VFP/2018"

ENDDEFINE

DEFINE CLASS iCalPropRDATE AS _iCalProperty

	ICName = "RDATE"
	xICName = "rdate"

	Value_DataType = "DATE-TIME"
	Value_AlternativeDataTypes = "DATE,PERIOD"
	Value_IsList = .T.

ENDDEFINE

DEFINE CLASS iCalPropRECURRENCE_ID AS _iCalProperty

	ICName = "RECURRENCE-ID"
	xICName = "recurrence-id"

	Value_DataType = "DATE-TIME"
	Value_AlternativeDataTypes = "DATE"

ENDDEFINE

DEFINE CLASS iCalPropRELATED_TO AS _iCalProperty

	ICName = "RELATED-TO"
	xICName = "related-to"

ENDDEFINE

DEFINE CLASS iCalPropREPEAT AS _iCalProperty

	ICName = "REPEAT"
	xICName = "repeat"

	DefaultValue = 0
	Value_DataType = "INTEGER"

ENDDEFINE

DEFINE CLASS iCalPropREQUEST_STATUS AS _iCalProperty

	ICName = "REQUEST-STATUS"
	xICName = "request-status"

	Value_IsComposite = .T.

ENDDEFINE

DEFINE CLASS iCalPropRESOURCES AS _iCalProperty

	ICName = "RESOURCES"
	xICName = "resources"

	Value_IsList = .T.

ENDDEFINE

DEFINE CLASS iCalPropRRULE AS _iCalProperty

	ICName = "RRULE"
	xICName = "rrule"

	Value_DataType = "RECUR"

	NextDate = .NULL.
	NextTzName = ""
	NextUTCDate = .NULL.
	PreviousDate = .NULL.
	PreviousTzName = ""
	PreviousUTCDate = .NULL.

	HIDDEN ByMonth(1)
	HIDDEN ByWeekNo(1)
	HIDDEN ByYearDay(1)
	HIDDEN ByMonthDay(1)
	HIDDEN ByDay(1)
	HIDDEN ByHour(1)
	HIDDEN ByMinute(1)
	HIDDEN BySecond(1)
	HIDDEN BySetPos(1)
	DIMENSION ByMonth(1)
	DIMENSION ByWeekNo(1)
	DIMENSION ByYearDay(1)
	DIMENSION ByMonthDay(1)
	DIMENSION ByDay(1)
	DIMENSION ByHour(1)
	DIMENSION ByMinute(1)
	DIMENSION BySecond(1)
	DIMENSION BySetPos(1)

	_MemberData = 	'<VFPData>' + ;
							'<memberdata name="nextdate" type="property" display="NextDate" />' + ;
							'<memberdata name="nextutcdate" type="property" display="NextUTCDate" />' + ;
							'<memberdata name="nexttzname" type="property" display="NextTzName" />' + ;
							'<memberdata name="previousdate" type="property" display="PreviousDate" />' + ;
							'<memberdata name="previousutcdate" type="property" display="PreviousUTCDate" />' + ;
							'<memberdata name="previoustzname" type="property" display="PreviousTzName" />' + ;
							'<memberdata name="calculateall" type="method" display="CalculateAll" />' + ;
							'<memberdata name="calculatenext" type="method" display="CalculateNext" />' + ;
							'<memberdata name="calculateperiod" type="method" display="CalculatePeriod" />' + ;
							'<memberdata name="calculateprevious" type="method" display="CalculatePrevious" />' + ;
						'</VFPData>'

	* Calculate all events determined by a recurrent rule (return a cursor name)
	FUNCTION CalculateAll (Start AS Datetime, Finish AS Datetime, TZ AS iCalCompTIMEZONE, AddDates AS iCalPropRDATE, Exceptions AS iCalPropEXDATE) AS String

		IF PCOUNT() = 2
			m.TZ = .NULL.
		ENDIF
		
		IF PCOUNT() < 4
			STORE .NULL. TO m.AddDates, m.Exceptions
		ENDIF

		RETURN This.Calculator(m.Start, m.Finish, m.TZ, m.AddDates, m.Exceptions, .F.)

	ENDFUNC

	* calculate the previous date of a recurrent event
	FUNCTION CalculatePrevious (Start AS Datetime, Finish AS Datetime, TZ AS iCalCompTIMEZONE, AddDates AS iCalPropRDATE, Exceptions AS iCalPropEXDATE)

		SAFETHIS

		LOCAL CalcCursor AS String
		LOCAL ARRAY PreviousDate(1)

		IF PCOUNT() = 2
			m.TZ = .NULL.
		ENDIF
		
		IF PCOUNT() < 4
			STORE .NULL. TO m.AddDates, m.Exceptions
		ENDIF

		IF !ISNULL(m.TZ)
			m.TZ.PushSavingTime()
		ENDIF

		This.PreviousDate = .NULL.

		m.CalcCursor = This.Calculator(m.Start, m.Finish, m.TZ, m.AddDates, m.Exceptions, .F.)
		SELECT LocalTime, TzName FROM (m.CalcCursor) WHERE LocalTime = (SELECT MAX(LocalTime) FROM (m.CalcCursor) WHERE LocalTime <= m.Finish) INTO ARRAY PreviousDate

		IF _Tally = 1
			This.PreviousDate = m.PreviousDate(1, 1)
			This.PreviousTzName = NVL(m.PreviousDate(1, 2), "")
		ENDIF

		USE IN (m.CalcCursor)

		IF !ISNULL(m.TZ)
			m.TZ.PopSavingTime()
		ENDIF

		RETURN This.PreviousDate

	ENDFUNC

	* calculate the limit dates of a recurrent event
	FUNCTION CalculatePeriod (Start AS Datetime, Finish AS Datetime, TZ AS iCalCompTIMEZONE, AddDates AS iCalPropRDATE, Exceptions AS iCalPropEXDATE)

		SAFETHIS

		LOCAL CalcCursor AS String
		LOCAL ARRAY LimitDates(1)

		IF PCOUNT() = 2
			m.TZ = .NULL.
		ENDIF
		
		IF PCOUNT() < 4
			STORE .NULL. TO m.AddDates, m.Exceptions
		ENDIF

		IF !ISNULL(m.TZ)
			m.TZ.PushSavingTime()
		ENDIF

		This.PreviousDate = .NULL.
		This.NextDate = .NULL.

		m.CalcCursor = This.Calculator(m.Start, m.Finish, m.TZ, m.AddDates, m.Exceptions, .T.)
		SELECT LocalTime, TzName FROM (m.CalcCursor) WHERE LocalTime = (SELECT MAX(LocalTime) FROM (m.CalcCursor) WHERE LocalTime <= m.Finish) ;
		UNION ALL ;
		SELECT LocalTime, TzName FROM (m.CalcCursor) WHERE LocalTime = (SELECT MIN(LocalTime) FROM (m.CalcCursor) WHERE LocalTime > m.Finish) ;
		ORDER BY 1 ;
		INTO ARRAY LimitDates

		IF _Tally = 2
			This.PreviousDate = m.LimitDates(1, 1)
			This.PreviousTzName = NVL(m.LimitDates(1, 2), "")
			This.NextDate = m.LimitDates(2, 1)
			This.NextTzName = NVL(m.LimitDates(2, 2), "")
		ENDIF

		IF !ISNULL(m.TZ)
			m.TZ.PopSavingTime()
		ENDIF

		USE IN (m.CalcCursor)

	ENDFUNC

	* calculate the next date of a recurrent event
	FUNCTION CalculateNext (Start AS Datetime, Finish AS Datetime, TZ AS iCalCompTIMEZONE, AddDates AS iCalPropRDATE, Exceptions AS iCalPropEXDATE) AS Datetime

		SAFETHIS

		LOCAL CalcCursor AS String
		LOCAL ARRAY NextDate(1)

		IF PCOUNT() = 2
			m.TZ = .NULL.
		ENDIF
		
		IF PCOUNT() < 4
			STORE .NULL. TO m.AddDates, m.Exceptions
		ENDIF

		IF !ISNULL(m.TZ)
			m.TZ.PushSavingTime()
		ENDIF

		This.NextDate = .NULL.

		m.CalcCursor = This.Calculator(m.Start, m.Finish, m.TZ, m.AddDates, m.Exceptions, .T.)
		SELECT LocalTime, TzName FROM (m.CalcCursor) WHERE LocalTime = (SELECT MIN(LocalTime) FROM (m.CalcCursor) WHERE LocalTime > m.Finish) INTO ARRAY NextDate

		IF _Tally = 1
			This.NextDate = m.NextDate(1, 1)
			This.NextTzName = NVL(m.NextDate(1, 2), "")
		ENDIF

		USE IN (m.CalcCursor)

		IF !ISNULL(m.TZ)
			m.TZ.PopSavingTime()
		ENDIF

		RETURN This.NextDate

	ENDFUNC

	* the RRULE processor
	HIDDEN FUNCTION Calculator (Start AS Datetime, Finish AS Datetime, TZ AS iCalCompVTIMEZONE, AddDates AS Collection, ExceptDates AS Collection, CalcNext AS Logical) AS String

		SAFETHIS

		ASSERT VARTYPE(m.Start) $ "DT" AND VARTYPE(m.Finish) $ "DT" AND VARTYPE(m.TZ) $ "OX" AND VARTYPE(m.AddDates) $ "OX" AND VARTYPE(m.ExceptDates) $ "OX" AND VARTYPE(m.CalcNext) == "L"

		LOCAL Rule AS iCalTypeRECUR
		LOCAL ReCursor AS String
		LOCAL TempDates AS Collection
		LOCAL ReCount AS Integer
		LOCAL Current AS Datetime
		LOCAL Until AS Datetime
		LOCAL Interval AS Integer
		LOCAL STUsage AS Logical
		LOCAL NextSavingTime AS Datetime
		LOCAL TzName AS String
		LOCAL DatePart AS Date
		LOCAL DatetimePart AS Datetime
		LOCAL AltDatetimePart AS Datetime
		LOCAL YearPart AS Integer
		LOCAL MonthPart AS Integer
		LOCAL DayPart AS Integer
		LOCAL HourPart AS Integer
		LOCAL MinutePart AS Integer
		LOCAL SecondPart AS Integer
		LOCAL StepIndex AS Integer
		LOCAL NextInterval AS Logical
		LOCAL LoopIndex AS Integer
		LOCAL AddDate AS iCalPropRDATE
		LOCAL ExceptDate AS iCalPropEXDATE
		LOCAL REXDValue

		* get the rule definitions in a comfortable manner
		m.Rule = This.GetValue()
		This.RuleCollectionToArray("ByMonth", m.Rule.ByMonth)
		This.RuleCollectionToArray("ByWeekNo", m.Rule.ByWeekNo)
		This.RuleCollectionToArray("ByYearDay", m.Rule.ByYearDay)
		This.RuleCollectionToArray("ByMonthDay", m.Rule.ByMonthDay)
		This.RuleCollectionToArray("ByDay", m.Rule.ByDay)
		This.RuleCollectionToArray("ByHour", m.Rule.ByHour)
		This.RuleCollectionToArray("ByMinute", m.Rule.ByMinute)
		This.RuleCollectionToArray("BySecond", m.Rule.BySecond)
		This.RuleCollectionToArray("BySetPos", m.Rule.BySetPos)
		m.Interval = NVL(m.Rule.Interval, 1)
		* set the Until date to local time, if necessary
		IF !ISNULL(m.Rule.Until)
			IF m.Rule.IsUTC AND !ISNULL(m.TZ)
				m.TZ.PushSavingTime()
				m.Until = m.TZ.ToLocalTime(m.Rule.Until)
				m.TZ.PopSavingTime()
			ELSE
				m.Until = m.Rule.Until
			ENDIF
		ELSE
			m.Until = .NULL.
		ENDIF

		* find an available name for the cursor
		DO WHILE EMPTY(m.ReCursor) OR USED(m.ReCursor)
			m.ReCursor = "RRule" + SYS(2015)
		ENDDO
		CREATE CURSOR (m.ReCursor) (LocalTime Datetime, TzName Varchar(32) NULL)
		INDEX ON LocalTime TAG LocalTime

		* the dates arranged for each interval
		m.TempDates = CREATEOBJECT("Collection")

		* a date to control the interval
		m.ReCount = 1
		* a safe date will be set as the starting point
		m.Current = This.GetCurrentDate(m.Rule, m.Start)

		* how is the date related to Timezones
		* for the time part interval: check on saving time changes affecting the interval
		m.STUsage = .F.
		IF !ISNULL(m.TZ)
			m.NextSavingTime = m.TZ.NextSavingTimeChange(m.Current)
			m.TzName = m.TZ.TzName
			m.STUsage = m.Rule.Freq == "HOURLY" OR m.Rule.Freq == "MINUTELY" OR m.Rule.Freq == "SECONDELY"
		ELSE
			m.NextSavingTime = .NULL.
			m.TzName = .NULL.
		ENDIF

		m.DatetimePart = m.Current

		m.YearPart = YEAR(m.Current)
		m.MonthPart = MONTH(m.Current)
		m.DayPart = DAY(m.Current)
		
		m.HourPart = HOUR(m.Current)
		m.MinutePart = MINUTE(m.Current)
		m.SecondPart = SEC(m.Current)

		* an extra interval may be required when the next date was requested
		m.NextInterval = .T.

		* while we didn't reach the finish date, or the count of dates, or the final date in the RRULE
		DO WHILE (m.Current <= m.Finish OR (m.CalcNext AND m.NextInterval)) ;
				AND m.ReCount <= NVL(m.Rule.Count, ICAL_ETERNITY) ;
				AND (ISNULL(m.Until) OR m.Current <= m.Until)

			* the start date will always be part of the date series
			IF RECCOUNT(m.ReCursor) = 0
				m.TempDates.Add(m.Start, TTOC(m.Start, 1))
			ENDIF

			* apply the By part rules, starting from the ByMonth part
			IF !ISNULL(m.DatetimePart) AND !EMPTY(m.DatetimePart)
				This.ApplyByMonth(m.Rule, m.Start, m.DatetimePart, m.TempDates)
				* and also the case where there are no By part rules
				This.ApplyNoBy(m.Rule, m.Start, m.DatetimePart, m.TempDates)
			ENDIF

			* get the dates in order from the collection of dates determined for the interval
			m.TempDates.KeySort = 2
			m.StepIndex = 0
			FOR EACH m.DatetimePart IN m.TempDates
				* if a positional By part is in effect, exclude all the dates that are not part of the positional collection
				* otherwise, include all (the others)
				m.StepIndex = m.StepIndex + 1
				IF ISNULL(m.Rule.BySetPos) ;
						OR ASCAN(This.BySetPos, m.StepIndex) != 0 ;
						OR ASCAN(This.BySetPos, -((m.TempDates.Count + 1) - m.StepIndex)) != 0
					IF m.DatetimePart >= m.Start
						IF !ISNULL(m.NextSavingTime) AND m.DatetimePart > m.NextSavingTime
							m.TZ.PushSavingTime()
							m.NextSavingTime = m.TZ.NextSavingTimeChange(m.DatetimePart, .T.)
							m.TzName = m.TZ.TzName
							m.TZ.PopSavingTime()
						ENDIF
						INSERT INTO (m.ReCursor) VALUES (m.DatetimePart, m.TzName)
						* we're processing and found the next date: this will be the last interval
						IF m.Current > m.Finish
							m.NextInterval = .F.
						ENDIF
					ENDIF
				ENDIF
			ENDFOR
			m.TempDates.KeySort = 0
			m.TempDates.Remove(-1)
			m.ReCount = RECCOUNT(m.ReCursor)

			* calculate the next interval, depending on the frequency (a valid date is always safely generated)
			DO CASE
			CASE m.Rule.Freq == "YEARLY"
				m.YearPart = m.YearPart + m.Interval
			CASE m.Rule.Freq == "MONTHLY"
				m.MonthPart = m.MonthPart + m.Interval
				IF m.MonthPart > 12
					m.YearPart = m.YearPart + INT((m.MonthPart - 1) / 12)
					m.MonthPart = (m.MonthPart - 1) % 12 + 1
				ENDIF
			CASE m.Rule.Freq == "WEEKLY"
				m.DatePart = TTOD(m.Current) + 7 * m.Interval
				m.YearPart = YEAR(m.DatePart)
				m.MonthPart = MONTH(m.DatePart)
				m.DayPart = DAY(m.DatePart)
			CASE m.Rule.Freq == "DAILY"
				m.DatePart = TTOD(m.Current) + m.Interval
				m.YearPart = YEAR(m.DatePart)
				m.MonthPart = MONTH(m.DatePart)
				m.DayPart = DAY(m.DatePart)
			OTHERWISE
				DO CASE
				CASE m.Rule.Freq == "HOURLY"
					m.DatetimePart = m.Current + HOUR_IN_SECONDS * m.Interval
				CASE m.Rule.Freq == "MINUTELY"
					m.DatetimePart = m.Current + MINUTE_IN_SECONDS * m.Interval
				CASE m.Rule.Freq == "SECONDELY"
					m.DatetimePart = m.Current + m.Interval
				OTHERWISE
					EXIT
				ENDCASE
				IF m.STUsage AND m.DatetimePart > m.NextSavingTime
					m.TZ.PushSavingTime()
					m.NextSavingTime = m.TZ.NextSavingTimeChange(m.NextSavingTime, .T.)
					m.TZ.PopSavingTime()
					m.DatetimePart = m.TZ.ToUTC(m.DatetimePart)
					m.TzName = m.TZ.TzName
				ENDIF
				m.YearPart = YEAR(m.DatetimePart)
				m.MonthPart = MONTH(m.DatetimePart)
				m.DayPart = DAY(m.DatetimePart)
				m.HourPart = HOUR(m.DatetimePart)
				IF !(m.Rule.Freq == "HOURLY")
					m.MinutePart = MINUTE(m.DatetimePart)
					IF !(m.Rule.Freq == "MINUTELY")
						m.SecondPart = SEC(m.DatetimePart)
					ENDIF
				ENDIF
			ENDCASE

			* the reference date for the next interval (unless we found eternity)
			IF m.YearPart <= 9999
				m.DatetimePart = DATETIME(m.YearPart, m.MonthPart, m.DayPart, m.HourPart, m.MinutePart, m.SecondPart)
				m.Current = m.DatetimePart
			ELSE
				EXIT
			ENDIF

		ENDDO

		* remove duplicates
		SELECT LocalTime, TzName ;
			FROM (m.ReCursor) ;
			GROUP BY LocalTime, TzName ;
			INTO CURSOR (m.ReCursor) READWRITE

		* a final selection may be required if a count was in effect
		IF !ISNULL(m.Rule.Count)
			SELECT TOP (m.Rule.Count) * ;
				FROM (m.ReCursor) ;
				ORDER BY LocalTime ;
				INTO CURSOR (m.ReCursor) READWRITE
		ENDIF
		INDEX ON LocalTime TAG LocalTime

		* add the RDATEs, if any
		IF !ISNULL(m.AddDates)
			FOR EACH m.AddDate IN m.AddDates
				FOR m.LoopIndex = 1 TO m.AddDate.GetValueCount()
					m.REXDValue = m.AddDate.GetValue(m.LoopIndex)
					IF VARTYPE(m.REXDValue) == "O"
						m.REXDValue = m.REXDValue.DateStart
					ENDIF
					LOCATE FOR LocalTime = m.REXDValue
					IF !FOUND()
						INSERT INTO (m.ReCursor) VALUES (m.REXDValue, .NULL.)
					ENDIF
				ENDFOR
			ENDFOR
		ENDIF
		* exclude the EXDATEs, if any
		IF !ISNULL(m.ExceptDates)
			FOR EACH m.ExceptDate IN m.ExceptDates
				FOR m.LoopIndex = 1 TO m.ExceptDate.GetValueCount()
					m.REXDValue = m.ExceptDate.GetValue(m.LoopIndex)
					DELETE FROM (m.ReCursor) WHERE LocalTime = m.REXDValue
				ENDFOR
			ENDFOR
			SELECT * FROM (m.ReCursor) WHERE !DELETED() INTO CURSOR (m.ReCursor) READWRITE
			INDEX ON LocalTime TAG LocalTime
		ENDIF

		RETURN m.ReCursor

	ENDFUNC

	* all By rule parts either expand or limit the current reference interval date
	* the ByDay is a special case because it must handle mixed limits/expansions, depending on the frequency and other By rule parts

	* refer to RFC 5545 for documentation

	HIDDEN FUNCTION ApplyByMonth (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .F.

		* rule set?
		IF !ISNULL(m.Rule.ByMonth)
			* expand
			IF m.Rule.Freq == "YEARLY"
				FOR m.Entry = 1 TO ALEN(This.ByMonth)
					m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, This.ByMonth(m.Entry), "RVSSSS")
					IF !EMPTY(m.CalcDate) AND This.ApplyByWeekNo(m.Rule, m.Start, m.CalcDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
							m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
						ENDIF
						m.Accepted = .T.
					ENDIF
				ENDFOR
			ELSE
			* limit
				m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, 0, "RRRSSS")
				IF ASCAN(This.ByMonth, MONTH(m.CalcDate)) != 0
					IF This.ApplyByWeekNo(m.Rule, m.Start, m.CalcDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
							m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
						ENDIF
						m.Accepted = .T.
					ENDIF
				ENDIF
			ENDIF
		ELSE
			* rule not set? pass to next rule
			m.Accepted = This.ApplyByWeekNo(m.Rule, m.Start, m.RefDate, m.Dates)
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyByWeekNo (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL DayInWeek AS Integer
		LOCAL FDWeek AS Integer
		LOCAL FDWYear AS Integer
		LOCAL FirstFourDaysWeek AS Datetime
		LOCAL FirstDay AS Date
		LOCAL NumberOfWeeks AS Integer
		LOCAL WeekNo AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .F.

		* rule set?
		IF !ISNULL(m.Rule.ByWeekNo) AND m.Rule.Freq == "YEARLY"

			m.FDWeek = VAL(STREXTRACT("SU:1:MO:2:TU:3:WE:4:TH:5:FR:6:SA:7:", NVL(m.Rule.WkSt, "MO") + ":", ":"))
			m.FirstDay = DATE(YEAR(m.RefDate), 1, 1)
			m.FDWYear = DOW(m.FirstDay, m.FDWeek) + 1
			m.CalcDate = m.FirstDay + (m.FDWeek - m.FDWYear)
			IF m.FirstDay - m.CalcDate >= 4
				m.CalcDate = m.CalcDate + 7
			ENDIF
			IF DATE(YEAR(m.RefDate) + 1, 1, 1) - (m.CalcDate + 52 * 7) >= 4
				m.NumberOfWeeks = 53
			ELSE
				m.NumberOfWeeks = 52
			ENDIF

			m.FirstFourDaysWeek = DATETIME(YEAR(m.CalcDate), MONTH(m.CalcDate), DAY(m.CalcDate), HOUR(m.RefDate), MINUTE(m.RefDate), SEC(m.RefDate))

			FOR m.Entry = 1 TO ALEN(This.ByWeekNo)
				m.WeekNo = This.ByWeekNo(m.Entry)
				IF ABS(m.WeekNo) <= m.NumberOfWeeks
					IF m.WeekNo < 0
						m.WeekNo = (m.NumberOfWeeks + 1) + m.WeekNo
					ENDIF
					m.CalcDate = m.FirstFourDaysWeek + (m.WeekNo - 1) * WEEK_IN_SECONDS
					FOR m.DayInWeek = 0 TO 6
						IF This.ApplyByYearDay(m.Rule, m.Start, m.CalcDate, m.Dates)
							IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
								m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
							ENDIF
							m.Accepted = .T.
						ENDIF
						m.CalcDate = m.CalcDate + DAY_IN_SECONDS
					ENDFOR
				ENDIF
			ENDFOR
		ELSE
			* rule not set? pass to next rule
			m.Accepted = This.ApplyByYearDay(m.Rule, m.Start, m.RefDate, m.Dates)
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyByYearDay (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL PosDay AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .F.

		* rule set?
		IF !ISNULL(m.Rule.ByYearDay)
			
			FOR m.Entry = 1 TO ALEN(This.ByYearDay)
				m.PosDay = This.ByYearDay(m.Entry)
				IF m.PosDay > 0
					m.CalcDate = This.GetStartBasedDate(m.Start, DATE(YEAR(m.RefDate), 1, 1), 0, "RRRSSS") + DAY_IN_SECONDS * (m.PosDay - 1)
				ELSE
					m.CalcDate = This.GetStartBasedDate(m.Start, GOMONTH(DATE(YEAR(m.RefDate), 1, 1), 12), 0, "RRRSSS") + DAY_IN_SECONDS * m.PosDay
				ENDIF
				IF m.Rule.Freq == "YEARLY"
					* expand
					IF YEAR(m.CalcDate) = YEAR(m.RefDate)
						IF This.ApplyByMonthDay(m.Rule, m.Start, m.CalcDate, m.Dates)
							IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
								m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
							ENDIF
							m.Accepted = .T.
						ENDIF
					ENDIF
				ELSE
					* limit
					IF m.CalcDate = m.RefDate
						IF This.ApplyByMonthDay(m.Rule, m.Start, m.CalcDate, m.Dates)
							IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
								m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
							ENDIF
							m.Accepted = .T.
						ENDIF
					ENDIF
				ENDIF
			ENDFOR
		ELSE
			* rule not set? pass to next rule
			m.Accepted = This.ApplyByMonthDay(m.Rule, m.Start, m.RefDate, m.Dates)
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyByMonthDay (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL PosDay AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .F.

		* rule set?
		IF !ISNULL(m.Rule.ByMonthDay)
			
			FOR m.Entry = 1 TO ALEN(This.ByMonthDay)
				TRY
					m.PosDay = This.ByMonthDay(m.Entry)
					IF m.PosDay > 0
						m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, m.PosDay, "RRVSSS")
					ELSE
						m.CalcDate = This.GetStartBasedDate(m.Start, GOMONTH(DATE(YEAR(m.RefDate), MONTH(m.RefDate), 1), 1), 1, "RRVSSS")
						m.CalcDate = m.CalcDate + 86400 * m.PosDay
					ENDIF
					* expand
					IF m.Rule.Freq == "YEARLY" OR m.Rule.Freq == "MONTHLY"
						IF MONTH(m.CalcDate) = MONTH(m.RefDate)
							IF This.ApplyByDay(m.Rule, m.Start, m.CalcDate, m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
								m.Accepted = .T.
							ENDIF
						ENDIF
					ELSE
					* limit
						IF m.CalcDate = m.RefDate
							IF This.ApplyByDay(m.Rule, m.Start, m.CalcDate, m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
								m.Accepted = .T.
							ENDIF
						ENDIF
					ENDIF
				CATCH
				ENDTRY
			ENDFOR
		ELSE
			* rule not set? pass to next rule
			m.Accepted = This.ApplyByDay(m.Rule, m.Start, m.RefDate, m.Dates)
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyByDay (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL TempDate AS Datetime
		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL FDWeek AS Integer
		LOCAL DWeek AS Integer
		LOCAL PosWeek AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .F.

		* rule set?
		IF !ISNULL(m.Rule.ByDay)

			DO CASE

			CASE m.Rule.Freq == "YEARLY" AND (!ISNULL(m.Rule.ByYearDay) OR !ISNULL(m.Rule.ByMonthDay))
				* limit
				m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, 0, "RRRSSS")
				IF ASCAN(This.ByDay, SUBSTR(":SUMOTUWETHFRSA", DOW(m.CalcDate) * 2, 2)) != 0
					IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
							m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
						ENDIF
					ENDIF
				ENDIF

			CASE m.Rule.Freq == "YEARLY" AND !ISNULL(m.Rule.ByMonth)
				* expand
				FOR m.Entry = 1 TO ALEN(This.ByDay)
					m.CalcDate = This.GetStartBasedDate(m.Start, DATE(YEAR(m.RefDate), MONTH(m.RefDate), 1), 0, "RRRSSS")
					m.PosWeek = VAL(This.ByDay(m.Entry))
					m.DWeek = VAL(STREXTRACT("SU:1:MO:2:TU:3:WE:4:TH:5:FR:6:SA:7:", IIF(m.PosWeek != 0, RIGHT(This.ByDay(m.Entry), 2), This.ByDay(m.Entry)) + ":", ":"))
					m.CalcDate = m.CalcDate + DAY_IN_SECONDS * IIF(m.DWeek >= DOW(m.CalcDate), m.DWeek - DOW(m.CalcDate), m.DWeek + 7 - DOW(m.CalcDate))
					IF m.PosWeek = 0
						DO WHILE MONTH(m.CalcDate) = MONTH(m.RefDate)
							IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
							ENDIF
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS
						ENDDO
					ELSE
						IF m.PosWeek > 0
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS * (m.PosWeek - 1)
						ELSE
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS * 4
							IF MONTH(m.CalcDate) != MONTH(m.RefDate)
								m.CalcDate = m.CalcDate - WEEK_IN_SECONDS
							ENDIF
							m.CalcDate = m.CalcDate - WEEK_IN_SECONDS * ABS(m.PosWeek + 1)
						ENDIF
						IF MONTH(m.CalcDate) = MONTH(m.RefDate)
							IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate,m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
							ENDIF
						ENDIF
					ENDIF
				ENDFOR

			CASE m.Rule.Freq == "YEARLY" AND ISNULL(m.Rule.ByWeekNo) AND ISNULL(m.Rule.ByMonth)
				* expand
				m.FDWeek = VAL(STREXTRACT("SU:1:MO:2:TU:3:WE:4:TH:5:FR:6:SA:7:", NVL(m.Rule.WkSt, "MO") + ":", ":"))
				FOR m.Entry = 1 TO ALEN(This.ByDay)
					m.CalcDate = This.GetStartBasedDate(m.Start, DATE(YEAR(m.RefDate), 1, 1), 0, "RRRSSS")
					m.PosWeek = VAL(This.ByDay(m.Entry))
					m.DWeek = VAL(STREXTRACT("SU:1:MO:2:TU:3:WE:4:TH:5:FR:6:SA:7:", IIF(m.PosWeek != 0, RIGHT(This.ByDay(m.Entry), 2), This.ByDay(m.Entry)) + ":", ":"))
					m.CalcDate = m.CalcDate + DAY_IN_SECONDS * IIF(m.DWeek >= DOW(m.CalcDate), m.DWeek - DOW(m.CalcDate), m.DWeek + 7 - DOW(m.CalcDate))
					IF m.PosWeek = 0
						DO WHILE YEAR(m.CalcDate) = YEAR(m.RefDate)
							IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
								m.Accepted = .T.
							ENDIF
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS
						ENDDO
					ELSE
						IF m.PosWeek > 0
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS * (m.PosWeek - 1)
						ELSE
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS * 52
							IF YEAR(m.CalcDate) != YEAR(m.RefDate)
								m.CalcDate = m.CalcDate - WEEK_IN_SECONDS
							ENDIF
							m.CalcDate = m.CalcDate - WEEK_IN_SECONDS * ABS(m.PosWeek + 1)
						ENDIF
						IF YEAR(m.CalcDate) = YEAR(m.RefDate)
							IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate,m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
								m.Accepted = .T.
							ENDIF
						ENDIF
					ENDIF
				ENDFOR

			CASE m.Rule.Freq == "MONTHLY" AND !ISNULL(m.Rule.ByMonthDay)
				* limit
				IF ASCAN(This.ByDay, SUBSTR(":SUMOTUWETHFRSA", DOW(m.RefDate) * 2, 2)) != 0
					IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.RefDate, 1)) = 0
							m.Dates.Add(m.RefDate, TTOC(m.RefDate, 1))
						ENDIF
					ENDIF
				ENDIF

			CASE m.Rule.Freq == "MONTHLY"
				* expand
				FOR m.Entry = 1 TO ALEN(This.ByDay)
					m.CalcDate = DATETIME(YEAR(m.RefDate), MONTH(m.RefDate), 1, HOUR(m.RefDate), MINUTE(m.RefDate), SEC(m.RefDate))
					m.PosWeek = VAL(This.ByDay(m.Entry))
					m.DWeek = VAL(STREXTRACT("SU:1:MO:2:TU:3:WE:4:TH:5:FR:6:SA:7:", IIF(m.PosWeek != 0, RIGHT(This.ByDay(m.Entry), 2), This.ByDay(m.Entry)) + ":", ":"))
					m.CalcDate = m.CalcDate + 86400 * IIF(m.DWeek >= DOW(m.CalcDate), m.DWeek - DOW(m.CalcDate), m.DWeek + 7 - DOW(m.CalcDate))
					IF m.PosWeek = 0
						DO WHILE MONTH(m.CalcDate) = MONTH(m.RefDate)
							IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
								m.Accepted = .T.
							ENDIF
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS
						ENDDO
					ELSE
						IF m.PosWeek > 0
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS * (m.PosWeek - 1)
						ELSE
							m.CalcDate = m.CalcDate + WEEK_IN_SECONDS * 4
							IF MONTH(m.CalcDate) != MONTH(m.RefDate)
								m.CalcDate = m.CalcDate - WEEK_IN_SECONDS
							ENDIF
							m.CalcDate = m.CalcDate - WEEK_IN_SECONDS * ABS(m.PosWeek + 1)
						ENDIF
						IF MONTH(m.CalcDate) = MONTH(m.RefDate)
							IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
								IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
									m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
								ENDIF
								m.Accepted = .T.
							ENDIF
						ENDIF
					ENDIF
				ENDFOR

			CASE m.Rule.Freq == "WEEKLY"
				* expand
				m.FDWeek = VAL(STREXTRACT("SU:1:MO:2:TU:3:WE:4:TH:5:FR:6:SA:7:", NVL(m.Rule.WkSt, "MO") + ":", ":"))
				m.CalcDate = m.RefDate - 86400 * (DOW(m.RefDate, m.FDWeek) - 1)
				FOR m.Entry = 1 TO 7
					IF ASCAN(This.ByDay, SUBSTR(":SUMOTUWETHFRSA", DOW(m.CalcDate) * 2, 2)) != 0
						IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
							IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
								m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
							ENDIF
							m.Accepted = .T.
						ENDIF
					ENDIF
					m.CalcDate = m.CalcDate + DAY_IN_SECONDS
				ENDFOR

			OTHERWISE
				* limit
				m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, 0, "RRRSSS")
				IF ASCAN(This.ByDay, SUBSTR(":SUMOTUWETHFRSA", DOW(m.CalcDate) * 2, 2)) != 0
					IF This.ApplyByHour(m.Rule, m.Start, m.CalcDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
							m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
						ENDIF
						m.Accepted = .T.
					ENDIF
				ENDIF
			ENDCASE
		ELSE
			* rule not set? pass to next rule
			m.Accepted = This.ApplyByHour(m.Rule, m.Start, m.RefDate, m.Dates)
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyByHour (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .F.

		* rule set?
		IF !ISNULL(m.Rule.ByHour)
			* limit
			IF m.Rule.Freq == "SECONDLY" OR m.Rule.Freq == "MINUTELY" OR m.Rule.Freq == "HOURLY"
				IF ASCAN(This.ByHour, HOUR(m.RefDate)) != 0
					IF This.ApplyByMinute(m.Rule, m.Start, m.RefDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.RefDate, 1)) = 0
							m.Dates.Add(m.RefDate, TTOC(m.RefDate, 1))
						ENDIF
						m.Accepted = .T.
					ENDIF
				ENDIF
			ELSE
			* expand
				FOR m.Entry = 1 TO ALEN(This.ByHour)
					m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, This.ByHour(m.Entry), "RRRVSS")
					IF This.ApplyByMinute(m.Rule, m.Start, m.CalcDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
							m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
						ENDIF
						m.Accepted = .T.
					ENDIF
				ENDFOR
			ENDIF
		ELSE
			* rule not set? pass to next rule
			m.Accepted = This.ApplyByMinute(m.Rule, m.Start, m.RefDate, m.Dates)
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyByMinute (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .F.

		* rule set?
		IF !ISNULL(m.Rule.ByMinute)
			* limit
			IF m.Rule.Freq == "SECONDLY" OR m.Rule.Freq == "MINUTELY"
				IF ASCAN(This.ByMinute, MINUTE(m.RefDate)) != 0
					IF This.ApplyBySecond(m.Rule, m.Start, m.RefDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.RefDate, 1)) = 0
							m.Dates.Add(m.RefDate, TTOC(m.RefDate, 1))
						ENDIF
						m.Accepted = .T.
					ENDIF
				ENDIF
			ELSE
			* expand
				FOR m.Entry = 1 TO ALEN(This.ByMinute)
					m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, This.ByMinute(m.Entry), "RRRRVS")
					IF This.ApplyBySecond(m.Rule, m.Start, m.CalcDate, m.Dates)
						IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
							m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
						ENDIF
						m.Accepted = .T.
					ENDIF
				ENDFOR
			ENDIF
		ELSE
			* rule not set? pass to next rule
			m.Accepted = This.ApplyBySecond(m.Rule, m.Start, m.RefDate, m.Dates)
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyBySecond (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection) AS Logical

		LOCAL CalcDate AS Datetime
		LOCAL Entry AS Integer
		LOCAL Accepted AS Logical

		m.Accepted = .T.
		* rule set?
		IF !ISNULL(m.Rule.BySecond)
			* limit
			IF m.Rule.Freq == "SECONDLY"
				IF ASCAN(This.BySecond, SEC(m.RefDate)) != 0
					IF m.Dates.GetKey(TTOC(m.RefDate, 1)) = 0
						m.Dates.Add(m.RefDate, TTOC(m.RefDate, 1))
					ENDIF
				ELSE
					m.Accepted = .F.
				ENDIF
			ELSE
			* expand
				FOR m.Entry = 1 TO ALEN(This.BySecond)
					m.CalcDate = This.GetStartBasedDate(m.Start, m.RefDate, This.BySecond(m.Entry), "RRRRRV")
					IF m.Dates.GetKey(TTOC(m.CalcDate, 1)) = 0
						m.Dates.Add(m.CalcDate, TTOC(m.CalcDate, 1))
					ENDIF
				ENDFOR
			ENDIF
		ENDIF

		RETURN m.Accepted

	ENDFUNC

	HIDDEN FUNCTION ApplyNoBy (Rule AS iCalTypeRECUR, Start AS Datetime, RefDate AS Datetime, Dates AS Collection)

		IF ISNULL(m.Rule.ByMonth) AND ISNULL(m.Rule.ByWeekNo) AND ISNULL(m.Rule.ByYearDay) AND ISNULL(m.Rule.ByMonthDay) ;
				AND ISNULL(m.Rule.ByDay) AND ISNULL(m.Rule.ByHour) AND ISNULL(m.Rule.ByMinute) AND ISNULL(m.Rule.BySecond)
			IF m.Dates.GetKey(TTOC(m.RefDate, 1)) = 0
				m.Dates.Add(m.RefDate, TTOC(m.RefDate, 1))
			ENDIF
		ENDIF

	ENDFUNC

	* get a safe starting interval date (next dates will always be valid)
	HIDDEN FUNCTION GetCurrentDate (Rule AS iCalTypeRECUR, Start AS Datetime)

		LOCAL MonthPart, DayPart AS Integer

		m.MonthPart = IIF(!m.Rule.Freq == "YEARLY", MONTH(m.Start), 1)
		m.DayPart = IIF(m.Rule.Freq == "YEARLY" OR m.Rule.Freq == "MONTHLY", 1, DAY(m.Start))

		RETURN DATETIME(YEAR(m.Start), m.MonthPart, m.DayPart, HOUR(m.Start), MINUTE(m.Start), SEC(m.Start))
		
	ENDFUNC

	* get a date based on a combination of start date, reference date, and designated value parts
	HIDDEN FUNCTION GetStartBasedDate (Start AS Datetime, RefDate AS Datetime, PartValue AS Integer, SourcePattern AS String)

		LOCAL CalcDate AS Datetime

		TRY
			m.CalcDate = DATETIME(IIF(SUBSTR(m.SourcePattern, 1, 1) == "V", m.PartValue, YEAR(IIF(SUBSTR(m.SourcePattern, 1, 1) == "S", m.Start, m.RefDate))), ;
											IIF(SUBSTR(m.SourcePattern, 2, 1) == "V", m.PartValue, MONTH(IIF(SUBSTR(m.SourcePattern, 2, 1) == "S", m.Start, m.RefDate))), ;
											IIF(SUBSTR(m.SourcePattern, 3, 1) == "V", m.PartValue, DAY(IIF(SUBSTR(m.SourcePattern, 3, 1) == "S", m.Start, m.RefDate))), ;
											IIF(SUBSTR(m.SourcePattern, 4, 1) == "V", m.PartValue, HOUR(IIF(SUBSTR(m.SourcePattern, 4, 1) == "S", m.Start, m.RefDate))), ;
											IIF(SUBSTR(m.SourcePattern, 5, 1) == "V", m.PartValue, MINUTE(IIF(SUBSTR(m.SourcePattern, 5, 1) == "S", m.Start, m.RefDate))), ;
											IIF(SUBSTR(m.SourcePattern, 6, 1) == "V", m.PartValue, SEC(IIF(SUBSTR(m.SourcePattern, 6, 1) == "S", m.Start, m.RefDate))))
		CATCH
			m.CalcDate = {:}
		ENDTRY

		RETURN m.CalcDate

	ENDFUNC

	* prepare a comfortable array to reference the items in a rule by collection
	HIDDEN FUNCTION RuleCollectionToArray (DestArrayName AS String, RuleCollection AS Collection)

		LOCAL Entry AS Integer
		LOCAL DestArray AS String

		IF !ISNULL(m.RuleCollection)
			m.DestArray = "This." + m.DestArrayName
			DIMENSION &DestArray.(m.RuleCollection.Count)
			FOR m.Entry	= 1 TO m.RuleCollection.Count
				&DestArray.(m.Entry) = m.RuleCollection(m.Entry)
			ENDFOR
		ENDIF

	ENDFUNC

ENDDEFINE

DEFINE CLASS iCalPropSEQUENCE AS _iCalProperty

	ICName = "SEQUENCE"
	xICName = "sequence"

	DefaultValue = 0
	Value_DataType = "INTEGER"

ENDDEFINE

DEFINE CLASS iCalPropSTATUS AS _iCalProperty

	ICName = "STATUS"
	xICName = "status"
	AlternativeClasses = "VEVENT,VJOURNAL,VTODO"

	Enumeration = "TENTATIVE,CONFIRMED,CANCELLED,NEEDS-ACTION,COMPLETED,IN-PROCESS,DRAFT,FINAL"

ENDDEFINE

DEFINE CLASS iCalPropSTATUS_VEvent AS iCalPropSTATUS

	Enumeration = "TENTATIVE,CONFIRMED,CANCELLED"

ENDDEFINE

DEFINE CLASS iCalPropSTATUS_VJournal AS iCalPropSTATUS

	Enumeration = "DRAFT,FINAL,CANCELLED"

ENDDEFINE

DEFINE CLASS iCalPropSTATUS_VToDo AS iCalPropSTATUS

	Enumeration = "NEEDS-ACTION,COMPLETED,IN-PROCESS,CANCELLED"

ENDDEFINE

DEFINE CLASS iCalPropSUMMARY AS _iCalProperty

	ICName = "SUMMARY"
	xICName = "summary"

ENDDEFINE

DEFINE CLASS iCalPropTRANSP AS _iCalProperty

	ICName = "TRANSP"
	xICName = "transp"

	Enumeration = "OPAQUE,TRANSPARENT"
	Extensions = .F.

ENDDEFINE

DEFINE CLASS iCalPropTRIGGER AS _iCalProperty

	ICName = "TRIGGER"
	xICName = "trigger"

	Value_DataType = "DURATION"
	Value_AlternativeDataTypes = "DATE-TIME"

ENDDEFINE

DEFINE CLASS iCalPropTZID AS _iCalProperty

	ICName = "TZID"
	xICName = "tzid"

ENDDEFINE

DEFINE CLASS iCalPropTZNAME AS _iCalProperty

	ICName = "TZNAME"
	xICName = "tzname"

ENDDEFINE

DEFINE CLASS iCalPropTZOFFSETFROM AS _iCalProperty

	ICName = "TZOFFSETFROM"
	xICName = "tzoffsetfrom"

	Value_DataType = "UTC-OFFSET"

ENDDEFINE

DEFINE CLASS iCalPropTZOFFSETTO AS _iCalProperty

	ICName = "TZOFFSETTO"
	xICName = "tzoffsetto"

	Value_DataType = "UTC-OFFSET"

ENDDEFINE

DEFINE CLASS iCalPropTZURL AS _iCalProperty

	ICName = "TZURL"
	xICName = "tzurl"

	Value_DataType = "URI"

ENDDEFINE

DEFINE CLASS iCalPropUID AS _iCalProperty

	ICName = "UID"
	xICName = "uid"

ENDDEFINE

DEFINE CLASS iCalPropURL AS _iCalProperty

	ICName = "URL"
	xICName = "url"

	Value_DataType = "URI"

ENDDEFINE

DEFINE CLASS iCalPropVERSION AS _iCalProperty

	ICName = "VERSION"
	xICName = "version"
	ReadOnly = .T.

	DefaultValue = "2.0"

ENDDEFINE

