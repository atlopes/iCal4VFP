*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
*!*	|                                                                    |
*!*	|    Source and docs: https://bitbucket.org/atlopes/ical4vfp         |
*!*	|                                                                    |
*!*	+--------------------------------------------------------------------+

*!*	iCalendar types classes

*!*	Adds the classes to support iCalendar structured data types:
*!*	DURATION, PERIOD and RECUR

*!*	These types have Parse and Serialize methods, and support the value initialization at instantiation time.
*!*	Value processing depends on the data structure which is exposed to the intantiator through VFP object properties.

*!*	To avoid a cross-dependency, icalendar.prg is not called, but its classes are assumed to be in scope.

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

#DEFINE	SAFETHIS		ASSERT !USED("This") AND TYPE("This") == "O"

DEFINE CLASS iCalTypeDURATION AS _iCalType

	Positive = .T.
	Weeks = .NULL.
	Days = .NULL.
	Hours = .NULL.
	Minutes = .NULL.
	Seconds = .NULL.

	_MemberData =	'<VFPData>' + ;
							'<memberdata name="positive" type="property" display="Positive"/>' + ;
							'<memberdata name="weeks" type="property" display="Weeks"/>' + ;
							'<memberdata name="days" type="property" display="Days"/>' + ;
							'<memberdata name="hours" type="property" display="Hours"/>' + ;
							'<memberdata name="minutes" type="property" display="Minutes"/>' + ;
							'<memberdata name="seconds" type="property" display="Seconds"/>' + ;
							'<memberdata name="calculate" type="property" display="Calculate"/>' + ;
						'</VFPData>'

	FUNCTION Serialize (Level AS String) AS String

		SAFETHIS

		ASSERT VARTYPE(m.Level) == "C"

		LOCAL Serialized AS String
		LOCAL Signature AS String

		m.Serialized = ""
		m.Signature = IIF(This.Positive, "", "-") + "P"
	
		IF !ISNULL(This.Weeks)
			m.Serialized = m.Signature + TRANSFORM(INT(This.Weeks)) + "W"
		ELSE
			IF !ISNULL(This.Days)
				m.Serialized = m.Signature + TRANSFORM(INT(This.Days)) + "D"
			ENDIF
			IF !ISNULL(This.Hours) OR !ISNULL(This.Minutes) OR !ISNULL(This.Seconds)
				m.Serialized = EVL(m.Serialized, m.Signature) + "T" + TRANSFORM(INT(NVL(This.Hours, 0))) + "H"
				IF !ISNULL(This.Minutes) OR !ISNULL(This.Seconds)
					m.Serialized = m.Serialized + TRANSFORM(INT(NVL(This.Minutes, 0))) + "M"
					IF !ISNULL(This.Seconds)
						m.Serialized = m.Serialized + TRANSFORM(INT(NVL(This.Seconds, 0))) + "S"
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		RETURN m.Serialized

	ENDFUNC

	FUNCTION Parse (Serialized AS String, Level AS String) AS Logical

		SAFETHIS

		ASSERT VARTYPE(m.Serialized) + VARTYPE(m.Level) == "CC"

		LOCAL RegExp AS VBScript_RegExp_55.RegExp
		LOCAL Matches AS VBScript_RegExp_55.MatchCollection
		LOCAL Match AS VBScript_RegExp_55.Match
		LOCAL Parsed AS Logical

		m.Parsed = .F.
		STORE .NULL. TO This.Weeks, This.Days, This.Hours, This.Minutes, This.Seconds

		m.RegExp = CREATEOBJECT("VBScript.RegExp")
		m.RegExp.Ignorecase = .T.

		m.RegExp.Pattern = "^([+-])?P((\d+W)|((\d+D)?T(\d+H)((\d+M)(\d+S)?)?|(\d+D))$)"
		m.Matches = m.RegExp.Execute(m.Serialized)
		IF m.Matches.Count > 0
			m.Match = m.Matches.Item(0)
			
			IF m.Match.Value == m.Serialized

				This.Positive = LEFT(m.Serialized, 1) != "-"

				This.Weeks = VAL(m.Match.SubMatches(2))
				This.Days = VAL(NVL(m.Match.SubMatches(4), m.Match.SubMatches(9)))
				This.Hours = MIN(VAL(m.Match.SubMatches(5)), 23)
				This.Minutes = MIN(VAL(m.Match.SubMatches(7)), 59)
				This.Seconds = MIN(VAL(m.Match.SubMatches(8)), 59)

				m.Parsed = .T.
			ENDIF
		ENDIF

		RETURN m.Parsed

	ENDFUNC

	FUNCTION Calculate () AS Integer

		SAFETHIS

		LOCAL Calculated AS Integer

		m.Calculated = NVL(This.Weeks, 0) * 7 * 86400
		m.Calculated = m.Calculated + NVL(This.Days, 0) * 86400
		m.Calculated = m.Calculated + NVL(This.Hours, 0) * 3600
		m.Calculated = m.Calculated + NVL(This.Minutes, 0) * 60
		m.Calculated = m.Calculated + NVL(This.Seconds, 0)

		IF !This.Positive
			m.Calculated = -m.Calculated
		ENDIF

		RETURN m.Calculated

	ENDFUNC

ENDDEFINE

DEFINE CLASS iCalTypePERIOD AS _iCalType

	DateStart = .NULL.
	DateEnd = .NULL.
	Duration = .NULL.

	_MemberData =	'<VFPData>' + ;
							'<memberdata name="datestart" type="property" display="DateStart"/>' + ;
							'<memberdata name="dateend" type="property" display="DateEnd"/>' + ;
							'<memberdata name="duration" type="property" display="Duration"/>' + ;
						'</VFPData>'

	FUNCTION Serialize (Level AS String) AS String

		SAFETHIS

		ASSERT VARTYPE(m.Level) == "C"

		LOCAL Serialized

		m.Serialized = ""
		IF !EMPTY(NVL(This.DateStart, {})) AND (!EMPTY(NVL(This.DateEnd, {})) OR !ISNULL(This.Duration))
			IF !EMPTY(NVL(This.DateEnd, {}))
				m.Serialized = TRANSFORM(TTOC(This.DateEnd, 1), "@R 99999999T999999") + IIF(This.IsUTC, "Z", "")
			ELSE
				m.Serialized = This.Duration.Serialize(m.Level)
			ENDIF
			IF !EMPTY(m.Serialized)
				m.Serialized = TRANSFORM(TTOC(This.DateStart, 1), "@R 99999999T999999") + IIF(This.IsUTC, "Z", "") + "/" + ;
									m.Serialized
			ENDIF
		ENDIF

		RETURN m.Serialized

	ENDFUNC
			
	FUNCTION Parse (Serialized AS String, Level AS String) AS Logical

		SAFETHIS

		ASSERT VARTYPE(m.Serialized) + VARTYPE(m.Level) == "CC"

		LOCAL RegExp AS VBScript_RegExp_55.RegExp
		LOCAL Matches AS VBScript_RegExp_55.MatchCollection
		LOCAL Match AS VBScript_RegExp_55.Match
		LOCAL Parsed AS Logical

		m.Parsed = .F.
		STORE .NULL. TO This.DateStart, This.DateEnd, This.Duration
		This.IsUTC = .F.

		m.RegExp = CREATEOBJECT("VBScript.RegExp")
		m.RegExp.Ignorecase = .T.

		m.RegExp.Pattern = "^(\d{8}T\d{6}Z?)/((\d{8}T\d{6}Z?)|(\w+))$"
		m.Matches = m.RegExp.Execute(m.Serialized)
		IF m.Matches.Count > 0
			m.Match = m.Matches.Item(0)
			
			IF m.Match.Value == m.Serialized

				This.DateStart = EVALUATE("{^" + TRANSFORM(CHRTRAN(m.Match.SubMatches(0), "T", ""), "@R 9999-99-99 99:99:99") + "}")
				This.IsUTC = "Z" $ m.Match.SubMatches(1)
				IF !ISNULL(m.Match.SubMatches(2))
					This.DateEnd = EVALUATE("{^" + TRANSFORM(CHRTRAN(m.Match.SubMatches(2), "T", ""), "@R 9999-99-99 99:99:99") + "}")
				ELSE
					This.Duration = CREATEOBJECT("iCalTypeDURATION", m.Match.SubMatches(3))
				ENDIF

				m.Parsed = !EMPTY(NVL(This.DateStart, {})) AND (!EMPTY(NVL(This.DateEnd, {})) OR !ISNULL(This.Duration))
			ENDIF
		ENDIF

		RETURN m.Parsed
	ENDFUNC

ENDDEFINE

DEFINE CLASS iCalTypeRECUR AS _iCalType

	Freq = .NULL.
	Until = .NULL.
	Count = .NULL.
	Interval = .NULL.
	BySecond = .NULL.
	ByMinute = .NULL.
	ByHour = .NULL.
	ByDay = .NULL.
	ByMonthDay = .NULL.
	ByYearDay = .NULL.
	ByWeekNo = .NULL.
	ByMonth = .NULL.
	BySetPos = .NULL.
	WkSt = .NULL.

	_MemberData =	'<VFPData>' + ;
							'<memberdata name="freq" type="property" display="Freq"/>' + ;
							'<memberdata name="until" type="property" display="Until"/>' + ;
							'<memberdata name="count" type="property" display="Count"/>' + ;
							'<memberdata name="interval" type="property" display="Interval"/>' + ;
							'<memberdata name="bysecond" type="property" display="BySecond"/>' + ;
							'<memberdata name="byminute" type="property" display="ByMinute"/>' + ;
							'<memberdata name="byhour" type="property" display="ByHour"/>' + ;
							'<memberdata name="byday" type="property" display="ByDay"/>' + ;
							'<memberdata name="bymonthday" type="property" display="ByMonthDay"/>' + ;
							'<memberdata name="byyearday" type="property" display="ByYearDay"/>' + ;
							'<memberdata name="byweekno" type="property" display="ByWeekNo"/>' + ;
							'<memberdata name="bymonth" type="property" display="ByMonth"/>' + ;
							'<memberdata name="bysetpos" type="property" display="BySetPos"/>' + ;
							'<memberdata name="wkst" type="property" display="WkSt"/>' + ;
						'</VFPData>'

	FUNCTION Serialize (Level AS String) AS String

		SAFETHIS

		ASSERT VARTYPE(m.Level) == "C"

		LOCAL Serialized AS String
		LOCAL ListEntry AS String
		LOCAL ARRAY ByCollections(1)
		LOCAL ByCollection AS String
		LOCAL RefCollection AS Collection

		m.Serialized = ""
		IF !ISNULL(This.Freq)

			m.Serialized = "FREQ=" + This.Freq
			IF !EMPTY(NVL(This.Until, {}))
				IF VARTYPE(This.Until) == "T"
					m.Serialized = m.Serialized + ";UNTIL=" + TRANSFORM(TTOC(This.Until, 1), "@R 99999999T999999") + IIF(This.IsUTC, "Z", "") 
				ELSE
					m.Serialized = m.Serialized + ";UNTIL=" + DTOS(This.Until)
				ENDIF
			ELSE
				IF !ISNULL(This.Count)
					m.Serialized = m.Serialized + ";COUNT=" + TRANSFORM(INT(This.Count))
				ENDIF
			ENDIF

			IF !ISNULL(This.Interval)
				m.Serialized = m.Serialized + ";INTERVAL=" + TRANSFORM(INT(This.Interval))
			ENDIF

			ALINES(m.ByCollections, "BySecond,ByMinute,ByHour,ByDay,ByMonthDay,ByYearDay,ByWeekNo,ByMonth,BySetPos", 0, ",")
			FOR EACH m.ByCollection IN m.ByCollections
				m.RefCollection = EVALUATE("This." + m.ByCollection)
				IF !ISNULL(m.RefCollection)
					m.Serialized = m.Serialized + ";" + UPPER(m.ByCollection) + "="
					FOR EACH m.ListEntry IN m.RefCollection
						m.Serialized = m.Serialized + TRANSFORM(m.ListEntry) + ","
					ENDFOR
					m.Serialized = LEFT(m.Serialized, LEN(m.Serialized) - 1)
				ENDIF
			ENDFOR

			IF !ISNULL(This.WkSt)
				m.Serialized = m.Serialized + ";WKST=" + This.WkSt
			ENDIF
			
		ENDIF

		RETURN m.Serialized

	ENDFUNC

	FUNCTION Parse (Serialized AS String, Level AS String) AS Logical

		SAFETHIS

		ASSERT VARTYPE(m.Serialized) + VARTYPE(m.Level) == "CC"

		LOCAL ARRAY Segments(1)
		LOCAL Segment AS String
		LOCAL SegmentName AS String
		LOCAL SegmentValue AS String
		LOCAL RegExp AS VBScript_RegExp_55.RegExp
		LOCAL Matches AS VBScript_RegExp_55.MatchCollection
		LOCAL Match AS VBScript_RegExp_55.Match
		LOCAL Parsed AS Logical
		LOCAL Abort AS Logical
		LOCAL MaxValue AS Number
		LOCAL RefCollection AS Collection
		LOCAL ARRAY SubSegments(1)
		LOCAL SubSegment AS String

		m.Parsed = .F.
		STORE .NULL. TO This.Freq, This.Until, This.Count, This.Interval, This.BySecond, This.ByMinute, This.ByHour, This.ByDay, ;
			This.ByMonthDay, This.ByYearDay, This.ByWeekNo, This.ByMonth, This.BySetPos, This.WkSt
		This.IsUTC = .F.

		ALINES(m.Segments, m.Serialized, 1, ";")

		m.RegExp = CREATEOBJECT("VBScript.RegExp")
		m.RegExp.Ignorecase = .T.

		FOR EACH m.Segment IN m.Segments

			m.SegmentName = UPPER(STREXTRACT(m.Segment, "", "="))
			m.SegmentValue = STREXTRACT(m.Segment, "=", "")

			IF ISNULL(This.Freq) AND !(m.SegmentName == "FREQ")
				EXIT
			ENDIF

			m.Abort = .T.
			DO CASE
			CASE m.SegmentName == "FREQ" AND ISNULL(This.Freq)
				m.RegExp.Pattern = "^(SECONDLY|MINUTELY|HOURLY|DAILY|WEEKLY|MONTHLY|YEARLY)$"
				IF m.RegExp.Test(m.SegmentValue)
					This.Freq = UPPER(m.SegmentValue)
					m.Abort = .F.
				ENDIF

			CASE m.SegmentName == "UNTIL" AND ISNULL(This.Until) AND ISNULL(This.Count)
				m.RegExp.Pattern = "^(\d{8}T\d{6}Z?)|(\d{8})$"
				IF m.RegExp.Test(m.SegmentValue)
					TRY
						IF "T" $ m.SegmentValue
							This.Until = EVALUATE("{^" + TRANSFORM(CHRTRAN(m.SegmentValue, "T", ""), "@R 9999-99-99 99:99:99") + "}")
						ELSE
							This.Until = EVALUATE("{^" + TRANSFORM(m.SegmentValue, "@R 9999-99-999") + "}")
						ENDIF
						IF "Z" $ m.SegmentValue
							This.IsUTC = .T.
						ENDIF
						m.Abort = .F.
					CATCH
					ENDTRY
				ENDIF

			CASE m.SegmentName == "COUNT" AND ISNULL(This.Count) AND ISNULL(This.Until)
				m.RegExp.Pattern = "^\d+$"
				IF m.RegExp.Test(m.SegmentValue)
					This.Count = VAL(m.SegmentValue)
					m.Abort = .F.
				ENDIF

			CASE m.SegmentName == "INTERVAL" AND ISNULL(This.Interval)
				m.RegExp.Pattern = "^\d+$"
				IF m.RegExp.Test(m.SegmentValue)
					This.Interval = VAL(m.SegmentValue)
					m.Abort = .F.
				ENDIF

			CASE m.SegmentName == "BYSECOND" OR m.SegmentName == "BYMINUTE" OR m.SegmentName == "BYHOUR"
				m.RegExp.Pattern = "^\d+(,\d+)*$"
				IF m.RegExp.Test(m.SegmentValue)
					m.RefCollection = CREATEOBJECT("Collection")
					DO CASE
					CASE m.SegmentName == "BYSECOND" AND ISNULL(This.BySecond)
						This.BySecond = m.RefCollection
						m.MaxValue = 60
					CASE m.SegmentName == "BYMINUTE" AND ISNULL(This.ByMinute)
						This.ByMinute = m.RefCollection
						m.MaxValue = 59
					CASE m.SegmentName == "BYHOUR" AND ISNULL(This.ByHour)
						This.ByHour = m.RefCollection
						m.MaxValue = 23
					OTHERWISE
						m.RefCollection = .NULL.
						EXIT
					ENDCASE
					ALINES(m.SubSegments, m.SegmentValue, 0, ",")
					FOR EACH m.SubSegment IN m.SubSegments
						m.RefCollection.Add(MIN(VAL(m.SubSegment), m.MaxValue))
					ENDFOR
					m.Abort = .F.
				ENDIF

			CASE m.SegmentName == "BYDAY" AND ISNULL(This.ByDay)
				IF This.Freq == "MONTHLY" OR This.Freq == "YEARLY"
					m.RegExp.Pattern = "^([+-]?\d+)?(SU|MO|TU|WE|TH|FR|SA)(,([+-]?\d+)?(SU|MO|TU|WE|TH|FR|SA))*$"
				ELSE
					m.RegExp.Pattern = "^(SU|MO|TU|WE|TH|FR|SA)(,(SU|MO|TU|WE|TH|FR|SA))*$"
				ENDIF
				IF m.RegExp.Test(m.SegmentValue)
					This.ByDay = CREATEOBJECT("Collection")
					ALINES(m.SubSegments, m.SegmentValue, 0, ",")
					FOR EACH m.SubSegment IN m.SubSegments
						This.ByDay.Add(UPPER(m.SubSegment))
					ENDFOR
					m.Abort = .F.
				ENDIF

			CASE m.SegmentName == "BYMONTHDAY" OR m.SegmentName == "BYYEARDAY" OR m.SegmentName == "BYWEEKNO" OR m.SegmentName == "BYSETPOS"
				m.RegExp.Pattern = "^-?\d+(,-?\d+)*$"
				IF m.RegExp.Test(m.SegmentValue)
					m.RefCollection = CREATEOBJECT("Collection")
					DO CASE
					CASE m.SegmentName == "BYMONTHDAY" AND ISNULL(This.ByMonthDay) AND !(This.Freq == "WEEKLY")
						This.ByMonthDay = m.RefCollection
						m.MaxValue = 31
					CASE m.SegmentName == "BYYEARDAY" AND ISNULL(This.ByYearDay) AND ;
							!(This.Freq == "DAILY") AND !(This.Freq == "WEEKLY") AND !(This.Freq == "MONTHLY")
						This.ByYearDay = m.RefCollection
						m.MaxValue = 366
					CASE m.SegmentName == "BYWEEKNO" AND ISNULL(This.ByWeekNo) AND This.Freq == "YEARLY"
						This.ByWeekNo = m.RefCollection
						m.MaxValue = 53
					CASE m.SegmentName == "BYSETPOS" AND ISNULL(This.BySetPos)
						This.BySetPos = m.RefCollection
						m.MaxValue = 366 
					OTHERWISE
						m.RefCollection = .NULL.
						EXIT
					ENDCASE
					ALINES(m.SubSegments, m.SegmentValue, 0, ",")
					FOR EACH m.SubSegment IN m.SubSegments
						IF LEFT(m.SubSegment, 1) == "-"
							m.RefCollection.Add(MAX(MIN(VAL(m.SubSegment), -1), -m.MaxValue))
						ELSE
							m.RefCollection.Add(MIN(MAX(VAL(m.SubSegment), 1), m.MaxValue))
						ENDIF
					ENDFOR
					m.Abort = .F.
				ENDIF

			CASE m.SegmentName == "BYMONTH" AND ISNULL(This.ByMonth)
				m.RegExp.Pattern = "^\d+(,\d+)*$"
				IF m.RegExp.Test(m.SegmentValue)
					This.ByMonth = CREATEOBJECT("Collection")
					ALINES(m.SubSegments, m.SegmentValue, 0, ",")
					FOR EACH m.SubSegment IN m.SubSegments
						This.ByMonth.Add(MIN(MAX(VAL(m.SubSegment), 1), 12))
					ENDFOR
					m.Abort = .F.
				ENDIF

			CASE m.SegmentName == "WKST" AND ISNULL(This.WkSt)
				m.RegExp.Pattern = "^(SU|MO|TU|WE|TH|FR|SA)$"
				IF m.RegExp.Test(m.SegmentValue)
					This.WkSt = UPPER(m.SegmentValue)
					m.Abort = .F.
				ENDIF

			ENDCASE

			IF m.Abort
				EXIT
			ENDIF
		ENDFOR

		m.Parsed = !m.Abort

		IF m.Parsed
			IF ISNULL(This.Interval)
				This.Interval = 1
			ENDIF
			IF ISNULL(This.WkSt)
				This.WkSt = "MO"
			ENDIF
		ENDIF

		RETURN m.Parsed
		
	ENDFUNC

ENDDEFINE

* the basic iCalendar structured data type

DEFINE CLASS _iCalType AS _iCalBase

	IsUTC = .F.

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="isutc" type="property" display="IsUTC"/>' + ;
						"</VFPData>"

	FUNCTION Init (InitialValue AS String, Level AS String)

		IF PCOUNT() != 0
			IF PCOUNT() = 1
				m.Level = "UNDEFINED"
			ENDIF
			RETURN This.Parse(m.InitialValue, m.Level)
		ENDIF

	ENDFUNC

ENDDEFINE

