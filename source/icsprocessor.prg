*!"	iCalendar file reader and writer class

* install dependencies
DO LOCFILE("icalendar.prg")

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

#DEFINE	SAFETHIS		ASSERT !USED("This") AND TYPE("This") == "O"

DEFINE CLASS ICSProcessor AS Custom

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="readfile" type="method" display="ReadFile"/>' + ;
						'<memberdata name="readtext" type="method" display="ReadText"/>' + ;
						'<memberdata name="icstocursor" type="method" display="ICSToCursor"/>' + ;
						"</VFPData>"

	* read an ICS file
	FUNCTION ReadFile (ICSFile AS String) AS iCalendar

		ASSERT VARTYPE(m.ICSFile) == "C"

		LOCAL FHandle AS Integer
		LOCAL Result AS iCalendar
		LOCAL ICSLine AS String
		LOCAL Assembled AS String

		m.Result = .NULL.

		TRY
			* try to open the file for reading
			m.FHandle = FOPEN(m.ICSFile)
			* if fail, retry with an .ics extension
			IF m.FHandle = -1
				m.FHandle = FOPEN(FORCEEXT(m.ICSFile, "ics"))
			ENDIF

			* we have a file to read?
			IF m.FHandle != -1

				* prepare an iCalendar Core Object
				m.Result = CREATEOBJECT("iCalendar")

				* the lines to parse must be previously assembled
				m.Assembled = ""

				* get the first line
				m.ICSLine = FGETS(m.FHandle)
				* and skip the UTF-8 Byte-Order-Mark, if it exists
				IF LEFT(m.ICSLine, 3) == "" + 0hEFBBBF
					m.ICSLine = SUBSTR(m.ICSLine, 4)
				ENDIF

				* read all ics file
				DO WHILE !FEOF(m.FHandle)

					DO CASE
					* lines that start with a white space must be assembled with previous line
					CASE LEFT(m.ICSLine, 1) == " " OR LEFT(m.ICSLine, 1) == CHR(9)
						m.Assembled = m.Assembled + SUBSTR(m.ICSLine, 2)

					* lines starting with something else start a property or a component
					OTHERWISE
						* parse any pending line
						IF !EMPTY(m.Assembled)
							m.Result.Parse(m.Assembled)
						ENDIF
						* and start a new assembling
						m.Assembled= m.ICSLine
					ENDCASE
					* read the next line
					m.ICSLine = FGETS(m.FHandle, 8192)
				ENDDO

				* the pending assemblage will be parsed at the end
				m.Result.Parse(m.Assembled)
				* if everything was ok, m.Result will hold an iCalendar object
			ENDIF
		CATCH
			* something went wrong, so signal it
			m.Result = .NULL.

		FINALLY
			* close the file handle, regardless
			IF m.FHandle != -1
				FCLOSE(m.FHandle)
			ENDIF
		ENDTRY

		RETURN m.Result

	ENDFUNC

	* read an ICS from memory
	FUNCTION Read (ICSText AS String) AS iCalendar

		ASSERT VARTYPE(m.ICSText) == "C"

		LOCAL Result AS iCalendar
		LOCAL ARRAY ICSLines(1)
		LOCAL ICSLine AS String
		LOCAL Assembled AS String
		LOCAL ICSTextNoBOM AS String

		TRY
			* skip the UTF-8 Byte-Order-Mark, if it exists at the beginning of the source string
			IF LEFT(m.ICSText, 3) == "" + 0hEFBBBF
				m.ICSTextNoBOM = SUBSTR(m.ICSText, 4)
			ELSE
				m.ICSTextNoBOM = m.ICSText
			ENDIF

			* break the source string into lines
			ALINES(m.ICSLines, m.ICSTextNoBOM)

			* prepare an iCalendar Core Object
			m.Result = CREATEOBJECT("iCalendar")

			* the lines to parse must be previously assembled
			m.Assembled = ""

			* go through each line
			FOR EACH m.ICSLine IN m.ICSLines

				DO CASE
				* lines that start with a white space must be assembled with previous line
				CASE LEFT(m.ICSLine, 1) == " " OR LEFT(m.ICSLine, 1) == CHR(9)
					m.Assembled = m.Assembled + SUBSTR(m.ICSLine, 2)

				* lines starting with something else start a property or a component
				OTHERWISE
					* parse any pending line
					IF !EMPTY(m.Assembled)
						m.Result.Parse(m.Assembled)
					ENDIF
					* and start a new assembling
					m.Assembled= m.ICSLine
				ENDCASE
			ENDFOR
			
			* the pending assemblage will be parsed at the end
			m.Result.Parse(m.Assembled)
			* if everything was ok, m.Result will have an iCalendar object

		CATCH
			* something went wrong, so signal it
			m.Result = .NULL.

		ENDTRY

		RETURN m.Result

	ENDFUNC

	* read the essential of an ICS object into a VFP cursor
	FUNCTION ICSToCursor (ICSSource AS StringOriCalendar, CursorType AS String, CursorName AS String) AS Integer

		SAFETHIS

		ASSERT VARTYPE(m.ICSSource) $ "CO" AND VARTYPE(m.CursorType) + VARTYPE(m.CursorName) == "CC"

		LOCAL Source AS iCalendar

		* from an iCalendar, a string, or a file...
		DO CASE
		CASE VARTYPE(m.ICSSource) == "O"
			m.Source = m.ICSSource

		CASE MEMLINES(m.ICSSource) > 1
			m.Source = This.Read(m.ICSSource)

		OTHERWISE
			m.Source = This.ReadFile(m.ICSSource)
		ENDCASE

		IF ISNULL(m.Source)
			RETURN -1
		ENDIF

		* supported iCalendar object types, so far: events and to-dos
		DO CASE
		CASE UPPER(m.CursorType) == "EVENTS"
			RETURN This.EventsToCursor(m.Source, m.CursorName)

		CASE UPPER(m.CursorType) == "TODOS"
			RETURN This.TodosToCursor(m.Source, m.CursorName)

		OTHERWISE
			RETURN -2
		ENDCASE
	ENDFUNC

	HIDDEN FUNCTION EventsToCursor (Source AS iCalendar, CursorName AS String) AS Integer

		LOCAL EventIndex AS Integer
		LOCAL ICComponent AS _iCalComponent
		LOCAL ICProperty AS _iCalProperty
		LOCAL ICParameter AS _iCalParameter
		LOCAL ObjBuffer AS Object

		CREATE CURSOR (m.CursorName) ;
			(IdEvent Varchar(200), ;
				Summary Varchar(200), ;
				Timezone Varchar(100), ;
				Start Datetime, ;
				End Datetime, ;
				Organizer Varchar(200), ;
				OrganizerName Varchar(100), ;
				Attendee Varchar(200), ;
				AttendeeName Varchar(100), ;
				Status Varchar(30), ;
				Location Varchar(200), ;
				Priority Varchar(200), ;
				URL Varchar(200), ;
				Comments Memo, ;
				Description Memo)

		FOR m.EventIndex = 1 TO m.Source.GetICComponentsCount("VEVENT")

			m.ICComponent = m.Source.GetICComponent("VEVENT", m.EventIndex)
			
			SCATTER MEMO BLANK NAME ObjBuffer

			m.ObjBuffer.IdEvent = NVL(m.ICComponent.GetICPropertyValue("UID"), "Undefined")
			m.ObjBuffer.Summary = NVL(m.ICComponent.GetICPropertyValue("SUMMARY"), "")
			m.ObjBuffer.Start = NVL(m.ICComponent.GetICPropertyValue("DTSTART"), {})
			m.ObjBuffer.End = NVL(m.ICComponent.GetICPropertyValue("DTEND"), {})
			m.ICProperty = m.ICComponent.GetICProperty("ORGANIZER")
			IF !ISNULL(m.ICProperty)
				m.ObjBuffer.Organizer = NVL(m.ICProperty.GetValue(), "")
				m.ObjBuffer.OrganizerName = NVL(m.ICProperty.GetICParameterValue("CN"), "")
			ENDIF
			m.ObjBuffer.Location = NVL(m.ICComponent.GetICPropertyValue("LOCATION"), "")
			m.ObjBuffer.Priority = NVL(m.ICComponent.GetICPropertyValue("PRIORITY"), "")
			m.ObjBuffer.URL = NVL(m.ICComponent.GetICPropertyValue("URL"), "")
			m.ObjBuffer.Comments = NVL(m.ICComponent.GetICPropertyValue("COMMENT"), "")
			m.ObjBuffer.Description = NVL(m.ICComponent.GetICPropertyValue("DESCRIPTION"), "")

			This.AttendeesToCursor(m.ICComponent, m.ObjBuffer)

		ENDFOR

		RETURN RECCOUNT(m.CursorName)

	ENDFUNC

	HIDDEN FUNCTION TodosToCursor (Source AS iCalendar, CursorName AS String) AS Integer

		LOCAL NumAttendees AS Integer
		LOCAL TodoIndex AS Integer
		LOCAL AttendeeIndex AS Integer
		LOCAL ICComponent AS _iCalComponent
		LOCAL ICProperty AS _iCalProperty
		LOCAL ICParameter AS _iCalParameter
		LOCAL ObjBuffer AS Object

		CREATE CURSOR (m.CursorName) ;
			(IdTodo Varchar(200), ;
				Summary Varchar(200), ;
				Timezone Varchar(100), ;
				Start Datetime, ;
				Due Datetime, ;
				Organizer Varchar(200), ;
				OrganizerName Varchar(100), ;
				Attendee Varchar(200), ;
				AttendeeName Varchar(100), ;
				Status Varchar(30), ;
				Location Varchar(200), ;
				Priority Varchar(200), ;
				URL Varchar(200), ;
				Comments Memo, ;
				Description Memo)

		FOR m.TodoIndex = 1 TO m.Source.GetICComponentsCount("VTODO")

			m.ICComponent = m.Source.GetICComponent("VTODO", m.TodoIndex)
			
			SCATTER MEMO BLANK NAME ObjBuffer

			m.ObjBuffer.IdTodo = NVL(m.ICComponent.GetICPropertyValue("UID"), "Undefined")
			m.ObjBuffer.Summary = NVL(m.ICComponent.GetICPropertyValue("SUMMARY"), "")
			m.ObjBuffer.Start = NVL(m.ICComponent.GetICPropertyValue("DTSTART"), {})
			m.ObjBuffer.Due = NVL(m.ICComponent.GetICPropertyValue("DUE"), {})
			m.ICProperty = m.ICComponent.GetICProperty("ORGANIZER")
			IF !ISNULL(m.ICProperty)
				m.ObjBuffer.Organizer = NVL(m.ICProperty.GetValue(), "")
				m.ObjBuffer.OrganizerName = NVL(m.ICProperty.GetICParameterValue("CN"), "")
			ENDIF
			m.ObjBuffer.Location = NVL(m.ICComponent.GetICPropertyValue("LOCATION"), "")
			m.ObjBuffer.Priority = NVL(m.ICComponent.GetICPropertyValue("PRIORITY"), "")
			m.ObjBuffer.URL = NVL(m.ICComponent.GetICPropertyValue("URL"), "")
			m.ObjBuffer.Comments = NVL(m.ICComponent.GetICPropertyValue("COMMENT"), "")
			m.ObjBuffer.Description = NVL(m.ICComponent.GetICPropertyValue("DESCRIPTION"), "")

			This.AttendeesToCursor(m.ICComponent, m.ObjBuffer)
		ENDFOR

		RETURN RECCOUNT(m.CursorName)

	ENDFUNC

	HIDDEN FUNCTION AttendeesToCursor (ICSource AS iCalendar, CursorBuffer AS Object)

		LOCAL NumAttendees AS Integer
		LOCAL AttendeeIndex AS Integer
		LOCAL Attendee AS iCalPropATTENDEE

		m.NumAttendees = m.ICSource.GetICPropertiesCount("ATTENDEE")
		IF m.NumAttendees > 0
			FOR m.AttendeeIndex = 1 TO m.NumAttendees

				m.Attendee = m.ICSource.GetICProperty("ATTENDEE", m.AttendeeIndex)

				m.CursorBuffer.Attendee = m.Attendee.GetValue()
				m.CursorBuffer.AttendeeName = NVL(m.Attendee.GetICParameterValue("CN"), "")
				m.CursorBuffer.Status = NVL(m.Attendee.GetICParameterValue("PARTSTAT"), "")

				APPEND BLANK
				GATHER MEMO NAME m.CursorBuffer
			ENDFOR
		ELSE
			APPEND BLANK
			GATHER MEMO NAME m.CursorBuffer
		ENDIF

	ENDFUNC

ENDDEFINE
