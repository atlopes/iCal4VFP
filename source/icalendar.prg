*!*	iCalendar VFP classes

*!*	In this file:
*!*		_iCalBase
*!*			_iCalElement
*!*				_iCalComponent
*!*				_iCalValueHandler
*!*					_iCalProperty
*!*					_iCalParameter
*!*			_iCalValue
*!*			_iCalValueInfo

*!*	In icaltypes.prg
*!*		[_iCalBase]
*!*			_iCalType
*!*				classes for iCalendar structured types

* install dependencies
DO LOCFILE("tokenizer.prg")
DO LOCFILE("icaltypes.prg")

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

#DEFINE	SAFETHIS		ASSERT !USED("This") AND TYPE("This") == "O"

* serialization constants

* the max width of iCalendar lines (for serialization)
#DEFINE	ICAL_LINE	72

* set to .F. to inhibit the ^-escape mechanism
#DEFINE	RFC6868	.T.

* EOL and continuation prefix
#DEFINE	CRLF			CHR(13) + CHR(10)
#DEFINE	HTAB			CHR(9)

* general class for iCalendar component objects

DEFINE CLASS _iCalComponent AS _iCalElement

	Level = "COMPONENT"

	* the properties that integrate the component
	DIMENSION ICProperties(1)
	ICPropertiesCount = 0
	* sub-components
	DIMENSION ICComponents(1)
	ICComponentsCount = 0

	* what components may be included
	IncludeComponents = "*NONE"

	* the component being ignored during parsing
	IgnoreComponent = .NULL.
	IgnoreCount = 0

	* the sub-component being parsed
	CurrentComponent = .NULL.

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="currentcomponent" type="property" display="CurrentComponent"/>' + ;
						'<memberdata name="iccomponents" type="property" display="ICComponents"/>' + ;
						'<memberdata name="iccomponentscount" type="property" display="ICComponentsCount"/>' + ;
						'<memberdata name="icproperties" type="property" display="ICProperties"/>' + ;
						'<memberdata name="icpropertiescount" type="property" display="ICPropertiesCount"/>' + ;
						'<memberdata name="ignorecomponent" type="property" display="IgnoreComponent"/>' + ;
						'<memberdata name="ignorecount" type="property" display="IgnoreCount"/>' + ;
						'<memberdata name="includecomponents" type="property" display="IncludeComponents"/>' + ;
						'<memberdata name="addicproperty" type="method" display="AddICProperty"/>' + ;
						'<memberdata name="addiccomponent" type="method" display="AddICComponent"/>' + ;
						'<memberdata name="geticcomponent" type="method" display="GetICComponent"/>' + ;
						'<memberdata name="geticcomponentscount" type="method" display="GetICComponentsCount"/>' + ;
						'<memberdata name="geticproperty" type="method" display="GetICProperty"/>' + ;
						'<memberdata name="geticpropertyvalue" type="method" display="GetICPropertyValue"/>' + ;
						'<memberdata name="geticpropertiescount" type="method" display="GetICPropertiesCount"/>' + ;
						"</VFPData>"

	* while destroying the object, clear the arrays of references to properties and components
	FUNCTION Destroy

		SAFETHIS

		LOCAL LoopIndex AS Integer

		FOR m.LoopIndex = This.ICPropertiesCount TO 1 STEP -1
			IF !ISNULL(This.ICProperties(m.LoopIndex))
				This.ICProperties(m.LoopIndex).Reset()
				This.ICProperties(m.LoopIndex) = .NULL.
			ENDIF
		ENDFOR

		FOR m.LoopIndex = This.ICComponentsCount TO 1 STEP -1
			IF !ISNULL(This.ICComponents(m.LoopIndex))
				This.ICComponents(m.LoopIndex).Reset()
				This.ICComponents(m.LoopIndex) = .NULL.
			ENDIF
		ENDFOR

	ENDFUNC

	* generate a component in iCalendar format
	FUNCTION Serialize () AS String

		SAFETHIS

		LOCAL Serialized AS String
		LOCAL Component AS _iCalComponent
		LOCAL Property AS _iCalProperty
		LOCAL LoopIndex AS Integer

		* mark the begin
		m.Serialized = "BEGIN:" + This.ICName + CRLF

		* generate all properties
		FOR m.LoopIndex = 1 TO This.ICPropertiesCount
			IF !ISNULL(This.ICProperties(m.LoopIndex))
				m.Serialized = m.Serialized + This.ICProperties(m.LoopIndex).Serialize()
			ENDIF
		ENDFOR

		* and then the components
		FOR m.LoopIndex = 1 TO This.ICComponentsCount
			IF !ISNULL(This.ICComponents(m.LoopIndex))
				m.Serialized = m.Serialized + This.ICComponents(m.LoopIndex).Serialize()
			ENDIF
		ENDFOR

		* mark that we are done
		m.Serialized = m.Serialized + "END:" + This.ICName + CRLF

		RETURN m.Serialized

	ENDFUNC

	* read an iCalendar formatted component part
	FUNCTION Parse (Serialized AS String) AS Logical

		SAFETHIS

		ASSERT VARTYPE(m.Serialized) == "C"

		LOCAL ComponentName AS String
		LOCAL ComponentObject AS _iCalComponent
		LOCAL PropertyName AS String
		LOCAL PropertyObject AS _iCalPropery

		DO CASE
		* in the process of ignoring a component?
		CASE !ISNULL(This.IgnoreComponent)
			* if an end of such component was found
			IF UPPER(TRIM(m.Serialized)) == "END:" + This.IgnoreComponent
				* decrease the counter (to avoid unbalanced definitions)
				This.IgnoreCount = This.IgnoreCount - 1
				IF This.IgnoreCount = 0
					* until we reached the end of the component definition
					This.IgnoreComponent = .NULL.
				ENDIF
			ELSE
				* found a recursive definition? just increase the counter
				IF UPPER(TRIM(m.Serialized)) == "BEGIN:" + This.IgnoreComponent
					This.IgnoreCount = This.IgnoreCount + 1
				ENDIF
				* all the rest is ignored...
			ENDIF

		* parsing a sub-component?
		CASE !ISNULL(This.CurrentComponent)
			* if the end of definition is reached, regain control
			IF UPPER(TRIM(m.Serialized)) == "END:" + This.CurrentComponent.ICName
				This.CurrentComponent = .NULL.
			ELSE
				* otherwise, pass the parsing downstream
				This.CurrentComponent.Parse(m.Serialized)
			ENDIF

		* a new sub-component?
		CASE UPPER(LEFT(m.Serialized, 6)) == "BEGIN:"
			m.ComponentName = UPPER(TRIM(SUBSTR(m.Serialized, 7)))
			m.ComponentObject = This.AddICComponent(This.iCalCreateObject("Comp", m.ComponentName, This.ICName))
			* the new sub-component is not recognized / accepted?
			IF ISNULL(m.ComponentObject)
				* then ignore its definition
				This.IgnoreComponent = m.ComponentName
				This.IgnoreCount = 1
			ELSE
				* otherwise, use it as a target for the following parsing
				This.CurrentComponent = m.ComponentObject
				m.ComponentObject = .NULL.
			ENDIF

		* ending our definition?
		CASE UPPER(TRIM(m.Serialized)) == "END:" + This.ICName
			RETURN

		* try to parse a property
		OTHERWISE
			m.PropertyName = UPPER(GETWORDNUM(m.Serialized, 1, ";:"))
			m.PropertyObject = This.AddICProperty(This.iCalCreateObject("Prop", m.PropertyName, This.ICName))
			* found a class for it? use the object's parser
			IF !ISNULL(m.PropertyObject)
				m.PropertyObject.Parse(m.Serialized)
			ENDIF
		ENDCASE

	ENDFUNC

	* add a propriety to the component
	FUNCTION AddICProperty (Property AS _iCalProperty) AS _iCalProperty

		SAFETHIS

		ASSERT VARTYPE(m.Property) $ "OX"

		IF !ISNULL(m.Property)
			This.ICPropertiesCount = This.ICPropertiesCount + 1
			DIMENSION This.ICProperties(This.ICPropertiesCount)
			This.ICProperties(This.ICPropertiesCount) = m.Property
		ENDIF
		
		RETURN m.Property

	ENDFUNC

	* add a sub-component to the component
	FUNCTION AddICComponent (Component AS _iCalComponent) AS _iCalComponent

		SAFETHIS

		ASSERT VARTYPE(m.Component) $ "OX"

		IF !ISNULL(m.Component)
			* but only if it is part of the allowed components
			IF (This.IncludeComponents == "*ANY" AND !(m.Component.ICName == This.ICName)) OR ATC("," + m.Component.ICName + ",", "," + This.IncludeComponents + ",") != 0
				This.ICComponentsCount = This.ICComponentsCount + 1
				DIMENSION This.ICComponents(This.ICComponentsCount)
				This.ICComponents(This.ICComponentsCount) = m.Component
				RETURN m.Component
			ENDIF
		ENDIF

		RETURN .NULL.

	ENDFUNC

	* get an iCalendar property that is added to the component
	FUNCTION GetICProperty (Id AS StringOrInteger, PropertyIndex AS Integer) AS _iCalProperty

		SAFETHIS

		ASSERT (PCOUNT() = 1 AND VARTYPE(m.Id) $ "NC") OR (PCOUNT() = 2 AND VARTYPE(m.Id) + VARTYPE(m.PropertyIndex) == "CN")

		LOCAL ARRAY Temp(1)
		ACOPY(This.ICProperties, m.Temp)

		IF PCOUNT() = 1
			IF VARTYPE(m.Id) == "N"
				RETURN This._GetIC(m.Id, 0, @m.Temp, This.ICPropertiesCount)
			ENDIF
			m.PropertyIndex = 1
		ENDIF

		RETURN This._GetIC(m.Id, m.PropertyIndex,  @m.Temp, This.ICPropertiesCount)

	ENDFUNC

	* get the value of a property that is added to the component
	FUNCTION GetICPropertyValue (Id AS StringOrInteger, PropertyIndex AS Integer) AS AnyType

		ASSERT (PCOUNT() = 1 AND VARTYPE(m.Id) $ "NC") OR (PCOUNT() = 2 AND VARTYPE(m.Id) + VARTYPE(m.PropertyIndex) == "CN")

		LOCAL ICProperty AS _iCalProperty

		IF PCOUNT() = 1
			IF VARTYPE(m.Id) == "N"
				m.ICProperty = This.GetICProperty(m.Id)
			ELSE
				m.ICProperty = This.GetICProperty(m.Id, 1)
			ENDIF
		ELSE
			m.ICProperty = This.GetICProperty(m.Id, m.PropertyIndex)
		ENDIF

		IF !ISNULL(m.ICProperty)
			RETURN m.ICProperty.GetValue()
		ENDIF

		RETURN .NULL.

	ENDFUNC

	* get the number of properties added to the component
	FUNCTION GetICPropertiesCount (Id AS String) AS Integer

		SAFETHIS

		ASSERT PCOUNT() = 0 OR VARTYPE(m.Id) == "C"

		LOCAL ARRAY Temp(1)
		ACOPY(This.ICProperties, m.Temp)

		IF PCOUNT() = 0
			m.Id = ""
		ENDIF

		RETURN This._GetICCount(m.Id, @m.Temp, This.ICPropertiesCount)

	ENDFUNC

	* get an iCalendar sub-component that is added to the component
	FUNCTION GetICComponent (Id AS StringOrInteger, ComponentIndex AS Integer) AS _iCalComponent

		SAFETHIS

		ASSERT (PCOUNT() = 1 AND VARTYPE(m.Id) $ "NC") OR (PCOUNT() = 2 AND VARTYPE(m.Id) + VARTYPE(m.ComponentIndex) == "CN")

		LOCAL ARRAY Temp(1)
		ACOPY(This.ICComponents, m.Temp)

		IF PCOUNT() = 1
			IF VARTYPE(m.Id) == "N"
				RETURN This._GetIC(m.Id, 0, @m.Temp, This.ICComponentsCount)
			ENDIF
			m.ComponentIndex = 1
		ENDIF

		RETURN This._GetIC(m.Id, m.ComponentIndex,  @m.Temp, This.ICComponentsCount)

	ENDFUNC

	* get the number of sub-components added to the component
	FUNCTION GetICComponentsCount (Id AS String) AS Integer

		SAFETHIS

		ASSERT PCOUNT() = 0 OR VARTYPE(m.Id) == "C"

		LOCAL ARRAY Temp(1)
		ACOPY(This.ICComponents, m.Temp)

		IF PCOUNT() = 0
			m.Id = ""
		ENDIF

		RETURN This._GetICCount(m.Id, @m.Temp, This.ICComponentsCount)

	ENDFUNC

ENDDEFINE

* general iCalendar property class definition

DEFINE CLASS _iCalProperty AS _iCalValueHandler

	Level = "PROPERTY"

	* the parameters that the property may have
	DIMENSION ICParameters(1)
	ICParametersCount = 0

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="icparameters" type="property" display="ICParameters"/>' + ;
						'<memberdata name="icparameterscount" type="property" display="ICParametersCount"/>' + ;
						'<memberdata name="addicparameter" type="method" display="AddICParameter"/>' + ;
						'<memberdata name="geticparameter" type="method" display="GetICParameter"/>' + ;
						'<memberdata name="geticparametervalue" type="method" display="GetICParameterValue"/>' + ;
						'<memberdata name="geticparameterscount" type="method" display="GetICParametersCount"/>' + ;
						"</VFPData>"

	* when terminating, remove all properties
	FUNCTION Destroy

		SAFETHIS

		LOCAL LoopIndex AS Integer

		FOR m.LoopIndex = This.ICParametersCount TO 1 STEP -1
			IF !ISNULL(This.ICParameters(m.LoopIndex))
				This.ICParameters(m.LoopIndex).Reset()
				This.ICParameters(m.LoopIndex) = .NULL.
			ENDIF
		ENDFOR

		DODEFAULT()

	ENDFUNC

	* add an iCalendar parameter to the property
	FUNCTION AddICParameter (Parameter AS _iCalParameter) AS _iCalParameter

		SAFETHIS

		ASSERT VARTYPE(m.Parameter) $ "OX"

		LOCAL ARRAY DataTypes(1)

		IF !ISNULL(m.Parameter)
			* the VALUE parameter is a special case: it may change the parent property data type
			IF m.Parameter.ICName == "VALUE" AND ;
					ALINES(m.DataTypes, This.HValue.Value.DataType + "," + This.HValue.Value.AlternativeDataTypes, 1 + 4, ",") > 1 AND ;
					(ASCAN(m.DataTypes, m.Parameter.GetValue(), -1, -1, 1, 2 + 4) != 0 OR This.Extensions)
				This.HValue.Value.DataType = UPPER(m.Parameter.GetValue())
			ENDIF

			This.ICParametersCount = This.ICParametersCount + 1
			DIMENSION This.ICParameters(This.ICParametersCount)
			This.ICParameters(This.ICParametersCount) = m.Parameter
		ENDIF

		RETURN m.Parameter

	ENDFUNC

	* get an iCalendar parameter that was added to the property
	FUNCTION GetICParameter (Id AS StringOrInteger, ParameterIndex AS Integer) AS _iCalParameter

		SAFETHIS

		ASSERT (PCOUNT() = 1 AND VARTYPE(m.Id) $ "NC") OR (PCOUNT() = 2 AND VARTYPE(m.Id) + VARTYPE(m.ParameterIndex) == "CN")

		LOCAL ARRAY Temp(1)
		ACOPY(This.ICParameters, m.Temp)

		IF PCOUNT() = 1
			IF VARTYPE(m.Id) == "N"
				RETURN This._GetIC(m.Id, 0, @m.Temp, This.ICParametersCount)
			ENDIF
			m.ParameterIndex = 1
		ENDIF

		RETURN This._GetIC(m.Id, m.ParameterIndex,  @m.Temp, This.ICParametersCount)

	ENDFUNC

	* get the value of an iCalendar parameter that was added to the property
	FUNCTION GetICParameterValue (Id AS StringOrInteger, ParameterIndex AS Integer) AS AnyType

		SAFETHIS

		ASSERT (PCOUNT() = 1 AND VARTYPE(m.Id) $ "NC") OR (PCOUNT() = 2 AND VARTYPE(m.Id) + VARTYPE(m.ParameterIndex) == "CN")

		LOCAL ICParameter AS _iCalParameter

		IF PCOUNT() = 1
			IF VARTYPE(m.Id) == "N"
				m.ICParameter = This.GetICParameter(m.Id)
			ELSE
				m.ICParameter = This.GetICParameter(m.Id, 1)
			ENDIF
		ELSE
			m.ICParameter = This.GetICParameter(m.Id, m.ParameterIndex)
		ENDIF

		IF !ISNULL(m.ICParameter)
			RETURN m.ICParameter.GetValue()
		ENDIF

		RETURN .NULL.

	ENDFUNC

	* get the number of paramters added to the property
	FUNCTION GetICParametersCount (Id AS String) AS Integer

		SAFETHIS

		ASSERT VARTYPE(m.Id) == "C"

		LOCAL ARRAY Temp(1)
		ACOPY(This.ICParameters, m.Temp)

		IF PCOUNT() = 0
			m.Id = ""
		ENDIF

		RETURN This._GetICCount(m.Id, @m.Temp, This.ICParametersCount)

	ENDFUNC

	* generate the property in iCalendar format
	FUNCTION Serialize () AS String

		SAFETHIS

		LOCAL Stream AS String
		LOCAL StreamLine AS String
		LOCAL StreamRest AS String
		LOCAL StreamUTF8 AS String
		LOCAL Serialized AS String
		LOCAL SerializedValue AS String
		LOCAL Parameter AS _iCalParameter
		LOCAL ParameterStream AS String
		LOCAL LoopIndex AS Integer

		m.Serialized = ""

		* get the property serialized value
		m.SerializedValue = This.HValue.Serialize(This.Level)

		* do not serialize empty properties
		IF !EMPTY(NVL(m.SerializedValue, ""))

			* at first, build a complete ANSI representation of the property
			m.Stream = This.ICName

			* including its parameters
			FOR m.LoopIndex = 1 TO This.ICParametersCount
				IF !ISNULL(This.ICParameters(m.LoopIndex))
					m.ParameterStream = This.ICParameters(m.LoopIndex).Serialize()
					IF !EMPTY(m.ParameterStream)
						m.Stream = m.Stream + ";" + m.ParameterStream
					ENDIF
				ENDIF
			ENDFOR

			m.Stream = m.Stream + ":" + m.SerializedValue

			* if needed, break the property contents into safe-width lines
			m.StreamLine = LEFT(m.Stream, ICAL_LINE)
			m.Stream = SUBSTR(m.Stream, ICAL_LINE + 1)
			DO WHILE !EMPTY(m.StreamLine)
				* encoded as utf-8
				m.StreamUTF8 = STRCONV(STRCONV(m.StreamLine, 1), 9)
				* if it goes over the limit
				IF LEN(m.StreamUTF8) > ICAL_LINE
					* reduce 4 ANSI characters at the time, by passing the last 4 to what is yet to be done, and try again
					m.Stream = RIGHT(m.StreamLine, 4) + m.Stream
					m.StreamLine = LEFT(m.StreamLine, LEN(m.StreamLine) - 4)
				ELSE
					* if it fits, then consider that it's serialized and move to the rest of the property contens
					m.Serialized = m.Serialized + IIF(!EMPTY(m.Serialized), HTAB, "") + m.StreamUTF8 + CRLF
					m.StreamLine = LEFT(m.Stream, ICAL_LINE)
					m.Stream = SUBSTR(m.Stream, ICAL_LINE + 1)
				ENDIF
			ENDDO

		ENDIF

		* done! (or empty...)
		RETURN m.Serialized

	ENDFUNC

	* read an iCalendar formatted property into a property object
	FUNCTION Parse (Serialized AS String) AS Logical

		SAFETHIS

		ASSERT VARTYPE(m.Serialized) == "C"

		LOCAL Parser AS Tokenizer

		LOCAL Parsed AS Logical

		LOCAL TokenIndex AS Integer

		LOCAL ParameterName AS String
		LOCAL ParameterValue AS String
		LOCAL ParameterSet AS Logical

		LOCAL ParameterObject AS _iCalParameter

		LOCAL SerializedANSI AS String

		* restart the property value and its parameters
		This.Destroy()

		m.SerializedANSI = STRCONV(STRCONV(m.Serialized, 11), 2)
		m.Parsed = .F.

		* prepare the tokenizer
		m.Parser = CREATEOBJECT("Tokenizer")
		m.Parser.IgnoreCase = .T.
		m.Parser.AddTokenPattern("^(([a-zA-Z-]+))[:;]", "NProp", 1, 0)
		m.Parser.AddTokenPattern("^(;([a-zA-Z-]+))=", "NParm", 1, 0)
		m.Parser.AddTokenPattern('^([=,]("[^"]*"))[,;:]', "VParm", 1, 0)
		m.Parser.AddTokenPattern('^([=,]([^,:;"]+))[:,;]', "VParm", 1, 0)
		m.Parser.AddTokenPattern("^(:(.+))", "VProp", 1, 0)

		* success, proceed to identify components, if a property with the same was identified
		IF m.Parser.GetTokens(m.SerializedANSI) AND m.Parser.Tokens(1).Type == "NProp" AND This.ICName == UPPER(m.Parser.Tokens(1).Value)

			* take care of the parameter section, leave the property value for last
			m.TokenIndex = 2
			DO WHILE m.TokenIndex <= m.Parser.Tokens.Count AND m.Parser.Tokens(m.TokenIndex).Type == "NParm"

				* the name of the parameter
				m.ParameterName = UPPER(m.Parser.Tokens(m.TokenIndex).Value)
				* now fetch its value(s)
				m.TokenIndex = m.TokenIndex + 1
				m.ParameterSet = .F.
				m.ParameterValue = ""

				DO WHILE m.TokenIndex <= m.Parser.Tokens.Count AND m.Parser.Tokens(m.TokenIndex).Type == "VParm"

					m.ParameterValue = m.ParameterValue + IIF(m.ParameterSet, ",", "") + m.Parser.Tokens(m.TokenIndex).Value
					m.ParameterSet = .T.
					m.TokenIndex = m.TokenIndex + 1

				ENDDO

				* a parameter was set?
				IF m.ParameterSet
					* try to instantiate an object, if known
					m.ParameterObject = This.iCalCreateObject("Parm", m.ParameterName, This.HostComponent)
					* a parameter class could be found?
					IF !ISNULL(m.ParameterObject)
						* set the value(s) and add the parameter to the property
						m.ParameterObject.Parse(m.ParameterValue)
						This.AddICParameter(m.ParameterObject)
						m.ParameterObject = .NULL.
					ENDIF
				ENDIF
			ENDDO

			* does the token collection end with a property value? then set it and signal success
			IF m.TokenIndex = m.Parser.Tokens.Count AND m.Parser.Tokens(m.TokenIndex).Type == "VProp"
				This.HValue.Parse(m.Parser.Tokens(m.TokenIndex).Value, This.Level)
				m.Parsed = .T.
			ENDIF
		ENDIF

		* get rid of any leftovers
		IF !m.Parsed
			This.Destroy()
		ENDIF

		RETURN m.Parsed

	ENDFUNC

ENDDEFINE

* general iCalendar property class definition

DEFINE CLASS _iCalParameter AS _iCalValueHandler

	Level = "PARAMETER"

	* generate the parameter in iCalendar format
	FUNCTION Serialize () AS String

		LOCAL Serialized AS String
		LOCAL SerializedValue AS String

		m.Serialized = ""
		m.SerializedValue = This.HValue.Serialize(This.Level)

		IF !EMPTY(NVL(m.SerializedValue, ""))
			m.Serialized = This.ICName + "=" + m.SerializedValue
		ENDIF

		RETURN m.Serialized
	ENDFUNC

	* read an iCalendar formatted parameter into a parameter object
	FUNCTION Parse (Serialized AS String)

		ASSERT VARTYPE(m.Serialized) == "C"

		This.HValue.Parse(m.Serialized, This.Level)

	ENDFUNC

ENDDEFINE

* a general class for iCalendar objects that hold values (that is, properties and parameters)

DEFINE CLASS _iCalValueHandler AS _iCalElement

	ADD OBJECT HValue AS _iCalValue

	Value_DataType = .NULL.
	Value_AlternativeDataTypes = .NULL.
	Value_IsList = .NULL.
	Value_IsUTC = .NULL.
	Value_IsComposite = .NULL.

	* initialize to default and intantiation values, and configuration
	FUNCTION Init (InitialValue AS AnyType)

		SAFETHIS

		IF !ISNULL(This.Value_DataType)
			This.HValue.Value.DataType = This.Value_DataType
			This.HValue.Value.OriginalDataType = This.Value_DataType
		ENDIF
		IF !ISNULL(This.Value_AlternativeDataTypes)
			This.HValue.Value.AlternativeDataTypes = This.Value_AlternativeDataTypes
		ENDIF
		IF !ISNULL(This.Value_IsList)
			This.HValue.Value.IsList = This.Value_IsList
		ENDIF
		IF !ISNULL(This.Value_IsUTC)
			This.HValue.Value.IsUTC = This.Value_IsUTC
		ENDIF
		IF !ISNULL(This.Value_IsComposite)
			This.HValue.Value.IsComposite = This.Value_IsComposite
		ENDIF

		IF PCOUNT() = 1
			This.SetValue(m.InitialValue)
		ELSE
			IF !ISNULL(This.DefaultValue)
				This.SetValue(This.DefaultValue)
			ENDIF
		ENDIF

	ENDFUNC

	FUNCTION Destroy

		This.HValue.UnsetValue()

	ENDFUNC

	* return the value (that may be an item in a list or composite value)
	FUNCTION GetValue (DataIndex AS Integer) AS AnyType

		ASSERT PCOUNT() = 0 OR VARTYPE(m.DataIndex) == "N"

		IF PCOUNT() = 1
			RETURN This.HValue.GetValue(m.DataIndex)
		ELSE
			RETURN This.HValue.GetValue()
		ENDIF

	ENDFUNC

	* how many values the property or parameter has?
	FUNCTION GetValueCount () AS Integer

		IF This.HValue.Value.IsList OR This.HValue.Value.IsComposite
			RETURN This.HValue.ValuesCount
		ELSE
			RETURN 1
		ENDIF

	ENDFUNC

	* set a value
	FUNCTION SetValue (NewValue AS AnyType)

		SAFETHIS

		ASSERT PCOUNT() = 1

		LOCAL ARRAY EnumerationTokens(1)
		LOCAL Accept AS Logical
		LOCAL ARRAY ExeStack(1)
		LOCAL Levels AS Integer

		m.Accept = .F.

		m.Levels = ASTACKINFO(m.ExeStack)
		* if the value is read-only, accept only from the object init; in any other case, receive the value
		IF !This.ReadOnly OR (m.Levels > 1 AND m.ExeStack(m.Levels - 1, 3) == LOWER(This.Class) + ".init")

			IF This.HValue.Value.DataType == "TEXT" AND !EMPTY(This.Enumeration) AND ALINES(m.EnumerationTokens, This.Enumeration, 1, ",") > 0

				IF ASCAN(m.EnumerationTokens, m.NewValue, -1, -1, 1, 1 + 2 + 4) != 0 OR (This.Extensions AND UPPER(LEFT(m.NewValue, 2)) == "X-")
					This.HValue.SetValue(m.NewValue)
					m.Accept = .T.
				ENDIF

			ELSE
				This.HValue.SetValue(m.NewValue)
				m.Accept = .T.
			ENDIF

		ENDIF

	ENDFUNC

	* clear the value
	FUNCTION UnsetValue (KeepDataType AS Logical)

		ASSERT VARTYPE(m.KeepDataType) == "L"

		This.HValue.UnsetValue(m.KeepDataType)

	ENDFUNC

ENDDEFINE

* a class to process values in properties and parameters

DEFINE CLASS _iCalValue AS _iCalBase

	* the value
	ADD OBJECT Value AS _iCalValueInfo

	* a list of values
	DIMENSION ValueList(1)
	ValuesCount = 0

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="value" type="property" display="Value"/>' + ;
						'<memberdata name="valuelist" type="property" display="ValueList"/>' + ;
						'<memberdata name="valuescount" type="property" display="ValuesCount"/>' + ;
						'<memberdata name="getvalue" type="method" display="GetValue"/>' + ;
						'<memberdata name="setvalue" type="method" display="SetValue"/>' + ;
						'<memberdata name="unsetvalue" type="method" display="UnsetValue"/>' + ;
						'<memberdata name="_parse" type="method" display="_Parse"/>' + ;
						'<memberdata name="_serialize" type="method" display="_Serialize"/>' + ;
						"</VFPData>"

	* remove the value, on exit
	FUNCTION Destroy

		This.UnsetValue()

	ENDFUNC

	* set the value
	PROCEDURE SetValue (NewValue AS AnyType)

		SAFETHIS

		ASSERT PCOUNT() = 1

		LOCAL DataValue AS AnyType

		This.Value.Data = .NULL.

		* for structured values, instantiate the objects if the value passed as a serialized version
		DO CASE
		CASE This.Value.DataType == "DURATION" AND VARTYPE(m.NewValue) == "C"
			m.DataValue = CREATEOBJECT("iCalTypeDURATION", m.NewValue)
		CASE This.Value.DataType == "PERIOD" AND VARTYPE(m.NewValue) == "C"
			m.DataValue = CREATEOBJECT("iCalTypePERIOD", m.NewValue)
		CASE This.Value.DataType == "RECUR" AND VARTYPE(m.NewValue) == "C"
			m.DataValue = CREATEOBJECT("iCalTypeRECUR", m.NewValue)
		OTHERWISE
			m.DataValue = m.NewValue
		ENDCASE

		* add the value to the list of values
		IF This.Value.IsList OR This.Value.IsComposite
			This.ValuesCount = This.ValuesCount + 1
			DIMENSION This.ValueList(This.ValuesCount)
			This.ValueList(This.ValuesCount) = m.DataValue
		ELSE
			* or set the single value
			This.Value.Data = m.DataValue
		ENDIF

		m.DataValue = .NULL.

	ENDPROC

	* get the value (for a list of values, an array may be returned if no index is given)
	FUNCTION GetValue (DataIndex AS Integer) AS AnyType
	
		SAFETHIS

		ASSERT PCOUNT() = 0 OR VARTYPE(m.DataIndex) == "N"

		IF This.Value.IsList OR This.Value.IsComposite
			IF PCOUNT() = 1
				IF This.ValuesCount != 0 AND BETWEEN(m.DataIndex, 1, This.ValuesCount)
					RETURN This.ValueList(m.DataIndex)
				ELSE
					RETURN .NULL.
				ENDIF
			ELSE
				RETURN @This.ValueList
			ENDIF
		ELSE
			RETURN This.Value.Data
		ENDIF

	ENDFUNC

	* clear the value (but hold the data type, in case it has changed)
	PROCEDURE UnsetValue (KeepDataType AS Logical)

		SAFETHIS

		ASSERT PCOUNT() = 0 OR VARTYPE(m.KeepDataType) == "L"
 
		LOCAL LoopIndex AS Integer

		IF This.Value.IsList OR This.Value.IsComposite
			FOR m.LoopIndex = This.ValuesCount TO 1 STEP -1
				This.ValueList(m.LoopIndex) = .NULL.
			ENDFOR
			DIMENSION This.ValueList(1)
			This.ValuesCount = 0
		ELSE
			This.Value.Data = .NULL.
		ENDIF

		IF !m.KeepDataType
			This.Value.DataType = This.Value.OriginalDataType
		ENDIF

	ENDPROC

	* serialize a value, or a list of values
	FUNCTION Serialize (Level AS String) AS String

		SAFETHIS

		ASSERT VARTYPE(m.Level) == "C"

		LOCAL SerializedList AS String
		LOCAL ListEntry AS String
		LOCAL ListSeparator AS String
		LOCAL LoopIndex AS Integer

		IF This.Value.IsList OR This.Value.IsComposite
			m.SerializedList = ""
			m.ListSeparator = IIF(This.Value.IsComposite, ";", ",")
			FOR m.LoopIndex = 1 TO This.ValuesCount
				m.ListEntry = This._Serialize(This.ValueList(m.LoopIndex), m.Level)
				IF !EMPTY(m.ListEntry)
					m.SerializedList = m.SerializedList + IIF(EMPTY(m.SerializedList), "" , m.ListSeparator) + m.ListEntry
				ENDIF
			ENDFOR
			RETURN m.SerializedList
		ELSE
			RETURN This._Serialize(This.Value.Data, m.Level)
		ENDIF

	ENDFUNC

	* serialize a single value (the level specifies if it is a parameter or property value)
	PROTECTED FUNCTION _Serialize (Value AS AnyType, Level AS String) AS String

		SAFETHIS

		LOCAL IsParameter AS Logical
		LOCAL IsProperty AS Logical
		LOCAL Encoded AS String

		STORE .F. TO m.IsParameter, m.IsProperty

		m.IsProperty = (m.Level == "PROPERTY")
		m.IsParameter = !m.IsProperty AND (m.Level == "PARAMETER")

		* the serialization depends on the data type and on the context (property or parameter)
		DO CASE
		CASE ISNULL(m.Value)
			RETURN ""

		CASE This.Value.DataType == "BINARY"
			RETURN STRCONV(m.Value, 13)

		CASE This.Value.DataType == "BOOLEAN"
			RETURN IIF(m.Value, "TRUE", "FALSE")

		CASE This.Value.DataType == "CAL-ADDRESS"
			IF m.IsParameter
				RETURN '"' + CHRTRAN(m.Value, '"', "") + '"'
			ELSE
				RETURN m.Value
			ENDIF

		CASE This.Value.DataType == "DATE"
			RETURN IIF(EMPTY(m.Value), "", DTOS(m.Value))

		CASE This.Value.DataType == "DATE-TIME"
			RETURN IIF(EMPTY(m.Value), "", TRANSFORM(TTOC(m.Value, 1), "@R 99999999T999999") + IIF(This.Value.IsUTC, "Z", ""))

		CASE This.Value.DataType == "DURATION"
			RETURN m.Value.Serialize(m.Level)

		CASE This.Value.DataType == "FLOAT"
			RETURN CHRTRAN(TRANSFORM(m.Value), SET("Point"), ".")

		CASE This.Value.DataType == "INTEGER"
			RETURN TRANSFORM(INT(m.Value))

		CASE This.Value.DataType == "PERIOD"
			RETURN m.Value.Serialize(m.Level)

		CASE This.Value.DataType == "RECUR"
			RETURN m.Value.Serialize(m.Level)

		CASE This.Value.DataType == "TEXT"
			IF m.IsProperty
				m.Encoded = STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(STRTRAN(m.Value, "\", "\\"), ";", "\;"), ",", "\,"), CRLF, "\n"), CHR(13), "\n"), CHR(10), "\n")
			ELSE
				m.Encoded = m.Value
				IF m.IsParameter
					IF RFC6868
						m.Encoded = STRTRAN(STRTRAN(m.Encoded, "^", "^^"), '"', "^'")
					ENDIF
					IF !(m.Encoded == CHRTRAN(m.Encoded, " ;:,", ""))
						m.Encoded = '"' + m.Encoded + '"'
					ENDIF
				ENDIF
			ENDIF
			RETURN m.Encoded

		CASE This.Value.DataType == "TIME"
			* special case: time uses the time portion in a Datetime value
			RETURN IIF(EMPTY(m.Value), "", SUBSTR(TTOC(m.Value, 1), 9) + IIF(This.Value.IsUTC, "Z", ""))

		CASE This.Value.DataType == "URI"
			IF m.IsParameter
				RETURN '"' + CHRTRAN(m.Value, '"', "%22") + '"'
			ELSE
				RETURN m.Value
			ENDIF

		CASE This.Value.DataType == "UTC-OFFSET"
			RETURN IIF(m.Value >= 0, "+", "-") + TRANSFORM(INT(INT(m.Value) / 60), "@L 99") + TRANSFORM(INT(INT(m.Value) % 60), "@L 99") + ;
						IIF(m.Value = INT(m.Value), "", TRANSFORM(INT((m.Value - INT(m.Value)) * 60), "@L 99"))
		
		OTHERWISE
			RETURN This.Value.Serialize(m.Value, m.Level)
		ENDCASE

	ENDFUNC

	* read the seralized version of a value, in iCalendar format, and ingest it into a VFP value (scalar or structured)
	FUNCTION Parse (Serialized AS String, Level AS String)

		SAFETHIS

		ASSERT VARTYPE(m.Serialized) + VARTYPE(m.Level) == "CC"

		LOCAL ListSeparator AS Character
		LOCAL ListEntry AS String
		LOCAL CharIndex AS Integer
		LOCAL Parsed AS String
		LOCAL Parser AS Tokenizer
		LOCAL Token AS Token
		LOCAL ARRAY ListEntries(1)

		This.UnsetValue(.T.)

		* read a single value
		IF !This.Value.IsList AND !This.Value.IsComposite
			This.SetValue(This._Parse(m.Serialized, m.Level))
		ELSE
			* or a comma or semicollon separated list of values
			m.ListSeparator = IIF(This.Value.IsComposite, ";", ",")
			* for text values, go through the value and look for unescaped separators
			IF This.Value.DataType == "TEXT"
				m.Parsed = m.Serialized
				m.CharIndex = 1
				DO WHILE !EMPTY(m.Parsed)
					DO CASE
					CASE SUBSTR(m.Parsed, m.CharIndex, 1) == "\"
						m.CharIndex = m.CharIndex + 2
					* a item in the list was found, parse it
					CASE SUBSTR(m.Parsed, m.CharIndex, 1) == m.ListSeparator
						This.SetValue(This._Parse(LEFT(m.Parsed, m.CharIndex - 1), m.Level))
						m.Parsed = SUBSTR(m.Parsed, m.CharIndex + 1)
						m.CharIndex = 1
					* reached the end of the list, parse and finish
					CASE m.CharIndex = LEN(m.Parsed)
						This.SetValue(This._Parse(m.Parsed, m.Level))
						m.Parsed = ""
					OTHERWISE
						m.CharIndex = m.CharIndex + 1
					ENDCASE
				ENDDO
			ELSE
			* for all other types, a regular expression will be able to extract the items in the list
				m.Parser = CREATEOBJECT("Tokenizer")
				m.Parser.AddTokenPattern(CHRTRAN('^((("[^"]+")|([^,]+)),?)', ',', m.ListSeparator), "Entry", 1, 0)
				IF m.Parser.GetTokens(m.Serialized)
					FOR EACH m.Token IN m.Parser.Tokens
						This.SetValue(This._Parse(m.Token.Value, m.Level))
					ENDFOR
				ENDIF
			ENDIF
		ENDIF

	ENDFUNC	

	* read a single value from its iCalendar expression into a VFP value / object
	PROTECTED FUNCTION _Parse (Serialized AS String, Level AS String) AS AnyType

		SAFETHIS

		LOCAL Parsed AS AnyValue
		LOCAL CharIndex AS Integer
		LOCAL ParsedChar AS Character
		LOCAL IsParameter AS Logical
		LOCAL IsProperty AS Logical
		LOCAL SerializedANSI AS String

		m.IsProperty = (m.Level == "PROPERTY")
		m.IsParameter = !m.IsProperty AND (m.Level == "PARAMETER")

		DO CASE
		CASE EMPTY(m.Serialized) OR ISNULL(m.Serialized)
			RETURN .NULL.

		CASE This.Value.DataType == "BINARY"
			RETURN STRCONV(m.Serialized, 14)

		CASE This.Value.DataType == "BOOLEAN"
			RETURN UPPER(m.Serialized) == "TRUE"

		CASE This.Value.DataType == "CAL-ADDRESS"
			RETURN CHRTRAN(m.Serialized, '"', "")

		CASE This.Value.DataType == "DATE"
			RETURN EVL(EVALUATE("{^" + TRANSFORM(m.Serialized, "@R 9999-99-99") + "}"), .NULL.)

		CASE This.Value.DataType == "DATE-TIME"
			m.Parsed = EVALUATE("{^" + TRANSFORM(CHRTRAN(m.Serialized, "T", ""), "@R 9999-99-99 99:99:99") + "}")
			IF !EMPTY(m.Parsed)
				This.Value.IsUTC = RIGHT(m.Serialized, 1) == "Z"
			ENDIF
			RETURN EVL(m.Parsed, .NULL.)

		CASE This.Value.DataType == "DURATION"
			m.Parsed = CREATEOBJECT("iCalTypeDURATION", m.Serialized)
			IF VARTYPE(m.Parsed) == "O" AND !ISNULL(m.Parsed)
				RETURN m.Parsed
			ELSE
				RETURN .NULL.
			ENDIF

		CASE This.Value.DataType == "FLOAT"
			RETURN VAL(CHRTRAN(m.Serialized, ".", SET("Point")))

		CASE This.Value.DataType == "INTEGER"
			RETURN INT(VAL(m.Serialized))

		CASE This.Value.DataType == "PERIOD"
			m.Parsed = CREATEOBJECT("iCalTypePERIOD", m.Serialized)
			IF VARTYPE(m.Parsed) == "O" AND !ISNULL(m.Parsed)
				RETURN m.Parsed
			ELSE
				RETURN .NULL.
			ENDIF

		CASE This.Value.DataType == "RECUR"
			m.Parsed = CREATEOBJECT("iCalTypeRECUR", m.Serialized)
			IF VARTYPE(m.Parsed) == "O" AND !ISNULL(m.Parsed)
				RETURN m.Parsed
			ELSE
				RETURN .NULL.
			ENDIF

		CASE This.Value.DataType == "TEXT"
			m.Parsed = ""
			m.CharIndex = 1

			* check if valid UTF-8
			IF m.Serialized == STRCONV(STRCONV(m.Serialized, 11), 9)
				m.SerializedANSI = STRCONV(STRCONV(m.Serialized, 11), 2)
			ELSE
			* if not, assume already in ANSI
				m.SerializedANSI = m.Serialized
			ENDIF

			* unescape the value
			DO WHILE m.CharIndex <= LEN(m.SerializedANSI)
				m.ParsedChar = SUBSTR(m.SerializedANSI, m.CharIndex, 1)
				DO CASE
				* unescape parameter values
				CASE m.ParsedChar == "^" AND m.IsParameter AND RFC6868
					m.ParsedChar = SUBSTR(m.SerializedANSI, m.CharIndex + 1, 1)
					DO CASE
					CASE m.ParsedChar == "^"
						m.Parsed = m.Parsed + "^"
					CASE m.ParsedChar == "'"
						m.Parsed = m.Parsed + '"'
					CASE m.ParsedChar == "n"
						m.Parsed = m.Parsed + CRLF
					OTHERWISE
						m.Parsed = m.Parsed + "^" + m.ParsedChar
					ENDCASE
					m.CharIndex = m.CharIndex + 2
				* unescape property values
				CASE m.ParsedChar == "\" AND m.IsProperty
					m.ParsedChar = SUBSTR(m.SerializedANSI, m.CharIndex + 1, 1)
					IF m.ParsedChar == "n" OR m.ParsedChar == "N"
						m.Parsed = m.Parsed + CRLF
					ELSE
						m.Parsed = m.Parsed + m.ParsedChar
					ENDIF
					m.CharIndex = m.CharIndex + 2
				OTHERWISE
					m.Parsed = m.Parsed + m.ParsedChar
					m.CharIndex = m.CharIndex + 1
				ENDCASE
			ENDDO

			* remove the surrounding quotes in parameter values
			IF m.IsParameter AND LEFT(m.Parsed, 1) == '"' AND RIGHT(m.Parsed, 1) == '"'
				m.Parsed = SUBSTR(m.Parsed, 2, LEN(m.Parsed) - 2)
			ENDIF

			RETURN EVL(m.Parsed, .NULL.)

		CASE This.Value.DataType == "TIME"
			* special case: time is stored as a Datetime value 
			m.Parsed = EVALUATE("{^1980-01-01 " + TRANSFORM(m.Serialized, "@R 99:99:99") + "}")
			IF !EMPTY(m.Parsed)
				This.Value.IsUTC = RIGHT(m.Serialized, 1) == "Z"
			ENDIF
			RETURN EVL(m.Parsed, .NULL.)

		CASE This.Value.DataType == "URI"
			RETURN CHRTRAN(m.Serialized, '"', "")

		CASE This.Value.DataType == "UTC-OFFSET"
			RETURN VAL(LEFT(m.Serialized, 3)) * 60 + VAL(SUBSTR(m.Serialized, 4, 2)) + IIF(LEN(m.Serialized) = 7, VAL(SUBSTR(m.Serialized, 6, 2) / 60), 0)

		OTHERWISE
			RETURN This.Value.Parse(m.Serialized, m.Level)
		ENDCASE

	ENDFUNC

ENDDEFINE

* general class to store information about a value

DEFINE CLASS _iCalValueInfo AS _iCalBase

	Data = .NULL.
	DataType = "TEXT"
	AlternativeDataTypes = ""
	OriginalDataType = "TEXT"
	IsList = .F.
	IsUTC = .F.
	IsComposite = .F.

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="alternativedatatypes" type="property" display="AlternativeDataTypes"/>' + ;
						'<memberdata name="data" type="property" display="Data"/>' + ;
						'<memberdata name="datatype" type="property" display="DataType"/>' + ;
						'<memberdata name="iscomposite" type="property" display="IsComposite"/>' + ;
						'<memberdata name="islist" type="property" display="IsList"/>' + ;
						'<memberdata name="isutc" type="property" display="IsUTC"/>' + ;
						"</VFPData>"

	PROCEDURE Destroy

		SAFETHIS

		IF VARTYPE(This.Data) == "O"
			This.Data = .NULL.
		ENDIF

	ENDPROC

	* just placeholders
	FUNCTION Serialize (Value AS AnyType, Level AS String) AS String
		RETURN ""
	ENDFUNC

	FUNCTION Parse (Serialized AS AnyType, Level AS String) AS String
		RETURN .NULL.
	ENDFUNC

ENDDEFINE

* the foundation class for iCalendar elements (components, properties, and parameters)

DEFINE CLASS _iCalElement AS _iCalBase

	* the iCalendar name
	ICName = ""
	* future use (xml serialization)
	xICName = ""

	* the devault value (if any)
	DefaultValue = .NULL.
	ReadOnly = .F.

	* a comma separated list of possible values
	Enumeration = ""
	* admit X- extensions
	Extensions = .T.

	* the name of the type of the host component
	HostComponent = .NULL.

	* a comma separated list of component types that may be used to identify component-dependent alternative classes
	AlternativeClasses = ""
	Level = ""

	_MemberData =	"<VFPData>" + ;
						'<memberdata name="alternativeclasses" type="property" display="AlternativeClasses"/>' + ;
						'<memberdata name="defaultvalue" type="property" display="DefaultValue"/>' + ;
						'<memberdata name="enumeration" type="property" display="Enumeration"/>' + ;
						'<memberdata name="extensions" type="property" display="Extensions"/>' + ;
						'<memberdata name="hostcomponent" type="property" display="HostComponent"/>' + ;
						'<memberdata name="icname" type="property" display="ICName"/>' + ;
						'<memberdata name="level" type="property" display="Level"/>' + ;
						'<memberdata name="readonly" type="property" display="ReadOnly"/>' + ;
						'<memberdata name="xicname" type="property" display="xICName"/>' + ;
						'<memberdata name="icalcreateobject" type="method" display="iCalCreateObject"/>' + ;
						'<memberdata name="getvalue" type="method" display="GetValue"/>' + ;
						'<memberdata name="getvaluecount" type="method" display="GetValueCount"/>' + ;
						'<memberdata name="reset" type="method" display="Reset"/>' + ;
						'<memberdata name="setvalue" type="method" display="SetValue"/>' + ;
						'<memberdata name="unsetvalue" type="method" display="UnsetValue"/>' + ;
						"</VFPData>"

	FUNCTION GetValue (DataIndex AS Integer) AS AnyType
	ENDFUNC

	FUNCTION GetValueCount () AS Integer
		RETURN 0
	ENDFUNC

	FUNCTION SetValue (NewValue AS AnyType)
	ENDFUNC

	FUNCTION UnsetValue (KeepDataType AS Logical)
	ENDFUNC

	FUNCTION Reset ()
		This.Destroy()
	ENDFUNC

	* a general method to create an iCalendar object of any type ("Comp", "Prop", or "Parm")
	FUNCTION iCalCreateObject (ClassType AS String, Name AS String, Host AS String) AS Object

		SAFETHIS

		ASSERT VARTYPE(m.ClassType) + VARTYPE(m.Name) == "CC" AND VARTYPE(m.Host) $ "CX"

		LOCAL SafeName AS String
		LOCAL Instance AS Object

		This.HostComponent = CHRTRAN(NVL(m.Host, ""), "-", "_")
		m.SafeName = "iCal" + m.ClassType + CHRTRAN(m.Name, "-", "_")

		TRY
			m.Instance = CREATEOBJECT(m.SafeName)
			IF !EMPTY(m.Instance.AlternativeClasses) AND ATC("," + This.HostComponent + ",", "," + m.Instance.AlternativeClasses + ",") != 0
				m.SafeName = m.SafeName + "_" + This.HostComponent
				m.Instance = .NULL.
				m.Instance = CREATEOBJECT(m.SafeName)
			ENDIF
		CATCH
			m.Instance = .NULL.
		ENDTRY

		RETURN m.Instance

	ENDFUNC

	* get an iCalendar object from an arrayed list of sub-elements (parameters in a property, properties in a component, ...),
	*  indexed by name, position, or name/position
	FUNCTION _GetIC (Id AS StringOrInteger, SubIndex AS Integer, SubArray AS Array, SubCount AS Integer) AS _iCalBase

		LOCAL LoopIndex AS Integer
		LOCAL ICNameIndex AS Integer
		LOCAL UICName AS String

		IF VARTYPE(m.Id) == "N"
			RETURN IIF(BETWEEN(m.Id, 1, m.SubCount), m.SubArray(m.Id), .NULL.)
		ENDIF

		m.ICNameIndex = 1
		m.UICName = UPPER(m.Id)

		FOR m.LoopIndex = 1 TO m.SubCount
			IF !ISNULL(m.SubArray(m.LoopIndex)) AND m.SubArray(m.LoopIndex).ICName == m.UICName
				IF m.ICNameIndex = m.SubIndex
					RETURN m.SubArray(m.LoopIndex)
				ENDIF
				m.ICNameIndex = m.ICNameIndex + 1
			ENDIF
		ENDFOR

		RETURN .NULL.
	ENDFUNC

	* return the number of iCalendar objects in an arrayed list of sub-elements,
	*  of all types or just from a specific type 
	FUNCTION _GetICCount (Id AS String, SubArray AS Array, SubCount AS Integer) AS Integer

		LOCAL LoopIndex AS Integer
		LOCAL UICName AS String
		LOCAL ICCount AS Integer

		IF EMPTY(m.Id)
			RETURN m.SubCount
		ENDIF

		m.UICName = UPPER(m.Id)
		m.ICCount = 0

		FOR m.LoopIndex = 1 TO m.SubCount
			IF !ISNULL(m.SubArray(m.LoopIndex)) AND m.SubArray(m.LoopIndex).ICName == m.UICName
				m.ICCount = m.ICCount + 1
			ENDIF
		ENDFOR

		RETURN m.ICCount

	ENDFUNC

ENDDEFINE

* the general base class upon which the iCalendar classes derive

DEFINE CLASS _iCalBase AS Custom

	_MemberData =	'<VFPData>' + ;
							'<memberdata name="parse" type="method" display="Parse"/>' + ;
							'<memberdata name="serialize" type="method" display="Serialize"/>' + ;
						'</VFPData>'

	FUNCTION Serialize () AS String
		RETURN ""
	ENDFUNC

	FUNCTION Parse (Serialized AS String) AS Logical
		RETURN .NULL.
	ENDFUNC

ENDDEFINE


