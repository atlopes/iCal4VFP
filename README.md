# iCal4VFP

A VFP library set to work with iCalendar objects.

## Overview

- Adheres to [RFC5545](https://tools.ietf.org/html/rfc5545 "RFC5545") specifications.
- Serializers and parsers for all Components, Properties, and Parameters.
- Includes structured iCalendar data types (RECUR, DURATION, and PERIOD)
- Integrates a Recurrence Rule processor that runs all RFC5545 examples.
- Reads iCalendar data from .ics files, or from strings.
- Allows for extensions at all levels (Components, Properties, Parameters, and Types).
- Processes timezone information, and translates between UTC and local time.
- Scalar types maps to VFP data types (Date, Datetime, String, Logical, .

## Classes

For every Component, Property, Parameter, and Structured Type specified by RFC5545 there is a corresponding class definition (for instance, `iCalCompVEVENT`, `iCalPropDTSTART`, `iCalParmTZID`, and `iCalTypeDURATION`). When the name of the iCalendar element includes an hyphen, the corresponding name in the VFP class set includes an underscore symbol, instead (for instance, `iCalPropLAST_MODIFIED`).

Values can be assigned to parameters and properties; parameters can be added to properties: properties and components can be added to components.

Values are also treated as objects, with specific methods to Set and Get them.

### Base classes

	_iCalBase
		_iCalElement
			_iCalComponent
			_iCalValueHandler
				_iCalProperty
				_iCalParameter
		_iCalValue
		_iCalValueInfo
		_iCalType
			iCalTypePERIOD
			iCalTypeDURATION
			iCalTypeRECUR

### ICSProcessor class

The ICSProcessor class is intended to process an iCalendar formatted source: an `.ics` file, or a iCalendar formatted string.

#### Methods
| Name  | Type  | Obs |
| ------------ | ------------ | ------------ |
| Read  | O  | Reads an iCalendar formatted string, and returns an `iCalendar` object  |
| ReadFile  | O  | Reads an `.ics` file, and returns an `iCalendar` object |
| ICSToCursor | N | Create a simplified flat cursor over some iCalendar source (an `iCalendar` object, an `.ics` file, or a formatted string) |

### Components classes

In iCalendar, components are container objects that integrate properties and, eventually, other components. They can be initialized and defined as a sequence of `CREATEOBJECT()` function calls, or trough a method of the ICSProcessor class.

	iCalendar
	iCalCompVEVENT
	iCalCompVTODO
	iCalCompVJOURNAL
	iCalCompVFREEBUSY
	iCalCompVTIMEZONE
	iCalCompVALARM
	iCalCompSTANDARD
	iCalCompDAYLIGHT


### Properties classes

	iCalPropACTION
	iCalPropATTACH
	iCalPropATTENDEE
	iCalPropCALSCALE
	iCalPropCATEGORIES
	iCalPropCLASS
	iCalPropCOMMENT
	iCalPropCOMPLETED
	iCalPropCONTACT
	iCalPropCREATED
	iCalPropDESCRIPTION
	iCalPropDTEND
	iCalPropDTSTAMP
	iCalPropDTSTART
	iCalPropDUE
	iCalPropDURATION
	iCalPropEXDATE
	iCalPropFREEBUSY
	iCalPropGEO
	iCalPropLAST_MODIFIED
	iCalPropLOCATION
	iCalPropMETHOD
	iCalPropORGANIZER
	iCalPropPERCENTCOMPLETE
	iCalPropPRIORITY
	iCalPropPRODID
	iCalPropRDATE
	iCalPropRECURRENCE_ID
	iCalPropRELATED_TO
	iCalPropREPEAT
	iCalPropREQUEST_STATUS
	iCalPropRESOURCES
	iCalPropRRULE
	iCalPropSEQUENCE
	iCalPropSTATUS
	iCalPropSUMMARY
	iCalPropTRANSP
	iCalPropTRIGGER
	iCalPropTZID
	iCalPropTZNAME
	iCalPropTZOFFSETFROM
	iCalPropTZOFFSETTO
	iCalPropTZURL
	iCalPropUID
	iCalPropURL
	iCalPropVERSION

### Parameters classes

	iCalParmALTREP
	iCalParmCN
	iCalParmCUTYPE
	iCalParmDELEGATED
	iCalParmDIR
	iCalParmENCODING
	iCalParmFMTTYPE
	iCalParmFBTYPE
	iCalParmLANGUAGE
	iCalParmMEMBER
	iCalParmPARTSTAT
	iCalParmRANGE
	iCalParmRELATED
	iCalParmRELTYPE
	iCalParmROLE
	iCalParmRSVP
	iCalParmSENT_BY
	iCalParmTZID
	iCalParmVALUE
