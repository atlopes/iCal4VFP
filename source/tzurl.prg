*!*	+--------------------------------------------------------------------+
*!*	|                                                                    |
*!*	|    iCal4VFP                                                        |
*!*	|                                                                    |
*!*	+--------------------------------------------------------------------+

*!*	a VFP class to use the TzURL timezone data at http://www.tzurl.org

* dependencies
IF _VFP.StartMode = 0
	SET PATH TO (JUSTPATH(SYS(16))) ADDITIVE
ENDIF
DO "icalloader.prg"

* install itself
IF !SYS(16) $ SET("Procedure")
	SET PROCEDURE TO (SYS(16)) ADDITIVE
ENDIF

DEFINE CLASS TzURL AS _iCalBase

	ADD OBJECT Timezones AS Collection

	HIDDEN Cache
	Cache = 30

	_memberdata = '<VFPData>' + ;
						'<memberdata name="timezones" type="property" display="Timezones"/>' + ;
						'<memberdata name="full" type="method" display="Full"/>' + ;
						'<memberdata name="minimal" type="method" display="Minimal"/>' + ;
						'<memberdata name="setcache" type="method" display="SetCache"/>' + ;
						'</VFPData>'

	FUNCTION Init

		LOCAL ARRAY TZIDs(1)
		LOCAL TZID AS String
		LOCAL TZURL AS String
		LOCAL LoopIndex AS Integer

		TEXT TO m.TZURL NOSHOW PRETEXT 1 + 2
		Africa/Abidjan
		Africa/Accra
		Africa/Addis_Ababa
		Africa/Algiers
		Africa/Asmara
		Africa/Bamako
		Africa/Bangui
		Africa/Banjul
		Africa/Bissau
		Africa/Blantyre
		Africa/Brazzaville
		Africa/Bujumbura
		Africa/Cairo
		Africa/Casablanca
		Africa/Ceuta
		Africa/Conakry
		Africa/Dakar
		Africa/Dar_es_Salaam
		Africa/Djibouti
		Africa/Douala
		Africa/El_Aaiun
		Africa/Freetown
		Africa/Gaborone
		Africa/Harare
		Africa/Johannesburg
		Africa/Juba
		Africa/Kampala
		Africa/Khartoum
		Africa/Kigali
		Africa/Kinshasa
		Africa/Lagos
		Africa/Libreville
		Africa/Lome
		Africa/Luanda
		Africa/Lubumbashi
		Africa/Lusaka
		Africa/Malabo
		Africa/Maputo
		Africa/Maseru
		Africa/Mbabane
		Africa/Mogadishu
		Africa/Monrovia
		Africa/Nairobi
		Africa/Ndjamena
		Africa/Niamey
		Africa/Nouakchott
		Africa/Ouagadougou
		Africa/Porto-Novo
		Africa/Sao_Tome
		Africa/Tripoli
		Africa/Tunis
		Africa/Windhoek
		America/Adak
		America/Anchorage
		America/Anguilla
		America/Antigua
		America/Araguaina
		America/Argentina/Buenos_Aires
		America/Argentina/Catamarca
		America/Argentina/Cordoba
		America/Argentina/Jujuy
		America/Argentina/La_Rioja
		America/Argentina/Mendoza
		America/Argentina/Rio_Gallegos
		America/Argentina/Salta
		America/Argentina/San_Juan
		America/Argentina/San_Luis
		America/Argentina/Tucuman
		America/Argentina/Ushuaia
		America/Aruba
		America/Asuncion
		America/Atikokan
		America/Bahia
		America/Bahia_Banderas
		America/Barbados
		America/Belem
		America/Belize
		America/Blanc-Sablon
		America/Boa_Vista
		America/Bogota
		America/Boise
		America/Cambridge_Bay
		America/Campo_Grande
		America/Cancun
		America/Caracas
		America/Cayenne
		America/Cayman
		America/Chicago
		America/Chihuahua
		America/Costa_Rica
		America/Creston
		America/Cuiaba
		America/Curacao
		America/Danmarkshavn
		America/Dawson
		America/Dawson_Creek
		America/Denver
		America/Detroit
		America/Dominica
		America/Edmonton
		America/Eirunepe
		America/El_Salvador
		America/Fort_Nelson
		America/Fortaleza
		America/Glace_Bay
		America/Godthab
		America/Goose_Bay
		America/Grand_Turk
		America/Grenada
		America/Guadeloupe
		America/Guatemala
		America/Guayaquil
		America/Guyana
		America/Halifax
		America/Havana
		America/Hermosillo
		America/Indiana/Indianapolis
		America/Indiana/Knox
		America/Indiana/Marengo
		America/Indiana/Petersburg
		America/Indiana/Tell_City
		America/Indiana/Vevay
		America/Indiana/Vincennes
		America/Indiana/Winamac
		America/Inuvik
		America/Iqaluit
		America/Jamaica
		America/Juneau
		America/Kentucky/Louisville
		America/Kentucky/Monticello
		America/Kralendijk
		America/La_Paz
		America/Lima
		America/Los_Angeles
		America/Lower_Princes
		America/Maceio
		America/Managua
		America/Manaus
		America/Marigot
		America/Martinique
		America/Matamoros
		America/Mazatlan
		America/Menominee
		America/Merida
		America/Metlakatla
		America/Mexico_City
		America/Miquelon
		America/Moncton
		America/Monterrey
		America/Montevideo
		America/Montserrat
		America/Nassau
		America/New_York
		America/Nipigon
		America/Nome
		America/Noronha
		America/North_Dakota/Beulah
		America/North_Dakota/Center
		America/North_Dakota/New_Salem
		America/Ojinaga
		America/Panama
		America/Pangnirtung
		America/Paramaribo
		America/Phoenix
		America/Port-au-Prince
		America/Port_of_Spain
		America/Porto_Velho
		America/Puerto_Rico
		America/Punta_Arenas
		America/Rainy_River
		America/Rankin_Inlet
		America/Recife
		America/Regina
		America/Resolute
		America/Rio_Branco
		America/Santarem
		America/Santiago
		America/Santo_Domingo
		America/Sao_Paulo
		America/Scoresbysund
		America/Sitka
		America/St_Barthelemy
		America/St_Johns
		America/St_Kitts
		America/St_Lucia
		America/St_Thomas
		America/St_Vincent
		America/Swift_Current
		America/Tegucigalpa
		America/Thule
		America/Thunder_Bay
		America/Tijuana
		America/Toronto
		America/Tortola
		America/Vancouver
		America/Whitehorse
		America/Winnipeg
		America/Yakutat
		America/Yellowknife
		Antarctica/Casey
		Antarctica/Davis
		Antarctica/DumontDUrville
		Antarctica/Macquarie
		Antarctica/Mawson
		Antarctica/McMurdo
		Antarctica/Palmer
		Antarctica/Rothera
		Antarctica/Syowa
		Antarctica/Troll
		Antarctica/Vostok
		Arctic/Longyearbyen
		Asia/Aden
		Asia/Almaty
		Asia/Amman
		Asia/Anadyr
		Asia/Aqtau
		Asia/Aqtobe
		Asia/Ashgabat
		Asia/Atyrau
		Asia/Baghdad
		Asia/Bahrain
		Asia/Baku
		Asia/Bangkok
		Asia/Barnaul
		Asia/Beirut
		Asia/Bishkek
		Asia/Brunei
		Asia/Chita
		Asia/Choibalsan
		Asia/Colombo
		Asia/Damascus
		Asia/Dhaka
		Asia/Dili
		Asia/Dubai
		Asia/Dushanbe
		Asia/Famagusta
		Asia/Gaza
		Asia/Hebron
		Asia/Ho_Chi_Minh
		Asia/Hong_Kong
		Asia/Hovd
		Asia/Irkutsk
		Asia/Istanbul
		Asia/Jakarta
		Asia/Jayapura
		Asia/Jerusalem
		Asia/Kabul
		Asia/Kamchatka
		Asia/Karachi
		Asia/Kathmandu
		Asia/Khandyga
		Asia/Kolkata
		Asia/Krasnoyarsk
		Asia/Kuala_Lumpur
		Asia/Kuching
		Asia/Kuwait
		Asia/Macau
		Asia/Magadan
		Asia/Makassar
		Asia/Manila
		Asia/Muscat
		Asia/Nicosia
		Asia/Novokuznetsk
		Asia/Novosibirsk
		Asia/Omsk
		Asia/Oral
		Asia/Phnom_Penh
		Asia/Pontianak
		Asia/Pyongyang
		Asia/Qatar
		Asia/Qyzylorda
		Asia/Riyadh
		Asia/Sakhalin
		Asia/Samarkand
		Asia/Seoul
		Asia/Shanghai
		Asia/Singapore
		Asia/Srednekolymsk
		Asia/Taipei
		Asia/Tashkent
		Asia/Tbilisi
		Asia/Tehran
		Asia/Thimphu
		Asia/Tokyo
		Asia/Tomsk
		Asia/Ulaanbaatar
		Asia/Urumqi
		Asia/Ust-Nera
		Asia/Vientiane
		Asia/Vladivostok
		Asia/Yakutsk
		Asia/Yangon
		Asia/Yekaterinburg
		Asia/Yerevan
		Atlantic/Azores
		Atlantic/Bermuda
		Atlantic/Canary
		Atlantic/Cape_Verde
		Atlantic/Faroe
		Atlantic/Madeira
		Atlantic/Reykjavik
		Atlantic/South_Georgia
		Atlantic/St_Helena
		Atlantic/Stanley
		Australia/Adelaide
		Australia/Brisbane
		Australia/Broken_Hill
		Australia/Currie
		Australia/Darwin
		Australia/Eucla
		Australia/Hobart
		Australia/Lindeman
		Australia/Lord_Howe
		Australia/Melbourne
		Australia/Perth
		Australia/Sydney
		Etc/GMT+0
		Etc/GMT+1
		Etc/GMT+10
		Etc/GMT+11
		Etc/GMT+12
		Etc/GMT+2
		Etc/GMT+3
		Etc/GMT+4
		Etc/GMT+5
		Etc/GMT+6
		Etc/GMT+7
		Etc/GMT+8
		Etc/GMT+9
		Etc/GMT-0
		Etc/GMT-1
		Etc/GMT-10
		Etc/GMT-11
		Etc/GMT-12
		Etc/GMT-13
		Etc/GMT-14
		Etc/GMT-2
		Etc/GMT-3
		Etc/GMT-4
		Etc/GMT-5
		Etc/GMT-6
		Etc/GMT-7
		Etc/GMT-8
		Etc/GMT-9
		Etc/GMT
		Etc/GMT0
		Etc/Greenwich
		Etc/UCT
		Etc/UTC
		Etc/Universal
		Etc/Zulu
		Europe/Amsterdam
		Europe/Andorra
		Europe/Astrakhan
		Europe/Athens
		Europe/Belgrade
		Europe/Berlin
		Europe/Bratislava
		Europe/Brussels
		Europe/Bucharest
		Europe/Budapest
		Europe/Busingen
		Europe/Chisinau
		Europe/Copenhagen
		Europe/Dublin
		Europe/Gibraltar
		Europe/Guernsey
		Europe/Helsinki
		Europe/Isle_of_Man
		Europe/Istanbul
		Europe/Jersey
		Europe/Kaliningrad
		Europe/Kiev
		Europe/Kirov
		Europe/Lisbon
		Europe/Ljubljana
		Europe/London
		Europe/Luxembourg
		Europe/Madrid
		Europe/Malta
		Europe/Mariehamn
		Europe/Minsk
		Europe/Monaco
		Europe/Moscow
		Europe/Nicosia
		Europe/Oslo
		Europe/Paris
		Europe/Podgorica
		Europe/Prague
		Europe/Riga
		Europe/Rome
		Europe/Samara
		Europe/San_Marino
		Europe/Sarajevo
		Europe/Saratov
		Europe/Simferopol
		Europe/Skopje
		Europe/Sofia
		Europe/Stockholm
		Europe/Tallinn
		Europe/Tirane
		Europe/Ulyanovsk
		Europe/Uzhgorod
		Europe/Vaduz
		Europe/Vatican
		Europe/Vienna
		Europe/Vilnius
		Europe/Volgograd
		Europe/Warsaw
		Europe/Zagreb
		Europe/Zaporozhye
		Europe/Zurich
		Indian/Antananarivo
		Indian/Chagos
		Indian/Christmas
		Indian/Cocos
		Indian/Comoro
		Indian/Kerguelen
		Indian/Mahe
		Indian/Maldives
		Indian/Mauritius
		Indian/Mayotte
		Indian/Reunion
		Pacific/Apia
		Pacific/Auckland
		Pacific/Bougainville
		Pacific/Chatham
		Pacific/Chuuk
		Pacific/Easter
		Pacific/Efate
		Pacific/Enderbury
		Pacific/Fakaofo
		Pacific/Fiji
		Pacific/Funafuti
		Pacific/Galapagos
		Pacific/Gambier
		Pacific/Guadalcanal
		Pacific/Guam
		Pacific/Honolulu
		Pacific/Kiritimati
		Pacific/Kosrae
		Pacific/Kwajalein
		Pacific/Majuro
		Pacific/Marquesas
		Pacific/Midway
		Pacific/Nauru
		Pacific/Niue
		Pacific/Norfolk
		Pacific/Noumea
		Pacific/Pago_Pago
		Pacific/Palau
		Pacific/Pitcairn
		Pacific/Pohnpei
		Pacific/Port_Moresby
		Pacific/Rarotonga
		Pacific/Saipan
		Pacific/Tahiti
		Pacific/Tarawa
		Pacific/Tongatapu
		Pacific/Wake
		Pacific/Wallis
		ENDTEXT

		FOR m.LoopIndex = 1 TO ALINES(m.TZIDs, m.TZURL)
			m.TZID = m.TZIDs(m.LoopIndex)
			This.Timezones.Add(m.TZID, m.TZID)
		ENDFOR
	ENDFUNC

	* reads the complete timezone definition from TZURL.org, or from cache
	FUNCTION Full (TzID AS String, StoredTz AS String) AS iCalCompVTIMEZONE

		ASSERT VARTYPE(m.TzID) == "C" AND (PCOUNT() = 1 OR VARTYPE(m.StoredTz) == "C")

		LOCAL Timezone AS iCalCompVTIMEZONE
		LOCAL ICS AS ICSProcessor
		LOCAL iCal AS iCalendar
		LOCAL LocalTz AS String
		LOCAL SafetySetting AS String
		LOCAL ARRAY IsFile[1]

		* load the timezone definition into an iCalendar
		m.ICS = CREATEOBJECT("ICSProcessor")
		m.iCal = .NULL.

		* before fetching the Timezone definition from TZURL, check if there is a fresh copy in local cache
		m.LocalTz = ADDBS(SYS(2023)) + "iCal4VFP\" + CHRTRAN(m.TzID, "/+", "__") + ".ics"
		IF !DIRECTORY(JUSTPATH(m.LocalTz))
			MKDIR (JUSTPATH(m.LocalTz))
		ENDIF

		* Cache = 0 : read always from TzURL
		* Cache < 0 : read always from cache, if available
		* Cache > 0 : read from TzURL if file in cache older than Cache days, otherwise from cache, if available
		TRY
			IF This.Cache != 0
				IF ADIR(m.IsFile, m.LocalTz) == 1
					IF This.Cache < 0 OR FDATE(m.LocalTz) + This.Cache >= DATE()
						m.iCal = m.ICS.ReadFile(m.LocalTz)
					ENDIF
				ENDIF
			ENDIF
		CATCH
			m.iCal = .NULL.
		ENDTRY

		* no (recent) local copy? fetch from the URL
		IF ISNULL(m.iCal)
			IF This.Timezones.GetKey(m.TzID) != 0
				m.iCal = m.ICS.ReadURL("https://www.tzurl.org/zoneinfo/" + This.Timezones.Item(m.TzID) + ".ics")
				* but save it for later if we are working with cache
				IF !ISNULL(m.iCal) AND This.Cache != 0
					m.SafetySetting = SET("Safety")
					SET SAFETY OFF
					STRTOFILE(m.iCal.Serialize(), m.LocalTz, 0)
					IF m.SafetySetting == "ON"
						SET SAFETY ON
					ENDIF
				ENDIF
			ENDIF
		ENDIF

		* if no iCalendar found, use the one that the application stored, if there is one
		IF ISNULL(m.iCal) AND PCOUNT() == 2
			IF CHR(13) $ m.StoredTz
				m.iCal = m.ICS.Read(m.StoredTz)
			ELSE
				m.iCal = m.ICS.ReadFile(m.StoredTz)
			ENDIF
		ENDIF

		* give up if an iCalendar .ics file could not be loaded
		IF ISNULL(m.iCal)
			RETURN .NULL.
		ENDIF

		m.Timezone = m.iCal.GetTimezone()

		* return a dettached version of the timezone, if a timezone was found
		RETURN IIF(ISNULL(m.Timezone), .NULL., m.Timezone.Recreate())

	ENDFUNC

	* create a minimal timezone definition, to be used from current DATETIME() - disregard the historical information
	FUNCTION Minimal (TzID AS String, StoredTz AS String) AS iCalCompVTIMEZONE

		ASSERT VARTYPE(m.TzID) == "C" AND (PCOUNT() = 1 OR VARTYPE(m.StoredTz) == "C")

		LOCAL Timezone AS iCalCompVTIMEZONE
		LOCAL Minimal AS iCalCompVTIMEZONE
		LOCAL TzComp AS _iCalComponent
		LOCAL Now AS Datetime
		LOCAL StandardStart AS Datetime
		LOCAL DaylightStart AS Datetime
		LOCAL DateStart AS Datetime
		LOCAL Standard AS iCalCompSTANDARD
		LOCAL FallbackStandard AS iCalCompSTANDARD
		LOCAL Daylight AS iCalCompDAYLIGHT
		LOCAL LoopIndex AS Integer
		LOCAL RRule AS iCalPropRRULE
		LOCAL VRule AS iCalTypeRECUR
		LOCAL UntilDate AS Datetime
		LOCAL SetDates AS Integer
		LOCAL AdditionalDate AS Datetime

		* get the full definition
		IF PCOUNT() = 1
			m.Timezone = This.Full(m.TzID)
		ELSE
			m.Timezone = This.Full(m.TzID, m.StoredTz)
		ENDIF
		IF ISNULL(m.Timezone)
			RETURN .NULL.
		ENDIF

		STORE .NULL. TO m.Standard, m.Daylight, m.Minimal
		STORE {:} TO m.StandardStart, m.DaylightStart
		m.Now = DATETIME()

		* locate the last STANDARD definition
		FOR m.LoopIndex = 1 TO m.Timezone.GetICComponentsCount("STANDARD")
			m.TzComp = m.Timezone.GetICComponent("STANDARD", m.LoopIndex)
			* when did it start
			m.DateStart = m.TzComp.GetICPropertyValue("DTSTART")
			* and some reset in the future
			FOR m.SetDates = 1 TO m.TzComp.GetICPropertiesCount("RDATE")
				m.AdditionalDate = m.TzComp.GetICPropertyValue("RDATE", m.SetDates)
				IF m.AdditionalDate > m.DateStart
					m.DateStart = m.AdditionalDate
				ENDIF
			ENDFOR
			* but consider expired definitions
			m.RRule = m.TzComp.GetICProperty("RRULE")
			IF !ISNULL(m.RRule)
				m.VRule = m.RRule.GetValue()
				m.UntilDate = m.VRule.Until
			ELSE
				m.UntilDate = .NULL.
			ENDIF
			* check if it is the most recent in the past
			IF m.DateStart > m.StandardStart AND m.DateStart < m.Now
				* use it instead of the previous one what was found
				IF m.Now <= NVL(m.UntilDate, m.Now)
					m.Standard = m.TzComp
					m.StandardStart = m.DateStart
				ELSE
					* unless it expired, meanwhile (but save last standard time, anyway)
					IF !ISNULL(m.Standard)
						m.FallbackStandard = m.Standard
					ENDIF
					m.Standard = .NULL.
					m.StandardStart = {/:}
				ENDIF
			ENDIF
		ENDFOR
		* use the fallback, if available and no standard definition for the time zone
		IF ISNULL(m.Standard) AND !ISNULL(m.FallbackStandard)
			m.Standard = m.FallbackStandard
		ENDIF

		* and now the last DAYLIGHT definition (same as above)
		FOR m.LoopIndex = 1 TO m.Timezone.GetICComponentsCount("DAYLIGHT")
			m.TzComp = m.Timezone.GetICComponent("DAYLIGHT", m.LoopIndex)
			m.DateStart = m.TzComp.GetICPropertyValue("DTSTART")
			FOR m.SetDates = 1 TO m.TzComp.GetICPropertiesCount("RDATE")
				m.AdditionalDate = m.TzComp.GetICPropertyValue("RDATE", m.SetDates)
				IF m.AdditionalDate > m.DateStart
					m.DateStart = m.AdditionalDate
				ENDIF
			ENDFOR
			m.RRule = m.TzComp.GetICProperty("RRULE")
			IF !ISNULL(m.RRule)
				m.VRule = m.RRule.GetValue()
				m.UntilDate = m.VRule.Until
			ELSE
				m.UntilDate = .NULL.
			ENDIF
			IF m.DateStart > m.DaylightStart AND m.DateStart < m.Now
				IF m.Now <= NVL(m.UntilDate, m.Now)
					m.DayLight = m.TzComp
					m.DaylightStart = m.DateStart
				ELSE
					m.Daylight = .NULL.
					m.DaylightStart = {/:}
				ENDIF
			ENDIF
		ENDFOR

		* we found at least one of them (as we should)
		IF !ISNULL(m.Standard) OR !ISNULL(m.Daylight)
			* get a dettached copy of the definition, considering only its properties
			m.Minimal = m.Timezone.Recreate(.T.)
			* if we found a standard time, add it to the minimal version
			IF !ISNULL(m.Standard)
				m.Minimal.AddICComponent(m.Standard.Recreate())
			ENDIF
			* if we found a daylight time, add it to the minimal version
			IF !ISNULL(m.Daylight)
				m.Minimal.AddICComponent(m.Daylight.Recreate())
			ENDIF
		ELSE
			* this is for safety
			m.Minimal = m.Timezone
		ENDIF

		RETURN m.Minimal

	ENDFUNC

	* set the cache expiration period, in days
	FUNCTION SetCache (Period AS Integer)

		ASSERT VARTYPE(m.Period) == "N"

		This.Cache = INT(m.Period)

	ENDFUNC

ENDDEFINE
