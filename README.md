# iCal4VFP

A VFP library set to work with iCalendar objects.

## Overview

- Adheres to [RFC 5545](https://tools.ietf.org/html/rfc5545 "RFC 5545") specifications.
- Serializers and parsers for all Components, Properties, and Parameters.
- Includes structured iCalendar data types (RECUR, DURATION, and PERIOD)
- Integrates a Recurrence Rule processor that runs all RFC 5545 examples.
- Reads iCalendar data from .ics files, or from strings.
- Allows for extensions at all levels (Components, Properties, Parameters, and Types).
- Processes timezone information, and translates between UTC and local time.
- Reads and integrates [TZURL](http://tzurl.org "TZURL") data.
- iCalendar scalar types map to VFP data types (Date, Datetime, String, Logical, Numeric).

Go to

- [Classes](classes.md "Classes")
- [Examples](examples.md "Examples")

## Using

iCal4VFP may be used either by managing iCalendar VFP objects, or by processing iCalendar formatted strings (in memory or in documents). Refer to [examples](examples.md "examples") for a quick introduction.

## Dependencies

iCal4VFP depends on [tokenizer](https://github.com/atlopes/tokenizer "tokenizer"), a class to extract tokens from a text.

## License

iCal4VFP is governed by an [UNLICENSE](UNLICENSE.md "UNLICENSE").

## Contributing

By using, testing, forking, pulling requests, pointing out issues.
