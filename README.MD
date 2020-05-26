# Calendar

`Calendar` returns the selected date on the created form.

Tested in Windows Excel 2019, but should apply to 2007+.

# Example

```vb
' (Add reference to Microsoft Forms 2.0 Object Library)
Dim Calendar As New clsCalendar

Debug.Print Calendar.getDate ' or Debug.Print Calendar
' -> returnes selected date

Set Calendar = Nothing
```

## Options

Calendar-VBA includes a few options for customizing starting interface if needed:
- __Caption__ (Default = `Calendar`) Use this option to rename calendar.
- __FirstDayOfTheWeek__ (Default = `vbUseSystemDayOfWeek`) Use this option to reassign the first day of the week.

## Installation
1. Download the [latest release](https://github.com/VBA-tools/VBA-JSON/releases)
2. Import `clsCalendar.cls`, `clsWidget.cls` and `clsEventsForWidget.cls` into your project (Open VBA Editor, `Alt + F11`; File > Import File)
3. Add `Dictionary` reference/class
   - For Windows-only, include a reference to "Microsoft Forms 2.0 Object Library"