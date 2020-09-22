<div align="center">

## Unix TimeStamp


</div>

### Description

My Code will Take the Current Time and Convert it to a Unix Time Stamp
 
### More Info
 
No inputs

Returns an String which will be the Unix Timestamp


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Christopher Trahan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/christopher-trahan.md)
**Level**          |Advanced
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/christopher-trahan-unix-timestamp__1-44477/archive/master.zip)





### Source Code

```
Public Function TimeStamp() As String
  Dim StartDate As String
  Dim EndTime As String
  Dim StartTime As String
  Dim EndDate As String
  Dim dblStart As Double
  Dim dblEnd As Double
  Dim DateTimeStart As Date
  Dim DateTimeEnd As Date
  Dim TotalHrs As String
  StartDate = "1/1/1970"
  StartTime = "00:00:00"
  EndDate = CStr(Date)
  EndTime = CStr(Time)
  DateTimeStart = FormatDateTime(StartDate & " " & StartTime)
  DateTimeEnd = FormatDateTime(EndDate & " " & EndTime)
  TimeStamp = DateDiff("s", DateTimeStart, DateTimeEnd, vbUseSystemDayOfWeek, _
  vbUseSystem)
End Function
```

