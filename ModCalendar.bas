Attribute VB_Name = "ModCalendar"
Public holidays() As Date
Public currentYear As Integer

Public Function loadHolidays(year)
'ReDim Preserve holidays(0)
Dim holidaysLine As String
Dim holidaysByMonth() As String
Dim countHolidays As Integer
Dim holidayDate As Date

countHolidays = getInitialCount()

For m = 1 To 12
    holidaysLine = ModIni.readPropertyFile(ModConfig.calendarPath, year & m, "")
    holidaysByMonth = Split(holidaysLine, "|")
    For d = 0 To UBound(holidaysByMonth)
        If (isValidDay(holidaysByMonth(d))) Then
            holidayDate = DateSerial(year, m, holidaysByMonth(d))
            If Not (existHoliday(holidayDate)) Then
                ReDim Preserve holidays(countHolidays)
                holidays(countHolidays) = holidayDate
                countHolidays = countHolidays + 1
            End If
        End If
    Next
Next
End Function

Public Function getInitialCount() As Integer
On Error GoTo out
getInitialCount = UBound(holidays) + 1
Exit Function
out:
getInitialCount = 0
End Function

Public Function existHoliday(holidayDate As Date) As Boolean
On Error GoTo out
For d = 0 To UBound(holidays)
    If (holidays(d) = Format(holidayDate, "dd/mm/yyyy")) Then
        existHoliday = True
        Exit Function
    End If
Next
existHoliday = False
Exit Function
out:
If Err.Number = 9 Then
    existHoliday = False
End If
End Function

Public Function getHolidaysNumberForYear(year As Integer)
Dim holidays As String
Dim holidaysByMonth() As String
Dim countHolidays As Integer
For m = 1 To 12
    holidays = ModIni.readPropertyFile(calendarPath, year & m, "")
    holidaysByMonth = Split(holidays, "|")
    For d = 0 To UBound(holidaysByMonth)
        If (isValidDay(holidaysByMonth(d))) Then
            countHolidays = countHolidays + 1
        End If
    Next
Next
MsgBox countHolidays
End Function

Private Function isValidDay(day As String) As Boolean
If (IsNumeric(day)) Then
    isValidDay = CInt(day) > 0 And CInt(day) < 32
Else
    isValidDay = False
End If
End Function



