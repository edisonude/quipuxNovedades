Attribute VB_Name = "ModCalendar"
Public calendarPath As String
Public holidays() As Date
Public currentYear As Integer

Public Function loadHolidays(year As Integer)
'ReDim Preserve holidays(0)
Dim holidaysLine As String
Dim holidaysByMonth() As String
Dim countHolidays As Integer
Dim holidayDate As Date
For m = 1 To 12
    holidaysLine = ModIni.readPropertyFile(calendarPath, year & m, "")
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

Public Function existHoliday(holidayDate As Date) As Boolean
On Error GoTo out
For d = 0 To UBound(holidays)
    If (holidays(d) = holidayDate) Then
        existHoliday = True
    Else
        existHoliday = False
    End If
Next
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

Sub main()
calendarPath = App.Path & "\calendar.ini"
currentYear = year(Now)
ModCalendar.loadHolidays (currentYear)

frmProcess.Show

End Sub

