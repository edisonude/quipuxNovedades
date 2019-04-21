Attribute VB_Name = "ModConfig"
Option Explicit

Public calendarPath As String
Public currentYear As String

Public ROW_START_READ As Integer
Public COL_TYPE_ROW As Integer
Public COL_DATE As Integer
Public COL_HOUR_INI As Integer
Public COL_HOUR_END As Integer
Public COL_HEDO As Integer
Public COL_HENO As Integer
Public COL_HEDF As Integer
Public COL_HENF As Integer
Public COL_RN As Integer
Public COL_RNF As Integer
Public COL_RF As Integer

Public HOUR_START_D As String
Public HOUR_END_D As String

Sub main()
calendarPath = App.Path & "\config.ini"
currentYear = year(Now)
Call reloadHolidays
Call loadExcelConfig

HOUR_START_D = ModIni.readPropertyFile(calendarPath, ModIni.K_HOUR_START_D, "06:00:00")
HOUR_END_D = ModIni.readPropertyFile(calendarPath, ModIni.K_HOUR_END_D, "21:00:00")
frmProcess.Show
End Sub

Public Sub reloadHolidays()
ReDim holidays(0)
Call ModCalendar.loadHolidays(currentYear - 1)
Call ModCalendar.loadHolidays(currentYear)
Call ModCalendar.loadHolidays(currentYear + 1)
End Sub

Public Function loadExcelConfig()
ROW_START_READ = ModIni.readPropertyFile(calendarPath, ModIni.K_ROW_START_READ, 9)
COL_TYPE_ROW = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_TYPE_ROW, 5)
COL_DATE = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_DATE, 6)
COL_HOUR_INI = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HOUR_INI, 7)
COL_HOUR_END = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HOUR_END, 8)
COL_HEDO = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HEDO, 10)
COL_HENO = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HENO, 11)
COL_HEDF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HEDF, 12)
COL_HENF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_HENF, 13)
COL_RN = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_RN, 14)
COL_RNF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_RNF, 15)
COL_RF = ModIni.readPropertyFile(calendarPath, ModIni.K_COL_RF, 16)
End Function
