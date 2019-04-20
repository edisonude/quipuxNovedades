VERSION 5.00
Begin VB.Form frmProcess 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9945
   LinkTopic       =   "Form2"
   Picture         =   "frmProcess.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProcess 
      Caption         =   "PROCESAR"
      Height          =   1935
      Left            =   720
      TabIndex        =   9
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "configurar excel"
      Height          =   1335
      Left            =   960
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hazlo todo"
      Height          =   855
      Left            =   2760
      TabIndex        =   7
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   5040
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   5640
      TabIndex        =   0
      Top             =   4440
      Width           =   3885
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Festivos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   420
         Left            =   900
         TabIndex        =   5
         Top             =   1470
         Width           =   1035
      End
      Begin VB.Label lHolidays 
         BackColor       =   &H00E0E0E0&
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A07430&
         Height          =   420
         Left            =   2145
         TabIndex        =   4
         Top             =   1440
         Width           =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00A07430&
         X1              =   60
         X2              =   3780
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lYear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A07430&
         Height          =   420
         Left            =   2145
         TabIndex        =   3
         Top             =   930
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIGURACIÓN CALENDARIO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D58417&
         Height          =   420
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   3720
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Año actual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   420
         Left            =   900
         TabIndex        =   1
         Top             =   960
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public excelApp As Excel.APPLICATION
Public workbook As Excel.workbook
Public sheet As Excel.Worksheet

Public hedo As Integer
Public heno As Integer
Public hedf As Integer
Public henf As Integer

Private Function loadExcel()
Dim pathFile As String
pathFile = "C:\git\quipuxNovedades\consolidado.xlsx"
Set excelApp = New Excel.APPLICATION
Set workbook = excelApp.Workbooks.Open(FileName:=pathFile)
Set sheet = workbook.Sheets(1)
End Function

Private Sub cmdProcess_Click()
On Error GoTo closeResources


sngTime = Timer





Call loadExcel

'Dim pathFile As String
'Dim excelApp As Excel.APPLICATION
'Dim workbook As Excel.workbook
'Dim sheet As Excel.Worksheet

'pathFile = "C:\git\quipuxNovedades\consolidado.xlsx"
'Set excelApp = New Excel.APPLICATION
'Set workbook = excelApp.Workbooks.Open(FileName:="C:\git\quipuxNovedades\consolidado.xlsx")
'Set sheet = workbook.Sheets(1)

Dim hasMoreRows As Boolean
Dim row As Integer
Dim rowsProcessed As Integer

hasMoreRows = True
row = ModConfig.ROW_START_READ
rowsProcessed = 0
'row = 1

While hasMoreRows
    value = sheet.Cells(row, ModConfig.COL_TYPE_ROW)
    If (value = "") Then
        hasMoreRows = False
    ElseIf (isTypeForProcess(row)) Then
        Dim dateStart As Date
        Dim dateEnd As Date
        Dim dateReference As Date
        Dim difGeneral As Integer
        Dim dif As Integer
        Dim isDiurnal As Boolean

        hedo = 0
        heno = 0
        hedf = 0
        henf = 0
        
        Dim dateStartDiurnal As Date
        Dim dateEndDiurnal As Date
        
        dateStart = getDateStart(row)
        dateEnd = getDateEnd(row, dateStart)
        
        dateReference = dateStart
        difGeneral = DateDiff("n", dateReference, dateEnd)

        While difGeneral > 0
            dateStartDiurnal = getDateStartDiurnal(dateReference)
continueDiffStartDate:
            dif = DateDiff("n", dateReference, dateStartDiurnal)
            isDiurnal = False
        
            If (dif <= 0) Then 'TODO: que pasa si es igual
                dateEndDiurnal = getDateEndDiurnal(dateReference)
                dif = DateDiff("n", dateReference, dateEndDiurnal)
                isDiurnal = True
                
                If (dif <= 0) Then
                    dateStartDiurnal = getDateStartDiurnal(DateAdd("d", 1, dateReference))
                    GoTo continueDiffStartDate
                Else
                    If (dif <= difGeneral) Then
                        Call assignDiference(dif, isDiurnal, dateReference, dateEndDiurnal)
                        dateReference = dateEndDiurnal
                        difGeneral = difGeneral - dif
                    Else
                        Call assignDiference(difGeneral, isDiurnal, dateReference, dateEndDiurnal)
                        difGeneral = difGeneral - difGeneral
                    End If
                End If
            Else
                If (dif <= difGeneral) Then
                    Call assignDiference(dif, isDiurnal, dateReference, dateStartDiurnal)
                    dateReference = dateStartDiurnal
                    difGeneral = difGeneral - dif
                Else
                    Call assignDiference(difGeneral, isDiurnal, dateReference, dateStartDiurnal)
                    difGeneral = difGeneral - difGeneral
                End If
            End If
        Wend
        rowsProcessed = rowsProcessed + 1
        Call writeResults(row)
    End If
    row = row + 1
Wend

workbook.Save
workbook.Close SaveChanges:=False

sngTime = Timer - sngTime
    
MsgBox "procesados " & rowsProcessed & vbNewLine & "Tiempo == " & FormatNumber(sngTime, 3)
'MsgBox "hedo " & Round(hedo / 60, 2) & vbNewLine & "heno " & Round(heno / 60, 2)

closeResources:
Call closeResources
End Sub

Private Sub writeResults(row)
If ("HORA EXTRA" = sheet.Cells(row, ModConfig.COL_TYPE_ROW)) Then
    sheet.Cells(row, ModConfig.COL_HEDO) = Round(hedo / 60, 2)
    sheet.Cells(row, ModConfig.COL_HENO) = Round(heno / 60, 2)
    sheet.Cells(row, ModConfig.COL_HEDF) = Round(hedf / 60, 2)
    sheet.Cells(row, ModConfig.COL_HENF) = Round(henf / 60, 2)
Else
    sheet.Cells(row, ModConfig.COL_RN) = Round(heno / 60, 2)
    sheet.Cells(row, ModConfig.COL_RF) = Round(hedf / 60, 2)
    sheet.Cells(row, ModConfig.COL_RNF) = Round(henf / 60, 2)
End If
End Sub
Private Sub assignDiference(dif As Integer, isDiurnal As Boolean, dateStart As Date, dateEnd As Date)
If (Not isDiurnal) Then
    If (day(dateStart) <> day(dateEnd)) Then
        Dim difD1 As Integer
        Dim difD2 As Integer
        
        difD1 = DateDiff("n", dateStart, getNextDay(dateStart))
        If (difD1 < dif) Then
            difD2 = dif - difD1
            Call assignDif(difD1, isDiurnal, dateStart)
            Call assignDif(difD2, isDiurnal, dateEnd)
            Exit Sub
        End If
    End If
End If
Call assignDif(dif, isDiurnal, dateStart)
End Sub

Private Function assignDif(dif As Integer, isDiurnal As Boolean, dateAsign As Date) As Date
If (isDiurnal) Then
    If (ModCalendar.existHoliday(dateAsign)) Then
        hedf = hedf + dif
    Else
        hedo = hedo + dif
    End If
Else
    If (ModCalendar.existHoliday(dateAsign)) Then
        henf = henf + dif
    Else
        heno = heno + dif
    End If
End If
End Function

Private Function getNextDay(d As Date) As Date
getNextDay = Format(DateAdd("d", 1, d), "dd/MM/yyyy 00:00:00")
End Function

Private Function getDateStartDiurnal(d As Date) As Date
getDateStartDiurnal = Format(d, "dd/MM/yyyy " & ModConfig.HOUR_START_D)
End Function

Private Function getDateEndDiurnal(d As Date) As Date
getDateEndDiurnal = Format(d, "dd/MM/yyyy " & ModConfig.HOUR_END_D)
End Function

Private Function getDateStart(row As Integer) As Date
getDateStart = CDate(sheet.Cells(row, ModConfig.COL_DATE)) & " " & CDate(sheet.Cells(row, ModConfig.COL_HOUR_INI))
End Function

Private Function getDateEnd(row As Integer, dateStart As Date) As Date
Dim dateEnd As Date
dateEnd = CDate(sheet.Cells(row, ModConfig.COL_DATE)) & " " & CDate(sheet.Cells(row, ModConfig.COL_HOUR_END))
If (dateStart > dateEnd) Then
    dateEnd = DateAdd("d", 1, dateEnd)
End If
getDateEnd = dateEnd
End Function

Private Function isTypeForProcess(row As Integer) As Boolean
isTypeForProcess = ("HORA EXTRA" = sheet.Cells(row, ModConfig.COL_TYPE_ROW)) Or ("RECARGO NOCTURNO" = sheet.Cells(row, ModConfig.COL_TYPE_ROW))
End Function

Private Function closeResources()
Set workbook = Nothing
If Not excelApp Is Nothing Then
    excelApp.Quit
    Set excelApp = Nothing
End If
End Function

Private Sub Command1_Click()
'ModCalendar.getHolidaysNumberForYear (Me.lYear)
frmCalendar.Show
End Sub


Private Sub Command2_Click()

    Dim diferencia As Double
    Dim hedo As Double
    Dim heno As Double
    Dim colBase As Integer

    Dim oApp As Excel.APPLICATION
    Dim oWB As Excel.workbook
    
    'Create an Excel instalce.
    Set oApp = New Excel.APPLICATION
    'Open the desired workbook
    Set oWB = oApp.Workbooks.Open(FileName:="C:\git\quipux\consolidado.xlsx")
    'Do any modifications to the workbook.
    '...
    
    Dim b As Excel.Worksheet
    Set b = oWB.Sheets(1)
    Dim continuar As Boolean
    continuar = True
    Dim fila As Integer
    Dim value As String
    Dim cuenta As Integer
    fila = 9
    While continuar
        value = b.Cells(fila, 1)
        If (value = "") Then
            continuar = False
        Else
            If (b.Cells(fila, 5) = "HORA EXTRA") Or (b.Cells(fila, 5) = "RECARGO NOCTURNO") Then
                diferencia = 0
                hedo = 0
                heno = 0

                Dim fechaIni As Date
                fechaIni = CDate(b.Cells(fila, 6)) & " " & CDate(b.Cells(fila, 7))
                
                Dim fechaFin As Date
                Dim totalInMin As Double
                totalInMin = CDate(b.Cells(fila, 9)) * 60 * 24
                fechaFin = DateAdd("n", totalInMin, fechaIni)
                
                fechaFin = CDate(Format(fechaFin, "dd/MM/yyyy")) & " " & CDate(b.Cells(fila, 8))
                
                
                diferencia = DateDiff("n", fechaIni, fechaFin)
'                MsgBox "fechaIni " & fechaIni
'                MsgBox "fechaFin " & fechaFin
                
                Dim dateIniDiurno As Date
                dateIniDiurno = Format(fechaIni, "dd/MM/yyyy 06:00:00")
                
                Dim dateIniNocturno As Date
                dateIniNocturno = Format(fechaIni, "dd/MM/yyyy 21:00:00")
                
                Dim dateFinDiurno As Date
                dateFinDiurno = Format(fechaFin, "dd/MM/yyyy 06:00:00")
                
                Dim dateFinNocturno As Date
                dateFinNocturno = Format(fechaFin, "dd/MM/yyyy 21:00:00")
                
                If (fechaIni >= dateIniDiurno And fechaIni <= dateIniNocturno) Then
                    hedo = DateDiff("n", fechaIni, dateIniNocturno)
                    If (diferencia < hedo) Then
                        hedo = diferencia
                    Else
                        heno = DateDiff("n", dateIniNocturno, fechaFin)
                    End If
                Else
                    heno = DateDiff("n", fechaIni, dateFinDiurno)
                    If (diferencia < heno Or heno < 0) Then
                        heno = diferencia
                    Else
                        hedo = DateDiff("n", dateFinDiurno, fechaFin)
                    End If
                End If
                
                hedo = Round(hedo / 60, 2)
                heno = Round(heno / 60, 2)
                
                If (b.Cells(fila, 5) = "HORA EXTRA") Then
                    b.Cells(fila, 10) = hedo
                    b.Cells(fila, 11) = heno
                Else
                    b.Cells(fila, 14) = heno
                End If
                
                fila = fila + 1
                cuenta = cuenta + 1
            End If
        End If
    Wend

    'Save the xls file
    ' oWB.SaveAs FileName:="C:\git\quipux\Libro.xlsx"
    oWB.Save
    'close and clean up resources
    oWB.Close SaveChanges:=False
    Set oWB = Nothing
    oApp.Quit
    Set oApp = Nothing
    
    MsgBox "end"
End Sub

Private Sub Command3_Click()
frmConfigExcel.Show
End Sub
