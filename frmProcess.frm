VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProcess 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Novedades"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProcess.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog excelDialog 
      Left            =   -240
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccione el archivo de novedades a procesar"
      Filter          =   "Archivos Excel (xlsx)|*.xlsx|Archivos Excel (xls)|*.xls"
   End
   Begin VB.PictureBox picProcessing 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   960
      Picture         =   "frmProcess.frx":1CB5A
      ScaleHeight     =   4800
      ScaleWidth      =   8220
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   8220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H001BE7E1&
      BorderWidth     =   4
      Index           =   2
      X1              =   6720
      X2              =   6705
      Y1              =   1620
      Y2              =   3990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H001BE7E1&
      BorderWidth     =   4
      Index           =   1
      X1              =   240
      X2              =   9405
      Y1              =   1605
      Y2              =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H001BE7E1&
      BorderWidth     =   4
      Index           =   0
      X1              =   240
      X2              =   9330
      Y1              =   3990
      Y2              =   3975
   End
   Begin VB.Image btnInforms 
      Height          =   2280
      Left            =   6720
      Picture         =   "frmProcess.frx":9C734
      Top             =   1755
      Width           =   2640
   End
   Begin VB.Image btnEnd 
      Height          =   630
      Left            =   2550
      Picture         =   "frmProcess.frx":B00F6
      Top             =   5640
      Width           =   4710
   End
   Begin VB.Label lHolidays 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BE7E1&
      Height          =   255
      Left            =   3195
      TabIndex        =   3
      Top             =   4785
      Width           =   300
   End
   Begin VB.Label lYear 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2018"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H001BE7E1&
      Height          =   255
      Left            =   2955
      TabIndex        =   2
      Top             =   4485
      Width           =   540
   End
   Begin VB.Image btnDiurnalNocturnal 
      Height          =   1065
      Left            =   5760
      Picture         =   "frmProcess.frx":B9C18
      Top             =   4425
      Width           =   1785
   End
   Begin VB.Image btnConfigExcel 
      Height          =   1065
      Left            =   3840
      Picture         =   "frmProcess.frx":C0032
      Top             =   4425
      Width           =   1770
   End
   Begin VB.Image btnCalendar 
      Height          =   1065
      Left            =   1920
      Picture         =   "frmProcess.frx":C6330
      Top             =   4425
      Width           =   1770
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   240
      Picture         =   "frmProcess.frx":CC62E
      Top             =   3600
      Width           =   2340
   End
   Begin VB.Image btnProcess 
      Height          =   630
      Left            =   1110
      Picture         =   "frmProcess.frx":CF424
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label lUpload 
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   5790
      TabIndex        =   1
      Top             =   1905
      Width           =   660
   End
   Begin VB.Label lExcelFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione el archivo de novedades"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   735
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   615
      Picture         =   "frmProcess.frx":D8E9E
      Top             =   1875
      Width           =   5865
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Picture         =   "frmProcess.frx":E5900
      Top             =   1200
      Width           =   4560
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public excelPath As String
Public excelApp As Excel.APPLICATION
Public workbook As Excel.workbook
Public sheet As Excel.Worksheet

Public totMins As Integer
Public hedo As Integer
Public heno As Integer
Public hedf As Integer
Public henf As Integer

Const MAX_HOURS_DAY = 2
Const MAX_HOURS_WEEK = 12
Const MAX_HOURS_MONTH = 48

'Data columns
Const INI_ROW = 2
Const COL_CEDULA = 1
Const COL_NOMBRE = 2
Const COL_CARGO = 3
Const COL_VICEPRE = 5
Const COL_FECHA = 7
Const COL_TOTHOR = 11

Private Function loadExcel()
Set excelApp = New Excel.APPLICATION
Set workbook = excelApp.Workbooks.Open(FileName:=excelPath)
Set sheet = workbook.Sheets(1)
End Function

Private Sub btnCalendar_Click()
frmCalendar.Show
Set frmCalendar.frmParent = Me
End Sub

Private Sub btnConfigExcel_Click()
frmConfigExcel.Show
End Sub

Private Sub btnDiurnalNocturnal_Click()
frmDiurnalNocturnal.Show
End Sub

Private Sub btnEnd_Click()
End
End Sub

Private Sub btnInforms_Click()
On Error GoTo closeResources

If (Me.excelPath = "") Then
    MsgBox "Debe seleccionar el archivo de novedades que quiere procesar", vbCritical
    Exit Sub
End If

Call showProcessing
Call loadExcel

Dim hasMoreRows As Boolean
Dim row As Integer
Dim rowsProcessed As Integer

hasMoreRows = True
row = INI_ROW
rowsProcessed = 0

Dim dict As Dictionary
Set dict = New Dictionary

Dim cedula As String
Dim he As Double
Dim infoEmpleado As CEntryReport

While hasMoreRows
    cedula = sheet.Cells(row, COL_CEDULA)
    If (cedula = "") Then
        hasMoreRows = False
    Else
        he = sheet.Cells(row, COL_TOTHOR)
        If dict.Exists(cedula) Then
            Set infoEmpleado = dict.Item(cedula)
            infoEmpleado.totalHoras = infoEmpleado.totalHoras + he
            Set dict.Item(cedula) = infoEmpleado
        Else
            Set infoEmpleado = New CEntryReport
            infoEmpleado.cedula = cedula
            infoEmpleado.nombre = sheet.Cells(row, COL_NOMBRE)
            infoEmpleado.cargo = sheet.Cells(row, COL_CARGO)
            infoEmpleado.vicepresidencia = sheet.Cells(row, COL_VICEPRE)
            infoEmpleado.totalHoras = sheet.Cells(row, COL_TOTHOR)
            dict.Add infoEmpleado.cedula, infoEmpleado
        End If
    End If
    row = row + 1
Wend

workbook.Save
workbook.Close SaveChanges:=False

MsgBox "en dict " & dict.Count

Call showProcessing
    
MsgBox "Finalizó con éxito el procesamiento de las novedades." & vbNewLine & vbNewLine & _
    "Se procesaron " & rowsProcessed & " regisros de novedades.", vbInformation


closeResources:
MsgBox Err.Description
Call closeResources
End Sub

Private Sub btnProcess_Click()
On Error GoTo closeResources

If (Me.excelPath = "") Then
    MsgBox "Debe seleccionar el archivo de novedades que quiere procesar", vbCritical
    Exit Sub
End If

Call showProcessing
Call loadExcel

Dim hasMoreRows As Boolean
Dim row As Integer
Dim rowsProcessed As Integer

hasMoreRows = True
row = ModConfig.ROW_START_READ
rowsProcessed = 0

While hasMoreRows
    Value = sheet.Cells(row, ModConfig.COL_TYPE_ROW)
    If (Value = "") Then
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
        totMins = difGeneral

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

Call showProcessing
    
MsgBox "Finalizó con éxito el procesamiento de las novedades." & vbNewLine & vbNewLine & _
    "Se procesaron " & rowsProcessed & " regisros de novedades.", vbInformation
closeResources:
Call closeResources
End Sub

Private Sub showProcessing()
If (Me.picProcessing.Visible = False) Then
    Me.picProcessing.Visible = True
    Me.picProcessing.Top = 1680
Else
    Me.picProcessing.Visible = False
End If
End Sub

Private Sub writeResults(row)
Dim totalToReport As Integer
If ("HORA EXTRA" = sheet.Cells(row, ModConfig.COL_TYPE_ROW)) Then
    sheet.Cells(row, ModConfig.COL_HEDO) = IIf(hedo > 0, Round(hedo / 60, 2), "")
    sheet.Cells(row, ModConfig.COL_HENO) = IIf(heno > 0, Round(heno / 60, 2), "")
    sheet.Cells(row, ModConfig.COL_HEDF) = IIf(hedf > 0, Round(hedf / 60, 2), "")
    sheet.Cells(row, ModConfig.COL_HENF) = IIf(henf > 0, Round(henf / 60, 2), "")
    totalToReport = hedo + heno + hedf + henf
Else
    sheet.Cells(row, ModConfig.COL_RN) = IIf(heno > 0, Round(heno / 60, 2), "")
    sheet.Cells(row, ModConfig.COL_RF) = IIf(hedf > 0, Round(hedf / 60, 2), "")
    sheet.Cells(row, ModConfig.COL_RNF) = IIf(henf > 0, Round(henf / 60, 2), "")
    totalToReport = heno + hedf + henf
End If

If (totMins <> totalToReport) Then
    Call markCheckWithError(row, ModConfig.COL_TYPE_ROW, "El total de horas reportadas no puedieron ser clasificadas en su totalidad según su tipo")
End If

Dim difMinReported As Integer
difMinReported = CDate(sheet.Cells(row, ModConfig.COL_TOT)) * 1440
If difMinReported <> totMins Then
    Call markCheckWithError(row, ModConfig.COL_TOT, "El total de horas reportadas no coinciden con la diferencia de las fechas reportadas")
End If
End Sub

Private Sub markCheckWithError(row, col, message)
sheet.Cells(row, col).Interior.Color = &H80FF&
sheet.Cells(row, col).ClearComments
sheet.Cells(row, col).AddComment message

sheet.Cells(row, col).Comment.Shape.ScaleHeight 2.26, msoFalse, msoScaleFromTopLeft
sheet.Cells(row, col).Comment.Shape.ScaleWidth 5.87, msoFalse, msoScaleFromTopLeft
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

Private Sub Form_Load()
Me.lYear = ModConfig.currentYear
Me.lHolidays = ModCalendar.getHolidaysNumberForYear(ModConfig.currentYear)
End Sub

Public Sub updateHolidays()
Me.lHolidays = ModCalendar.getHolidaysNumberForYear(ModConfig.currentYear)
End Sub

Private Sub Image2_Click()

End Sub

Private Sub lUpload_Click()
excelDialog.ShowOpen
If excelDialog.FileName <> "" Then
    excelPath = excelDialog.FileName
    Me.lExcelFile = excelDialog.FileTitle
Else
    excelPath = ""
    Me.lExcelFile = "Seleccione el archivo de novedades"
End If
End Sub


