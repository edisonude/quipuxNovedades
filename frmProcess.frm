VERSION 5.00
Begin VB.Form frmProcess 
   Caption         =   "Form2"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10650
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   1335
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   2760
      TabIndex        =   7
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   4875
      TabIndex        =   0
      Top             =   480
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
Private Sub Command1_Click()
ModCalendar.getHolidaysNumberForYear (Me.lYear)
End Sub


Private Sub Command2_Click()
    Dim diferencia As Double
    Dim hedo As Double
    Dim heno As Double
    Dim colBase As Integer

    Dim oApp As Excel.APPLICATION
    Dim oWB As Excel.Workbook
    
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
