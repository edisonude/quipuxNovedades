VERSION 5.00
Begin VB.Form frmCalendar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración del calendario"
   ClientHeight    =   10125
   ClientLeft      =   360
   ClientTop       =   390
   ClientWidth     =   15045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCalendar.frx":0000
   ScaleHeight     =   10125
   ScaleWidth      =   15045
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2445
      Left            =   10560
      TabIndex        =   12
      Top             =   240
      Width           =   4110
      Begin VB.ComboBox cmbYears 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Image btnSave 
         Height          =   525
         Left            =   840
         Picture         =   "frmCalendar.frx":1527CE
         Top             =   1680
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Año actual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   420
         Left            =   540
         TabIndex        =   14
         Top             =   930
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
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
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   3720
      End
   End
   Begin VB.Label lOct 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   11
      Top             =   8250
      Width           =   255
   End
   Begin VB.Label lNov 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   10
      Top             =   8250
      Width           =   255
   End
   Begin VB.Label lDic 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   7110
      TabIndex        =   9
      Top             =   8250
      Width           =   255
   End
   Begin VB.Label lJul 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   8
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label lAgo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   7
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label lSep 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   7110
      TabIndex        =   6
      Top             =   5760
      Width           =   255
   End
   Begin VB.Label lAbr 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   5
      Top             =   3270
      Width           =   255
   End
   Begin VB.Label lMay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   3270
      Width           =   255
   End
   Begin VB.Label lJun 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   7110
      TabIndex        =   3
      Top             =   3270
      Width           =   255
   End
   Begin VB.Label lMar 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   7110
      TabIndex        =   2
      Top             =   780
      Width           =   255
   End
   Begin VB.Label lFeb 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   780
      Width           =   255
   End
   Begin VB.Label lEne 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   780
      Width           =   255
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public yearCalendar As Integer
Public frmParent As Form

Private Sub btnSave_Click()
Dim holidaysLine As String
For i = 1 To 12
    firstDay = Weekday(DateSerial(Me.yearCalendar, i, 1), vbMonday)
    lastDay = day(DateSerial(Me.yearCalendar, i + 1, 0))
    holidaysLine = ""
    For d = 1 To 42
        Select Case i
            Case 1
                addDayIfIsMarkAsHoliday Me.lEne(d), holidaysLine
            Case 2
                addDayIfIsMarkAsHoliday Me.lFeb(d), holidaysLine
            Case 3
                addDayIfIsMarkAsHoliday Me.lMar(d), holidaysLine
            Case 4
                addDayIfIsMarkAsHoliday Me.lAbr(d), holidaysLine
            Case 5
                addDayIfIsMarkAsHoliday Me.lMay(d), holidaysLine
            Case 6
                addDayIfIsMarkAsHoliday Me.lJun(d), holidaysLine
            Case 7
                addDayIfIsMarkAsHoliday Me.lJul(d), holidaysLine
            Case 8
                addDayIfIsMarkAsHoliday Me.lAgo(d), holidaysLine
            Case 9
                addDayIfIsMarkAsHoliday Me.lSep(d), holidaysLine
            Case 10
                addDayIfIsMarkAsHoliday Me.lOct(d), holidaysLine
            Case 11
                addDayIfIsMarkAsHoliday Me.lNov(d), holidaysLine
            Case 12
                addDayIfIsMarkAsHoliday Me.lDic(d), holidaysLine
        End Select

    Next
    Call ModIni.savePropertyFile(ModConfig.calendarPath, Me.yearCalendar & i, holidaysLine)
Next
ModConfig.reloadHolidays
MsgBox "¡Calendario actualizado!", vbInformation
Me.frmParent.updateHolidays
End Sub

Private Sub cmbYears_Click()
Me.yearCalendar = Me.cmbYears.Text
Call loadCalendar
End Sub

Private Sub cmbYears_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub addDayIfIsMarkAsHoliday(lday As Control, holidaysLine As String)
If (lday.ForeColor = vbRed) Then
    holidaysLine = holidaysLine & lday.Caption & "|"
End If
End Sub

Private Sub Form_Load()
Me.yearCalendar = ModConfig.currentYear

Me.cmbYears.AddItem Me.yearCalendar - 1
Me.cmbYears.AddItem Me.yearCalendar
Me.cmbYears.AddItem Me.yearCalendar + 1
Me.cmbYears.Text = Me.yearCalendar

Call createLabelsCalendar
Call loadCalendar
End Sub

Private Sub loadCalendar()
Call clearCalendar
Call loadCalendarDays
Call loadHolidays
End Sub

Private Sub clearCalendar()
For j = 0 To 5
    For i = 0 To 6
        pos = j * 7 + i + 1
        Call clearLabelCalendar(lEne(pos))
        Call clearLabelCalendar(lFeb(pos))
        Call clearLabelCalendar(lMar(pos))
        Call clearLabelCalendar(lAbr(pos))
        Call clearLabelCalendar(lMay(pos))
        Call clearLabelCalendar(lJun(pos))
        Call clearLabelCalendar(lJul(pos))
        Call clearLabelCalendar(lAgo(pos))
        Call clearLabelCalendar(lSep(pos))
        Call clearLabelCalendar(lOct(pos))
        Call clearLabelCalendar(lNov(pos))
        Call clearLabelCalendar(lDic(pos))
    Next
Next
End Sub

Private Sub clearLabelCalendar(label As label)
label.Caption = ""
label.ForeColor = &H404040
label.FontBold = False
End Sub

Private Sub createLabelsCalendar()
Dim pos As Integer
Dim alto As Integer
Dim m As Integer
alto = lEne(0).Top
For j = 0 To 5
    For i = 0 To 6
        pos = j * 7 + i + 1
        Load lEne(pos)
        Load lFeb(pos)
        Load lMar(pos)
        Load lAbr(pos)
        Load lMay(pos)
        Load lJun(pos)
        Load lJul(pos)
        Load lAgo(pos)
        Load lSep(pos)
        Load lOct(pos)
        Load lNov(pos)
        Load lDic(pos)
        
        lEne(pos).Visible = True
        lFeb(pos).Visible = True
        lMar(pos).Visible = True
        lAbr(pos).Visible = True
        lMay(pos).Visible = True
        lJun(pos).Visible = True
        lJul(pos).Visible = True
        lAgo(pos).Visible = True
        lSep(pos).Visible = True
        lOct(pos).Visible = True
        lNov(pos).Visible = True
        lDic(pos).Visible = True
        
        If (j > 0) Then
            lEne(pos).Left = lEne(pos - (j * 7)).Left
            lFeb(pos).Left = lFeb(pos - (j * 7)).Left
            lMar(pos).Left = lMar(pos - (j * 7)).Left
            lAbr(pos).Left = lAbr(pos - (j * 7)).Left
            lMay(pos).Left = lMay(pos - (j * 7)).Left
            lJun(pos).Left = lJun(pos - (j * 7)).Left
            lJul(pos).Left = lJul(pos - (j * 7)).Left
            lAgo(pos).Left = lAgo(pos - (j * 7)).Left
            lSep(pos).Left = lSep(pos - (j * 7)).Left
            lOct(pos).Left = lOct(pos - (j * 7)).Left
            lNov(pos).Left = lNov(pos - (j * 7)).Left
            lDic(pos).Left = lDic(pos - (j * 7)).Left
        Else
            lEne(pos).Left = lEne(pos - 1).Width + lEne(pos - 1).Left + 120
            lFeb(pos).Left = lFeb(pos - 1).Width + lFeb(pos - 1).Left + 120
            lMar(pos).Left = lMar(pos - 1).Width + lMar(pos - 1).Left + 120
            lAbr(pos).Left = lAbr(pos - 1).Width + lAbr(pos - 1).Left + 120
            lMay(pos).Left = lMay(pos - 1).Width + lMay(pos - 1).Left + 120
            lJun(pos).Left = lJun(pos - 1).Width + lJun(pos - 1).Left + 120
            lJul(pos).Left = lJul(pos - 1).Width + lJul(pos - 1).Left + 120
            lAgo(pos).Left = lAgo(pos - 1).Width + lAgo(pos - 1).Left + 120
            lSep(pos).Left = lSep(pos - 1).Width + lSep(pos - 1).Left + 120
            lOct(pos).Left = lOct(pos - 1).Width + lOct(pos - 1).Left + 120
            lNov(pos).Left = lNov(pos - 1).Width + lNov(pos - 1).Left + 120
            lDic(pos).Left = lDic(pos - 1).Width + lDic(pos - 1).Left + 120
        End If
        lEne(pos).Top = lEne(0).Top + ((lEne(0).Height + 31) * j)
        lFeb(pos).Top = lFeb(0).Top + ((lFeb(0).Height + 31) * j)
        lMar(pos).Top = lMar(0).Top + ((lMar(0).Height + 31) * j)
        lAbr(pos).Top = lAbr(0).Top + ((lAbr(0).Height + 31) * j)
        lMay(pos).Top = lMay(0).Top + ((lMay(0).Height + 31) * j)
        lJun(pos).Top = lJun(0).Top + ((lJun(0).Height + 31) * j)
        lJul(pos).Top = lJul(0).Top + ((lJul(0).Height + 31) * j)
        lAgo(pos).Top = lAgo(0).Top + ((lAgo(0).Height + 31) * j)
        lSep(pos).Top = lSep(0).Top + ((lMar(0).Height + 31) * j)
        lOct(pos).Top = lOct(0).Top + ((lOct(0).Height + 31) * j)
        lNov(pos).Top = lNov(0).Top + ((lNov(0).Height + 31) * j)
        lDic(pos).Top = lDic(0).Top + ((lDic(0).Height + 31) * j)
    Next
Next
End Sub

Private Function loadCalendarDays()
Dim firstDay As Integer
Dim lastDay As Integer

For i = 1 To 12
    firstDay = Weekday(DateSerial(Me.yearCalendar, i, 1), vbMonday)
    lastDay = day(DateSerial(Me.yearCalendar, i + 1, 0))
    For d = 1 To lastDay
        Select Case i
            Case 1
                Me.lEne(firstDay + d - 1) = d
            Case 2
                Me.lFeb(firstDay + d - 1) = d
            Case 3
                Me.lMar(firstDay + d - 1) = d
            Case 4
                Me.lAbr(firstDay + d - 1) = d
            Case 5
                Me.lMay(firstDay + d - 1) = d
            Case 6
                Me.lJun(firstDay + d - 1) = d
            Case 7
                Me.lJul(firstDay + d - 1) = d
            Case 8
                Me.lAgo(firstDay + d - 1) = d
            Case 9
                Me.lSep(firstDay + d - 1) = d
            Case 10
                Me.lOct(firstDay + d - 1) = d
            Case 11
                Me.lNov(firstDay + d - 1) = d
            Case 12
                Me.lDic(firstDay + d - 1) = d
        End Select
    Next
Next
End Function

Private Function loadHolidays()
For d = 0 To UBound(ModCalendar.holidays)
    Call markDateAsHoliday(ModCalendar.holidays(d))
Next
End Function

Private Sub markDateAsHoliday(holidayDate As Date)
Dim m As Integer
Dim firstDay As Integer
Dim posHoliday As Integer

m = Month(holidayDate)
If (Not Me.yearCalendar = year(holidayDate)) Then
    Exit Sub
End If

firstDay = Weekday(DateSerial(year(holidayDate), m, 1), vbMonday) - 1
posHoliday = firstDay + day(holidayDate)

Select Case m
    Case 1
        checkHoliday Me.lEne(posHoliday)
    Case 2
        checkHoliday Me.lFeb(posHoliday)
    Case 3
        checkHoliday Me.lMar(posHoliday)
    Case 4
        checkHoliday Me.lAbr(posHoliday)
    Case 5
        checkHoliday Me.lMay(posHoliday)
    Case 6
        checkHoliday Me.lJun(posHoliday)
    Case 7
        checkHoliday Me.lJul(posHoliday)
    Case 8
        checkHoliday Me.lAgo(posHoliday)
    Case 9
        checkHoliday Me.lSep(posHoliday)
    Case 10
        checkHoliday Me.lOct(posHoliday)
    Case 11
        checkHoliday Me.lNov(posHoliday)
    Case 12
        checkHoliday Me.lDic(posHoliday)
End Select
End Sub

Private Sub lEne_Click(Index As Integer)
Call checkHoliday(lEne(Index))
End Sub
Private Sub lFeb_Click(Index As Integer)
Call checkHoliday(lFeb(Index))
End Sub
Private Sub lMar_Click(Index As Integer)
Call checkHoliday(lMar(Index))
End Sub
Private Sub lAbr_Click(Index As Integer)
Call checkHoliday(lAbr(Index))
End Sub
Private Sub lMay_Click(Index As Integer)
Call checkHoliday(lMay(Index))
End Sub
Private Sub lJun_Click(Index As Integer)
Call checkHoliday(lJun(Index))
End Sub
Private Sub lJul_Click(Index As Integer)
Call checkHoliday(lJul(Index))
End Sub
Private Sub lAgo_Click(Index As Integer)
Call checkHoliday(lAgo(Index))
End Sub
Private Sub lSep_Click(Index As Integer)
Call checkHoliday(lSep(Index))
End Sub
Private Sub lOct_Click(Index As Integer)
Call checkHoliday(lOct(Index))
End Sub
Private Sub lNov_Click(Index As Integer)
Call checkHoliday(lNov(Index))
End Sub
Private Sub lDic_Click(Index As Integer)
Call checkHoliday(lDic(Index))
End Sub

Private Function checkHoliday(lday As Control)
If (lday.ForeColor = vbRed) Then
    lday.ForeColor = &H404040
    lday.FontBold = False
Else
    lday.ForeColor = vbRed
    lday.FontBold = True
End If
End Function
