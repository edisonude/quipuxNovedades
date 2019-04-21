VERSION 5.00
Begin VB.Form frmDiurnalNocturnal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración jornada diurna/noctura"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDiurnalNoctural.frx":0000
   ScaleHeight     =   3855
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tEndDiurnal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   3240
      TabIndex        =   1
      Text            =   "21:00:00"
      Top             =   2415
      Width           =   2490
   End
   Begin VB.TextBox tStartDiurnal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   3240
      TabIndex        =   0
      Text            =   "06:00:00"
      Top             =   1350
      Width           =   2490
   End
   Begin VB.Image btnSave 
      Height          =   510
      Left            =   2565
      Picture         =   "frmDiurnalNoctural.frx":1801E
      Top             =   3165
      Width           =   2130
   End
   Begin VB.Image Image2 
      Height          =   2100
      Left            =   720
      Picture         =   "frmDiurnalNoctural.frx":1B938
      Top             =   840
      Width           =   5835
   End
End
Attribute VB_Name = "frmDiurnalNocturnal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()
Dim rex As RegExp
Set rex = New RegExp
rex.Pattern = "^([0-2]{1})([0-9]{1})\:([0-9]{2})\:([0-9]{2})$"
If Not rex.Test(Me.tStartDiurnal) Then
    MsgBox "La fecha de inicio de la jornada diurna no cumple con el formato 00:00:00", vbCritical
    Exit Sub
End If

If Not rex.Test(Me.tEndDiurnal) Then
    MsgBox "La fecha de inicio de la jornada diurna no cumple con el formato 00:00:00", vbCritical
    Exit Sub
End If

ModConfig.HOUR_START_D = Me.tStartDiurnal
ModConfig.HOUR_END_D = Me.tEndDiurnal
Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_HOUR_START_D, Me.tStartDiurnal)
Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_HOUR_END_D, Me.tEndDiurnal)

MsgBox "Se actualizó correctamente las horas de la jornada diurna/nocturna", vbInformation
End Sub

Private Sub Form_Load()
Me.tStartDiurnal = ModConfig.HOUR_START_D
Me.tEndDiurnal = ModConfig.HOUR_END_D
End Sub
