VERSION 5.00
Begin VB.Form frmConfigExcel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración plantilla Excel"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   19785
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   19785
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tColRf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   18285
      TabIndex        =   12
      Text            =   "16"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColRnf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   17250
      TabIndex        =   11
      Text            =   "15"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColRn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   16080
      TabIndex        =   10
      Text            =   "14"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColHenf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   15000
      TabIndex        =   9
      Text            =   "13"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColHedf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   13860
      TabIndex        =   8
      Text            =   "12"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColHeno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   12705
      TabIndex        =   7
      Text            =   "11"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColHedo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   11595
      TabIndex        =   6
      Text            =   "10"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColHEnd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   9315
      TabIndex        =   5
      Text            =   "8"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColHStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   8220
      TabIndex        =   4
      Text            =   "7"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   7080
      TabIndex        =   3
      Text            =   "6"
      Top             =   3060
      Width           =   510
   End
   Begin VB.TextBox tColType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H001BE7E1&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   360
      Left            =   5940
      TabIndex        =   2
      Text            =   "5"
      Top             =   3060
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   360
      Picture         =   "frmConfigExcel.frx":0000
      ScaleHeight     =   5055
      ScaleWidth      =   19095
      TabIndex        =   0
      Top             =   1560
      Width           =   19095
      Begin VB.TextBox tColTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001BE7E1&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   360
         Left            =   10095
         TabIndex        =   13
         Text            =   "9"
         Top             =   1500
         Width           =   510
      End
      Begin VB.TextBox tRowStartRead 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001BE7E1&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   360
         Left            =   0
         TabIndex        =   1
         Text            =   "9"
         Top             =   3105
         Width           =   510
      End
   End
   Begin VB.Image btnSave 
      Height          =   450
      Left            =   7560
      Picture         =   "frmConfigExcel.frx":12A1E2
      Top             =   7200
      Width           =   2250
   End
   Begin VB.Image btnCancel 
      Height          =   450
      Left            =   10080
      Picture         =   "frmConfigExcel.frx":12D71C
      Top             =   7200
      Width           =   2250
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   360
      X2              =   18360
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   -195
      Picture         =   "frmConfigExcel.frx":130C56
      Top             =   225
      Width           =   8700
   End
End
Attribute VB_Name = "frmConfigExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
If (isValidField(Me.tRowStartRead)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_ROW_START_READ, Me.tRowStartRead.Text)
End If
If (isValidField(Me.tColType)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_TYPE_ROW, Me.tColType.Text)
End If
If (isValidField(Me.tColDate)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_DATE, Me.tColDate.Text)
End If
If (isValidField(Me.tColHStart)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_HOUR_INI, Me.tColHStart.Text)
End If
If (isValidField(Me.tColHEnd)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_HOUR_END, Me.tColHEnd.Text)
End If
If (isValidField(Me.tColHedo)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_HEDO, Me.tColHedo.Text)
End If
If (isValidField(Me.tColHeno)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_HENO, Me.tColHeno.Text)
End If
If (isValidField(Me.tColHedf)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_HEDF, Me.tColHedf.Text)
End If
If (isValidField(Me.tColHenf)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_HENF, Me.tColHenf.Text)
End If
If (isValidField(Me.tColRn)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_RN, Me.tColRn.Text)
End If
If (isValidField(Me.tColRnf)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_RNF, Me.tColRnf.Text)
End If
If (isValidField(Me.tColRf)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_RF, Me.tColRf.Text)
End If
If (isValidField(Me.tColTotal)) Then
    Call ModIni.savePropertyFile(ModConfig.calendarPath, ModIni.K_COL_TOT, Me.tColTotal.Text)
End If

Call ModConfig.loadExcelConfig
MsgBox "Configuración Actualizada", vbInformation
End Sub

Private Function isValidField(txt As TextBox)
If (txt.Text <> "" And IsNumeric(txt.Text)) Then
    isValidField = True
Else
    isValidField = False
End If
End Function

Private Sub Form_Load()
Me.tRowStartRead = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_ROW_START_READ, 9)
Me.tColType = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_TYPE_ROW, 5)
Me.tColDate = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_DATE, 6)
Me.tColTotal = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_TOT, 9)
Me.tColHStart = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_HOUR_INI, 7)
Me.tColHEnd = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_HOUR_END, 8)
Me.tColHedo = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_HEDO, 10)
Me.tColHeno = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_HENO, 11)
Me.tColHedf = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_HEDF, 12)
Me.tColHenf = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_HENF, 13)
Me.tColRn = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_RN, 14)
Me.tColRnf = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_RNF, 15)
Me.tColRf = ModIni.readPropertyFile(ModConfig.calendarPath, ModIni.K_COL_RF, 16)
End Sub
