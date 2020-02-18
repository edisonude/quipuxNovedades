VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportExtraHours 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de horas extras"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReportExtraHours.frx":0000
   ScaleHeight     =   9285
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tResult 
      Height          =   5460
      Left            =   9765
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   1785
      Width           =   3120
   End
   Begin VB.PictureBox picProcessing 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   1020
      Picture         =   "frmReportExtraHours.frx":1D322
      ScaleHeight     =   4800
      ScaleWidth      =   8220
      TabIndex        =   0
      Top             =   8865
      Visible         =   0   'False
      Width           =   8220
   End
   Begin VB.TextBox tColName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   3285
      TabIndex        =   35
      Text            =   "L"
      Top             =   5595
      Width           =   405
   End
   Begin VB.TextBox tColCharge 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   3285
      TabIndex        =   34
      Text            =   "O"
      Top             =   6420
      Width           =   405
   End
   Begin VB.TextBox tColCell 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   3285
      TabIndex        =   32
      Text            =   "N"
      Top             =   6015
      Width           =   405
   End
   Begin VB.TextBox tColIdLeader 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   8985
      TabIndex        =   28
      Text            =   "D"
      Top             =   5190
      Width           =   405
   End
   Begin VB.TextBox tColNameLeader 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   8970
      TabIndex        =   27
      Text            =   "E"
      Top             =   5625
      Width           =   405
   End
   Begin VB.TextBox tColCellLeader 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   8985
      TabIndex        =   26
      Text            =   "F"
      Top             =   6060
      Width           =   405
   End
   Begin VB.TextBox tMaxMonth 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   7920
      TabIndex        =   25
      Text            =   "48"
      Top             =   4400
      Width           =   525
   End
   Begin VB.TextBox tMaxWeek 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   5640
      TabIndex        =   24
      Text            =   "12"
      Top             =   4400
      Width           =   525
   End
   Begin VB.TextBox tMaxDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Left            =   3600
      TabIndex        =   23
      Text            =   "2"
      Top             =   4400
      Width           =   525
   End
   Begin VB.TextBox tRowStart 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   3285
      TabIndex        =   18
      Text            =   "8"
      Top             =   6840
      Width           =   405
   End
   Begin VB.TextBox tColTotHours 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   6135
      TabIndex        =   16
      Text            =   "AB"
      Top             =   6465
      Width           =   405
   End
   Begin VB.TextBox tColEndHour 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   6135
      TabIndex        =   14
      Text            =   "S"
      Top             =   6060
      Width           =   405
   End
   Begin VB.TextBox tColHourStart 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   6135
      TabIndex        =   12
      Text            =   "R"
      Top             =   5625
      Width           =   405
   End
   Begin VB.TextBox tColDate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   8985
      TabIndex        =   10
      Text            =   "Q"
      Top             =   6480
      Width           =   405
   End
   Begin VB.TextBox tColType 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   6135
      TabIndex        =   8
      Text            =   "P"
      Top             =   5190
      Width           =   405
   End
   Begin VB.TextBox tColId 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   3270
      TabIndex        =   7
      Text            =   "K"
      Top             =   5190
      Width           =   405
   End
   Begin VB.ComboBox cmbMonths 
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
      Height          =   405
      ItemData        =   "frmReportExtraHours.frx":9CEFC
      Left            =   840
      List            =   "frmReportExtraHours.frx":9CF24
      TabIndex        =   3
      Top             =   4380
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog excelDialog 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccione el archivo de novedades a procesar"
      Filter          =   "Archivos Excel (xlsx)|*.xlsx|Archivos Excel (xls)|*.xls"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. nombre empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   13
      Left            =   1020
      TabIndex        =   37
      Top             =   5610
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. cargo empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   12
      Left            =   1020
      TabIndex        =   36
      Top             =   6435
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. célula empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   11
      Left            =   1020
      TabIndex        =   33
      Top             =   6030
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. cédula líder"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   10
      Left            =   6720
      TabIndex        =   31
      Top             =   5205
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. nombre líder"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   9
      Left            =   6720
      TabIndex        =   30
      Top             =   5640
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. célula líder"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   8
      Left            =   6720
      TabIndex        =   29
      Top             =   6075
      Width           =   2220
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max. Horas x mes"
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
      Height          =   285
      Index           =   3
      Left            =   7200
      TabIndex        =   22
      Top             =   4080
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max. Horas x semana"
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
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   21
      Top             =   4080
      Width           =   2235
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max. Horas x dia"
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
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   20
      Top             =   4080
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fila inicio lectura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   7
      Left            =   1020
      TabIndex        =   19
      Top             =   6855
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. Total de horas #"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   6
      Left            =   3870
      TabIndex        =   17
      Top             =   6480
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. Hora de fin"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   5
      Left            =   3870
      TabIndex        =   15
      Top             =   6075
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. Hora de inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   4
      Left            =   3870
      TabIndex        =   13
      Top             =   5640
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. Fecha del reporte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   3
      Left            =   6720
      TabIndex        =   11
      Top             =   6495
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. Tipo de novedad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   2
      Left            =   3870
      TabIndex        =   9
      Top             =   5205
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Col. cédula empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Index           =   1
      Left            =   1020
      TabIndex        =   6
      Top             =   5205
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   3
      X1              =   780
      X2              =   9420
      Y1              =   4830
      Y2              =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Configuración de columnas:"
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
      Height          =   285
      Index           =   0
      Left            =   780
      TabIndex        =   5
      Top             =   4860
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mes a procesar"
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
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   4065
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H001BE7E1&
      BorderWidth     =   4
      Index           =   2
      X1              =   480
      X2              =   9570
      Y1              =   3975
      Y2              =   3960
   End
   Begin VB.Image btnProcess 
      Height          =   855
      Left            =   2850
      Picture         =   "frmReportExtraHours.frx":9CF8E
      Top             =   2760
      Width           =   4680
   End
   Begin VB.Image iMenuNovedades 
      Height          =   1065
      Left            =   4110
      Picture         =   "frmReportExtraHours.frx":AA038
      Top             =   7440
      Width           =   1905
   End
   Begin VB.Label lExcelFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione el archivo de horas extras"
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
      Left            =   2415
      TabIndex        =   2
      Top             =   2115
      Width           =   4935
   End
   Begin VB.Label lUpload 
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   7470
      TabIndex        =   1
      Top             =   1980
      Width           =   660
   End
   Begin VB.Image btnEnd 
      Height          =   630
      Left            =   2910
      Picture         =   "frmReportExtraHours.frx":B0AFA
      Top             =   8565
      Width           =   4710
   End
   Begin VB.Line Line1 
      BorderColor     =   &H001BE7E1&
      BorderWidth     =   4
      Index           =   0
      X1              =   495
      X2              =   9705
      Y1              =   7320
      Y2              =   7305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H001BE7E1&
      BorderWidth     =   4
      Index           =   1
      X1              =   480
      X2              =   9645
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   285
      Picture         =   "frmReportExtraHours.frx":BA61C
      Top             =   1125
      Width           =   7005
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   2295
      Picture         =   "frmReportExtraHours.frx":CC38E
      Top             =   1950
      Width           =   5865
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   480
      Picture         =   "frmReportExtraHours.frx":D8DF0
      Top             =   3600
      Width           =   2340
   End
End
Attribute VB_Name = "frmReportExtraHours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public excelPath As String
Public excelApp As Excel.APPLICATION
Public workbook As Excel.workbook
Public sheet As Excel.Worksheet

Dim colId, colName, colCharge, colCell As Integer
Dim colLeaderId, colLeaderName, colLeaderCell As Integer
Dim colDate, colHourStart, colHourEnd, colType, colTotHours, rowStart As Integer

Dim idsToReport() As String
Dim countToReport As Integer
Dim infoToReport()


Private Function loadInfoCols()
colId = ModConfig.convertExcelColToInt(Me.tColId)
colName = ModConfig.convertExcelColToInt(Me.tColName)
colCharge = ModConfig.convertExcelColToInt(Me.tColCharge)
colCell = ModConfig.convertExcelColToInt(Me.tColCell)

colLeaderId = ModConfig.convertExcelColToInt(Me.tColIdLeader)
colLeaderName = ModConfig.convertExcelColToInt(Me.tColNameLeader)
colLeaderCell = ModConfig.convertExcelColToInt(Me.tColCellLeader)

colDate = ModConfig.convertExcelColToInt(Me.tColDate)
colHourStart = ModConfig.convertExcelColToInt(Me.tColHourStart)
colHourEnd = ModConfig.convertExcelColToInt(Me.tColEndHour)
colType = ModConfig.convertExcelColToInt(Me.tColType)
colTotHours = ModConfig.convertExcelColToInt(Me.tColTotHours)
rowStart = Val(Me.tRowStart)
End Function

Private Sub btnEnd_Click()
End
End Sub

Private Sub btnProcess_Click()
On Error GoTo closeResources

If (Me.excelPath = "") Then
    MsgBox "Debe seleccionar el archivo para generar el reporte de horas extras.", vbCritical
    Exit Sub
End If

Call loadInfoCols

Call showProcessing
Call loadExcel

countToReport = 0

Dim hasMoreRows As Boolean
Dim row As Integer
Dim rowsProcessed As Integer

hasMoreRows = True
row = rowStart
rowsProcessed = 0

While hasMoreRows
    Value = sheet.Cells(row, colType)
    If (Value = "") Then
        hasMoreRows = False
    ElseIf (shouldProcess(row)) Then
        Dim he As Double
        Dim document As String
        Dim dateProcessed As Date
    
        document = sheet.Cells(row, colId)
        dateProcessed = sheet.Cells(row, colDate)
        he = getHe(row, document, dateProcessed)
        
        If (he > Val(Me.tMaxDay)) Then
            Call addRowToReport(row, he, Val(Me.tMaxDay))
        End If
        
        
        ' MsgBox "total horas doc: " & document & " fehca = " & dateProcessed & " tot = " & he
    
    
'        Dim dateStart As Date
'        Dim dateEnd As Date
'        Dim dateReference As Date
'        Dim difGeneral As Integer
'        Dim dif As Integer
'        Dim isDiurnal As Boolean
'
'        hedo = 0
'        heno = 0
'        hedf = 0
'        henf = 0
'
'        Dim dateStartDiurnal As Date
'        Dim dateEndDiurnal As Date
'
'        dateStart = getDateStart(row)
'        dateEnd = getDateEnd(row, dateStart)
'
'        dateReference = dateStart
'        difGeneral = DateDiff("n", dateReference, dateEnd)
'        totMins = difGeneral
'
'        While difGeneral > 0
'            dateStartDiurnal = getDateStartDiurnal(dateReference)
'continueDiffStartDate:
'            dif = DateDiff("n", dateReference, dateStartDiurnal)
'            isDiurnal = False
'
'            If (dif <= 0) Then 'TODO: que pasa si es igual
'                dateEndDiurnal = getDateEndDiurnal(dateReference)
'                dif = DateDiff("n", dateReference, dateEndDiurnal)
'                isDiurnal = True
'
'                If (dif <= 0) Then
'                    dateStartDiurnal = getDateStartDiurnal(DateAdd("d", 1, dateReference))
'                    GoTo continueDiffStartDate
'                Else
'                    If (dif <= difGeneral) Then
'                        Call assignDiference(dif, isDiurnal, dateReference, dateEndDiurnal)
'                        dateReference = dateEndDiurnal
'                        difGeneral = difGeneral - dif
'                    Else
'                        Call assignDiference(difGeneral, isDiurnal, dateReference, dateEndDiurnal)
'                        difGeneral = difGeneral - difGeneral
'                    End If
'                End If
'            Else
'                If (dif <= difGeneral) Then
'                    Call assignDiference(dif, isDiurnal, dateReference, dateStartDiurnal)
'                    dateReference = dateStartDiurnal
'                    difGeneral = difGeneral - dif
'                Else
'                    Call assignDiference(difGeneral, isDiurnal, dateReference, dateStartDiurnal)
'                    difGeneral = difGeneral - difGeneral
'                End If
'            End If
'        Wend
'        rowsProcessed = rowsProcessed + 1
'        Call writeResults(row)
    End If
    Me.tResult = "fila " & row & " - procesada" & vbNewLine & Me.tResult
    DoEvents
    row = row + 1
Wend

'Escribir resultados
Dim wsResult As Excel.Worksheet
Set wsResult = workbook.Worksheets.Add(Type:=xlWorksheet)
With wsResult
    .name = "ReportePorDia"
    Dim rowsResult As Integer
    
    'Escribir los encabezados
    Dim rowHeader As Integer
    rowHeader = rowStart - 1
    .Cells(1, 1) = sheet.Cells(rowHeader, colLeaderId)
    .Cells(1, 2) = sheet.Cells(rowHeader, colLeaderName)
    .Cells(1, 3) = sheet.Cells(rowHeader, colLeaderCell)
    .Cells(1, 4) = sheet.Cells(rowHeader, colId)
    .Cells(1, 5) = sheet.Cells(rowHeader, colName)
    .Cells(1, 6) = sheet.Cells(rowHeader, colCharge)
    .Cells(1, 7) = sheet.Cells(rowHeader, colCell)
    .Cells(1, 8) = sheet.Cells(rowHeader, colDate)
    .Cells(1, 9) = "H. EXTRAS REPORTADAS"
    .Cells(1, 10) = "H. EXTRAS EXCEDIDAS"
    
    For j = 1 To 10
        .Cells(1, j).Interior.Color = vbBlack
        .Cells(1, j).Font.Color = vbWhite
    Next
    .Columns("A:J").AutoFit
    
    Dim rowResult As Integer
    For rowResult = 1 To UBound(infoToReport)
        .Cells(rowResult + 1, 1) = infoToReport(rowResult).leaderId
        .Cells(rowResult + 1, 2) = infoToReport(rowResult).leaderName
        .Cells(rowResult + 1, 3) = infoToReport(rowResult).leaderCell
        .Cells(rowResult + 1, 4) = infoToReport(rowResult).id
        .Cells(rowResult + 1, 5) = infoToReport(rowResult).name
        .Cells(rowResult + 1, 6) = infoToReport(rowResult).charge
        .Cells(rowResult + 1, 7) = infoToReport(rowResult).cell
        .Cells(rowResult + 1, 8) = infoToReport(rowResult).dateHe
        .Cells(rowResult + 1, 9) = infoToReport(rowResult).hours
        .Cells(rowResult + 1, 10) = infoToReport(rowResult).hoursExceed
    Next
End With

workbook.Save
workbook.Close SaveChanges:=False

Call showProcessing

MsgBox "Finalizó la generación del reporte de horas extras.", vbInformation

closeResources:
If (Err.Number <> 0) Then
    MsgBox Err.Description
End If
Call closeResources
End Sub

Private Function addRowToReport(currentRow As Integer, hours As Double, maxHours As Integer)
countToReport = countToReport + 1
ReDim Preserve idsToReport(1 To countToReport)
idsToReport(countToReport) = getKeyRow(currentRow)

Dim entryReport As CEntryReportHe
Set entryReport = New CEntryReportHe
entryReport.leaderId = sheet.Cells(currentRow, colLeaderId)
entryReport.leaderName = sheet.Cells(currentRow, colLeaderName)
entryReport.leaderCell = sheet.Cells(currentRow, colLeaderCell)
entryReport.id = sheet.Cells(currentRow, colId)
entryReport.name = sheet.Cells(currentRow, colName)
entryReport.charge = sheet.Cells(currentRow, colCharge)
entryReport.cell = sheet.Cells(currentRow, colCell)
entryReport.dateHe = sheet.Cells(currentRow, colDate)
entryReport.hours = hours
entryReport.hoursExceed = hours - maxHours

ReDim Preserve infoToReport(1 To countToReport)
Set infoToReport(countToReport) = entryReport
End Function

Private Function getKeyRow(row As Integer) As String
getKeyRow = sheet.Cells(row, colId) & "-" & Format(sheet.Cells(row, colDate), "dd/MM/yyyy")
End Function

Private Function getHe(currentRow, document As String, currentDate As Date) As Double
Dim hasMoreRows As Boolean
hasMoreRows = True
Dim row As Integer
row = currentRow + 1
Dim heTotal As Double
heTotal = sheet.Cells(currentRow, colTotHours)

While hasMoreRows
    Value = sheet.Cells(row, colType)
    If (Value = "") Then
        hasMoreRows = False
    ElseIf (shouldProcess(row) And currentRow <> row) Then
        Dim documentReport As String
        Dim dateReport As Date
        documentReport = sheet.Cells(row, colId)
        dateReport = sheet.Cells(row, colDate)
        If (dateReport > currentDate) Then
                hasMoreRows = False
            Else
            If (document = documentReport And currentDate = dateReport) Then
                heTotal = heTotal + sheet.Cells(row, colTotHours)
            End If
        End If
    End If
    row = row + 1
Wend
getHe = heTotal
End Function

Private Function shouldProcess(row As Integer) As Boolean
Dim typeRow As String
Dim dateRow As Date
typeRow = sheet.Cells(row, colType)
dateRow = sheet.Cells(row, colDate)
shouldProcess = "HORA EXTRA" = typeRow And Month(dateRow) = Me.cmbMonths.ListIndex + 1 And Not keyAlreadyExist(row)
End Function

Private Function keyAlreadyExist(row As Integer) As Boolean
On Error GoTo returnFalse
Dim Key As String
Key = getKeyRow(row)
Dim i As Integer
For i = 1 To UBound(idsToReport)
    If (Key = idsToReport(i)) Then
        keyAlreadyExist = True
        Exit Function
    End If
Next
returnFalse:
keyAlreadyExist = False
End Function

Private Function loadExcel()
Set excelApp = New Excel.APPLICATION
Set workbook = excelApp.Workbooks.Open(FileName:=excelPath)
Set sheet = workbook.Sheets(1)
End Function

Private Function closeResources()
Set workbook = Nothing
If Not excelApp Is Nothing Then
    excelApp.Quit
    Set excelApp = Nothing
End If
End Function

Private Sub cmbMonths_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
Me.cmbMonths.ListIndex = Month(Now) - 1
End Sub

Private Sub iMenuNovedades_Click()
frmProcess.Show
Unload Me
End Sub

Private Sub lUpload_Click()
excelDialog.ShowOpen
If excelDialog.FileName <> "" Then
    excelPath = excelDialog.FileName
    Me.lExcelFile = excelDialog.FileTitle
Else
    excelPath = ""
    Me.lExcelFile = "Seleccione el archivo de horas extras"
End If
End Sub

Private Sub showProcessing()
If (Me.picProcessing.Visible = False) Then
    Me.picProcessing.Visible = True
    Me.picProcessing.Top = 1680
Else
    Me.picProcessing.Visible = False
End If
End Sub
