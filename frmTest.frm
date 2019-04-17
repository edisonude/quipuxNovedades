VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Add a reference to MS Excel xx.0 Object Library
Private Sub Command1_Click()
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

