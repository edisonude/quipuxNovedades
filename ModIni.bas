Attribute VB_Name = "ModIni"
Option Explicit

Const APPLICATION As String = "Quipux"
Public Const K_ROW_START_READ As String = "R_SR"
Public Const K_COL_TYPE_ROW As String = "C_TR"
Public Const K_COL_DATE As String = "C_D"
Public Const K_COL_HOUR_INI As String = "C_HS"
Public Const K_COL_HOUR_END As String = "C_HE"
Public Const K_COL_HEDO As String = "C_HEDO"
Public Const K_COL_HENO As String = "C_HENO"
Public Const K_COL_HEDF As String = "C_HEDF"
Public Const K_COL_HENF As String = "C_HENF"
Public Const K_COL_RN As String = "C_RN"
Public Const K_COL_RNF As String = "C_RNF"
Public Const K_COL_RF As String = "C_RF"
Public Const K_HOUR_START_D  As String = "HSD"
Public Const K_HOUR_END_D As String = "HED"

'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
    
'Almacena la ruta del archivo de configuraciones
Dim fileConfigPath As String

'Lee un dato _
-----------------------------
'Recibe la ruta del archivo, la clave a leer y _
 el valor por defecto en caso de que la Key no exista
Public Function readPropertyFile(Path_INI As String, Key As String, default As Variant) As String

Dim bufer As String * 256
Dim Len_Value As Long

        Len_Value = GetPrivateProfileString(APPLICATION, _
                                         Key, _
                                         default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
        
        readPropertyFile = Left$(bufer, Len_Value)

End Function

'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave

Public Function savePropertyFile(Path_INI As String, Key As String, Valor As Variant) As String

    WritePrivateProfileString APPLICATION, _
                                         Key, _
                                         Valor, _
                                         Path_INI

End Function


