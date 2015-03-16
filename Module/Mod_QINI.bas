Attribute VB_Name = "Mod_QINI"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub WriteText(IniFile As String, appName As String, keyName As String, valueNew As String)
    Dim X As Long
    X = WritePrivateProfileString(appName, keyName, valueNew, IniFile)
End Sub

Public Function GetText(IniFile As String, appName As String, keyName As String) As String
    Dim strDefault As String
    Dim lngBuffLen As Long
    Dim strResu As String
    Dim X As Long

    strResu = String$(1025, vbNullChar): lngBuffLen = 1025
    strDefault = ""
    X = GetPrivateProfileString(appName, keyName, _
                                strDefault, strResu, lngBuffLen, IniFile)
    GetText = Left$(strResu, X)
End Function

