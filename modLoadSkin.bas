Attribute VB_Name = "modLoadSkin"
  '==========================================================================
  '                                                             This code is written by Charon (2008).
  '                                        If you have any problems using this Control, please contact me.
  '                                                         My E-mial Address: astrophsyics@126.com
  '==========================================================================
  
Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

'读取INI文件  Read an INI File
Function GetINI(AppName As String, KeyName As String, FileName As String) As String
    Dim RetStr As String
    RetStr = String(10000, Chr(0))
    GetINI = Left(RetStr, GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), FileName))
End Function

'写入NI文件  Write an INI File
Function SaveINI(AppName As String, KeyName As String, value As String, FileName As String)
     WritePrivateProfileString AppName, KeyName, value, FileName
End Function

'从文件路径得到文件名  Get the filename from a specific path
Public Function StripPath(t$) As String
    Dim x%, ct%
    StripPath$ = t$
    x% = InStr(t$, "\")
    Do While x%
        ct% = x%
        x% = InStr(ct% + 1, t$, "\")
    Loop
    If ct% > 0 Then StripPath$ = Mid$(t$, ct% + 1)
End Function

'从文件路径得到目录 Get the folder path from a specific file path
Public Function GetFolderPath(FilePath As String) As String
Dim tempPath As String
    tempPath = StripPath(FilePath)
    GetFolderPath = Left(FilePath, Len(FilePath) - Len(tempPath))
End Function
