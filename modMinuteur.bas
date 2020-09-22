Attribute VB_Name = "modMinuteur"
Option Explicit

'APIs
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function AlphaBlend Lib "MSIMG32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
'
'DECLARATION DE LA COLLECTION QUI VA STOCKER L'ID DE CHAQUE CLASSE QUI APPELLERA UN MINUTEUR
Public IDClsMinuteur As New Collection

Public Sub AjoutColl(IDClasse As Long, IDMinuteur As Long)
    
    IDClsMinuteur.Add IDClasse, "M" & IDMinuteur

End Sub

Public Sub EnleveColl(IDClasse As Long)
    Dim i As Integer
    
    For i = 1 To IDClsMinuteur.Count
        If Str(IDClsMinuteur.Item(i)) = Str(IDClasse) Then
            IDClsMinuteur.Remove i
            Exit For
        End If
    Next

End Sub

Public Sub MinuteurProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal iEvent As Long, ByVal iTime As Long)

On Error Resume Next

Dim cM As Minuteur
Dim hLng As Long

    hLng = CLng(IDClsMinuteur.Item("M" & iEvent))
    If hLng = 0 Then Exit Sub
    CopyMemory cM, hLng, 4&
    cM.LancementAction
    CopyMemory cM, 0&, 4
    
If err Then Exit Sub
    
End Sub
