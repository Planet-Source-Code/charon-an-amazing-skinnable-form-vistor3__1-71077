Attribute VB_Name = "modResize"
Option Explicit

Public OldWindowProc As Long, procOld As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Type POINTAPI
  x As Long
  y As Long
End Type
Public Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type
Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = -4
Global lpPrevWndProc As Long
Public udtMMI As MINMAXINFO
Global MinSizeY As Long, MinSizeX As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case iMsg
Case WM_GETMINMAXINFO
Dim udtMINMAXINFO As MINMAXINFO
CopyMemory udtMINMAXINFO, ByVal lParam, 40&
With udtMINMAXINFO
    .ptMaxSize.x = udtMMI.ptMaxSize.x
    .ptMaxSize.y = udtMMI.ptMaxSize.y
    .ptMaxPosition.x = 0
    .ptMaxPosition.y = 0
    .ptMaxTrackSize.x = .ptMaxSize.x
    .ptMaxTrackSize.y = .ptMaxSize.y
    .ptMinTrackSize.x = udtMMI.ptMinTrackSize.x
    .ptMinTrackSize.y = udtMMI.ptMinTrackSize.y
End With
CopyMemory ByVal lParam, udtMINMAXINFO, 40&
WindowProc = False
Exit Function
End Select
WindowProc = CallWindowProc(procOld, hwnd, iMsg, wParam, lParam)
End Function

Public Function LockWindow(hwnd As Long, Optional MinWidth As Long, Optional MinHeight As Long, Optional maxwidth As Long, Optional maxheight As Long) As Boolean
With udtMMI
'Ö¸¶¨´°Ìå×î´ó¿í¶È
If maxwidth = 0 Then .ptMaxSize.x = Screen.Width \ Screen.TwipsPerPixelX Else .ptMaxSize.x = maxwidth
'Ö¸¶¨´°Ìå×î´ó¸ß¶È
If maxheight = 0 Then .ptMaxSize.y = Screen.Width \ Screen.TwipsPerPixelX Else .ptMaxSize.y = maxheight
End With
procOld = SetWindowLong(hwnd, -4, AddressOf WindowProc)
End Function

Public Function WndMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long

' Watch for the pertinent message to come in
  If Msg = WM_GETMINMAXINFO Then
    Dim MinMax As MINMAXINFO
' ¡¡This is necessary because the structure was passed by its address and there
' ¡¡is currently no intrinsic way to use an address in Visual Basic

    CopyMemory MinMax, ByVal lp, Len(MinMax)
    
    MinMax.ptMinTrackSize.x = MinSizeX '740
    MinMax.ptMinTrackSize.y = MinSizeY '400
    MinMax.ptMaxTrackSize.x = Screen.Width \ Screen.TwipsPerPixelX
    MinMax.ptMaxTrackSize.y = Screen.Height \ Screen.TwipsPerPixelY

' Here we copy the datastructure back up to the address passed in the parameters
' because Windows will look there for the information.

    CopyMemory ByVal lp, MinMax, Len(MinMax)

' This message tells Windows that the message was handled successfully

    WndMessage = 1
    Exit Function

  End If

' Here, we forward all irrelevant messages on to the default message handler.
WndMessage = CallWindowProc(OldWindowProc, hwnd, Msg, wp, lp)

End Function
