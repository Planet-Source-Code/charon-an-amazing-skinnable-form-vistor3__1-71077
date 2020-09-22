Attribute VB_Name = "modSubMain"
  '==========================================================================
  '                                                             This code is written by Charon (2008).
  '                                        If you have any problems using this Control, please contact me.
  '                                                         My E-mial Address: astrophsyics@126.com
  '==========================================================================
  
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2  '±íÊ¾°Ñ´°ÌåÉèÖÃ³É°ëÍ¸Ã÷ÑùÊ½  To make the form transparent
' »æÖÆÒõÓ°(Á¥ÊôSetClassLong)   Used to draw Shadow(Belongs to SetClassLong)
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = -26
Public Const WM_CLOSE = &H10      ' ¸Ã³£Á¿ÓÃÓÚ¹Ø±Õ´°¿Ú(Á¥ÊôSendMessage)   Used to close a Form(Belongs to SendMessage)
Public Const SW_MINIMIZE = 6        ' ¸Ã³£Á¿ÓÃÓÚ×îÐ¡»¯´°¿Ú(Á¥ÊôShowWindow)   Used to Minimize a Form(Belongs to ShowWindow)
Public Const SW_MAXIMIZE = 3
Public Const SW_NORMAL = 1

Private Sub Main()
' Æô¶¯²âÊÔ´°Ìå  Load frmDemo
    frmDemo.Show
End Sub

Public Function FillDeskV(Desk As Object, Source As PictureBox, Top As Integer, Left As Integer)
'ÓÃÈÎÒâÍ¼Æ¬×ÝÏòÌî³äÄ³Ò»ÈÝÆ÷    Use any picture to fill a container vertically
    Dim i As Integer
    For i = Top To Desk.Height Step Source.Height
            Desk.PaintPicture Source.Picture, Left, i, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Source.ScaleWidth, Source.ScaleHeight
    Next i
End Function

Public Function FillDeskH(Desk As Object, Source As PictureBox, Top As Integer, Left As Integer)
'ÓÃÈÎÒâÍ¼Æ¬ºáÏòÌî³äÄ³Ò»ÈÝÆ÷    Use any picture to fill a container horizontally
    Dim j As Integer
        For j = Left To Desk.Width Step Source.Width
            Desk.PaintPicture Source.Picture, j, Top, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Source.ScaleWidth, Source.ScaleHeight
        Next j
End Function

Public Function FillDesk(Desk As Object, Source As PictureBox)
'ÓÃÈÎÒâÍ¼Æ¬Ìî³äÄ³Ò»ÈÝÆ÷    Use any picture to fill a container
    Dim i As Integer, j As Integer
    For i = 0 To Desk.Height Step Source.Height
        For j = 0 To Desk.Width Step Source.Width
            Desk.PaintPicture Source.Picture, j, i, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Source.ScaleWidth, Source.ScaleHeight
        Next j
    Next i
End Function

Public Sub GetDirOp(ByVal Path$, Box As ComboBox)
'¶ÁÈ¡Ä³Ò»Ä¿Â¼ÏÂµÄ×ÓÎÄ¼þ¼Ð
    Dim vDirName As String, LastDir As String
    If Right(Path$, 1) <> "\" Then Path$ = Path$ & "\"
    vDirName = Dir(Path, vbDirectory) ' Retrieve the first entry.
    Do While Not vDirName = ""
        If vDirName <> "." And vDirName <> ".." Then
            If (GetAttr(Path & vDirName) And vbDirectory) = vbDirectory Then
                LastDir = vDirName
                ' vDirName
                Box.AddItem vDirName
                vDirName = Dir(Path$, vbDirectory)
                Do Until vDirName = LastDir Or vDirName = ""
                    vDirName = Dir
                Loop
                If vDirName = "" Then Exit Do
            End If
        End If
        vDirName = Dir
    Loop
End Sub
