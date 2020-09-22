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
Public Const LWA_ALPHA = &H2  '表示把窗体设置成半透明样式  To make the form transparent
' 绘制阴影(隶属SetClassLong)   Used to draw Shadow(Belongs to SetClassLong)
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = -26
Public Const WM_CLOSE = &H10      ' 该常量用于关闭窗口(隶属SendMessage)   Used to close a Form(Belongs to SendMessage)
Public Const SW_MINIMIZE = 6        ' 该常量用于最小化窗口(隶属ShowWindow)   Used to Minimize a Form(Belongs to ShowWindow)
Public Const SW_MAXIMIZE = 3
Public Const SW_NORMAL = 1

Private Sub Main()
' 启动测试窗体  Load frmDemo
    frmDemo.Show
End Sub

Public Function FillDeskV(Desk As Object, Source As PictureBox, Top As Integer, Left As Integer)
'用任意图片纵向填充某一容器    Use any picture to fill a container vertically
    Dim i As Integer
    For i = Top To Desk.Height Step Source.Height
            Desk.PaintPicture Source.Picture, Left, i, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Source.ScaleWidth, Source.ScaleHeight
    Next i
End Function

Public Function FillDeskH(Desk As Object, Source As PictureBox, Top As Integer, Left As Integer)
'用任意图片横向填充某一容器    Use any picture to fill a container horizontally
    Dim j As Integer
        For j = Left To Desk.Width Step Source.Width
            Desk.PaintPicture Source.Picture, j, Top, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Source.ScaleWidth, Source.ScaleHeight
        Next j
End Function

Public Function FillDesk(Desk As Object, Source As PictureBox)
'用任意图片填充某一容器    Use any picture to fill a container
    Dim i As Integer, j As Integer
    For i = 0 To Desk.Height Step Source.Height
        For j = 0 To Desk.Width Step Source.Width
            Desk.PaintPicture Source.Picture, j, i, Source.ScaleWidth, Source.ScaleHeight, 0, 0, Source.ScaleWidth, Source.ScaleHeight
        Next j
    Next i
End Function

Public Sub GetDirOp(ByVal Path$, Box As ComboBox)
'读取某一目录下的子文件夹
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
