VERSION 5.00
Begin VB.UserControl ctlVistorForm 
   Alignable       =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9525
   ControlContainer=   -1  'True
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   Begin VB.PictureBox picFormBtn 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   7815
      Picture         =   "ctlVistorForm.ctx":0000
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   95
      TabIndex        =   12
      Top             =   0
      Width           =   1425
      Begin Vistor3.ctlVistorButton cmdMinBtn 
         Height          =   270
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   476
         ImgUp           =   "ctlVistorForm.ctx":0BC2
         ImgDown         =   "ctlVistorForm.ctx":11B4
         ImgOver         =   "ctlVistorForm.ctx":17A6
         Caption         =   ""
         UseAlphaBlend   =   -1  'True
      End
      Begin Vistor3.ctlVistorButton cmdCloseBtn 
         Height          =   270
         Left            =   765
         TabIndex        =   14
         Top             =   0
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   476
         ImgUp           =   "ctlVistorForm.ctx":1D98
         ImgDown         =   "ctlVistorForm.ctx":2732
         ImgOver         =   "ctlVistorForm.ctx":30CC
         Caption         =   ""
         UseAlphaBlend   =   -1  'True
      End
      Begin Vistor3.ctlVistorButton cmdMaxBtn 
         Height          =   270
         Left            =   390
         TabIndex        =   15
         Top             =   0
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   476
         ImgUp           =   "ctlVistorForm.ctx":3A66
         ImgDown         =   "ctlVistorForm.ctx":4058
         ImgOver         =   "ctlVistorForm.ctx":464A
         Caption         =   ""
         UseAlphaBlend   =   -1  'True
      End
   End
   Begin VB.PictureBox picHead_1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "ctlVistorForm.ctx":4C3C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   213
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.PictureBox picBottom_1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   720
      Picture         =   "ctlVistorForm.ctx":977E
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox picBottom_3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   1200
      Picture         =   "ctlVistorForm.ctx":A6C0
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.PictureBox picBottom_2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   5400
      Picture         =   "ctlVistorForm.ctx":C382
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picRight_3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   960
      Picture         =   "ctlVistorForm.ctx":C464
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picLeft_3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   720
      Picture         =   "ctlVistorForm.ctx":CABE
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picRight_2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   720
      Picture         =   "ctlVistorForm.ctx":D118
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picRight_1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   360
      Picture         =   "ctlVistorForm.ctx":D1D2
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picLeft_1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   120
      Picture         =   "ctlVistorForm.ctx":D8A4
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picLeft_2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   960
      Picture         =   "ctlVistorForm.ctx":DF76
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picHead_2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3720
      Picture         =   "ctlVistorForm.ctx":E018
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picHead_3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   4560
      Picture         =   "ctlVistorForm.ctx":E3A2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.Image imgIcon 
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vistor Skinnable Form"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   480
      TabIndex        =   16
      Top             =   240
      Width           =   1845
   End
End
Attribute VB_Name = "ctlVistorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '==========================================================================
  '                                                           This Control is created by Charon (2008).
  '                                        If you have any problems using this Control, please contact me.
  '                                                         My E-mial Address: astrophsyics@126.com
  '==========================================================================
  
  Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
  Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
  Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
  Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
  Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
  Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
  Private Declare Function GetDesktopWindow Lib "user32" () As Long
  Private hBrush  As Long, outRgn As Long, myfont  As StdFont, userFrmBtnT As Integer, userFrmBtnR As Integer
  Private bCloseOnly As Boolean, bMax As Boolean, bEnMax As Boolean, userRgn1 As Integer, userRgn2 As Integer
  Private bDraggable As Boolean, bShadow As Boolean, intSizerBottom As Integer, intSizerRight As Integer, intTransparency As Integer
  Private Const GWL_WNDPROC = (-4)
  Private Type RECT
                  Left   As Long
                  Top   As Long
                  Right   As Long
                  Bottom   As Long
  End Type
  
  Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Event DblClick()
  Event Click()
  Event FormBtnClick(FormButton As Integer)   'FormButton:  1-Minimize  2-Maximize  3-Close
  Event SkinChanged() 'µ±Íê³É¼ÓÔØÆ¤·ôÊ±±»¼¤·¢  Raise when a skin is loaded

Private Sub DoDrag()
On Error Resume Next
'ÍÏ×§´°Ìå  Drag the Form
    ReleaseCapture
    If Not bDraggable Then Exit Sub
    SendMessage UserControl.ContainerHwnd, &HA1, 2, 0&
End Sub

Private Sub DoMaximize()
'´°Ìå×î´ó»¯  Maximize the form
If bCloseOnly Or bEnMax = False Then Exit Sub
If Not bMax Then
        Call SetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC, procOld)
        Dim hTaskBar As Long
        hTaskBar = FindWindow("Shell_TrayWnd", vbNullString)
        Dim RC As RECT
        Dim i As Long
        i = GetWindowRect(hTaskBar, RC)
        Dim taskheight As Long
        taskheight = RC.Bottom - RC.Top 'ÈÎÎñÀ¸¸ß¶È Taskbar Height
        i = GetWindowRect(GetDesktopWindow, RC)
        Dim maxwidth As Long
        Dim maxheight As Long
        maxwidth = RC.Right - RC.Left '»ñÈ¡ÆÁÄ»¿í¶È Screen Width
        maxheight = RC.Bottom - RC.Top - taskheight
        LockWindow UserControl.ContainerHwnd, , , maxwidth, maxheight
    Call ShowWindow(UserControl.ContainerHwnd, SW_MAXIMIZE)
    bMax = True: bDraggable = False
Else
    Call ShowWindow(UserControl.ContainerHwnd, SW_NORMAL)
    OldWindowProc = GetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC)
    Call SetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC, AddressOf WndMessage)
    bMax = False: bDraggable = True
End If
Call UserControl_Resize
End Sub

Private Sub DoLayOut(SkinPath As String, Form As Form)
'¸ù¾ÝÆ¤·ôÎÄ¼þÖØÐÂµ÷Õû¿Ø¼þ´óÐ¡ºÍÎ»ÖÃ  Resize and reposition all the controls according to the skin file
On Error Resume Next
Dim nFolderPath As String, nKeyVaule As String
    ' ÅÐ¶ÏÂ·¾¶ÊÇ·ñºÏ·¨ To see whether or not the path is valid
    If Len(Dir(Trim(SkinPath), vbArchive)) = 0 Then MsgBox "Invalid Skin Path! Make sure if the file exists.", , "": Exit Sub
    nFolderPath = GetFolderPath(SkinPath)
    nKeyVaule = GetINI("LayOut", "minbtnT", SkinPath)
        cmdMinBtn.Top = Val(nKeyVaule)
    nKeyVaule = GetINI("LayOut", "maxbtnT", SkinPath)
        cmdMaxBtn.Top = Val(nKeyVaule)
    nKeyVaule = GetINI("LayOut", "closebtnT", SkinPath)
        cmdCloseBtn.Top = Val(nKeyVaule)
    nKeyVaule = GetINI("LayOut", "minbtnL", SkinPath)
        cmdMinBtn.Left = Val(nKeyVaule)
    nKeyVaule = GetINI("LayOut", "maxbtnL", SkinPath)
        cmdMaxBtn.Left = Val(nKeyVaule)
    nKeyVaule = GetINI("LayOut", "closebtnL", SkinPath)
        cmdCloseBtn.Left = Val(nKeyVaule)
    nKeyVaule = GetINI("LayOut", "frmbtnsT", SkinPath)
        userFrmBtnT = Val(nKeyVaule)
        picFormBtn.Top = Val(nKeyVaule)
    nKeyVaule = GetINI("LayOut", "frmbtnsR", SkinPath)
        userFrmBtnR = Val(nKeyVaule)
        picFormBtn.Left = UserControl.ScaleWidth - Val(nKeyVaule)
    Dim x1, x2 As Integer
    x1 = Val(GetINI("LayOut", "WinMinW", SkinPath))
    x2 = Val(GetINI("LayOut", "WinMinH", SkinPath))
        Call SetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC, OldWindowProc)
        Call LimitSize(x1, x2)
    userRgn1 = GetINI("LayOut", "Rgn1", SkinPath)
    userRgn2 = GetINI("LayOut", "Rgn2", SkinPath)
        Call RgnForm(Form, userRgn1, userRgn2)
    intSizerBottom = GetINI("LayOut", "ResizerB", SkinPath)
    intSizerRight = GetINI("LayOut", "ResizerR", SkinPath)
End Sub

Private Sub RgnForm(ByVal frmbox As Form, ByVal fw As Long, ByVal fh As Long)
Dim w As Long, h As Long
    w = frmbox.ScaleX(frmbox.Width + 10, vbTwips, vbPixels)
    h = frmbox.ScaleY(frmbox.Height + 10, vbTwips, vbPixels)
    outRgn = CreateRoundRectRgn(0, 0, w, h, fw, fh)
    Call SetWindowRgn(frmbox.hwnd, outRgn, True)
End Sub

Public Function LoadSkin(SkinPath As String, Form As Form) ' ´ÓÂ·¾¶¼ÓÔØ×Ô¶¨ÒåÆ¤·ô  Load a skin from a specific skin file
On Error GoTo err
Dim nFolderPath As String, nKeyVaule As String
    ' ÅÐ¶ÏÂ·¾¶ÊÇ·ñºÏ·¨ To see whether or not the path is valid
    If Len(Dir(Trim(SkinPath), vbArchive)) = 0 Then MsgBox "Invalid Skin Path! Make sure if the file exists.", , "": Exit Function
    nFolderPath = GetFolderPath(SkinPath)
    If bMax Then Call DoMaximize
    With cmdMinBtn
        nKeyVaule = GetINI("Skin", "minbtn_1", SkinPath)
        .ImgUp = LoadPicture(nFolderPath & nKeyVaule)
        nKeyVaule = GetINI("Skin", "minbtn_2", SkinPath)
        .ImgOver = LoadPicture(nFolderPath & nKeyVaule)
        nKeyVaule = GetINI("Skin", "minbtn_3", SkinPath)
        .ImgDown = LoadPicture(nFolderPath & nKeyVaule)
        .DoFadeOut
    End With
    With cmdMaxBtn
        nKeyVaule = GetINI("Skin", "maxbtn_1", SkinPath)
        .ImgUp = LoadPicture(nFolderPath & nKeyVaule)
        nKeyVaule = GetINI("Skin", "maxbtn_2", SkinPath)
        .ImgOver = LoadPicture(nFolderPath & nKeyVaule)
        nKeyVaule = GetINI("Skin", "maxbtn_3", SkinPath)
        .ImgDown = LoadPicture(nFolderPath & nKeyVaule)
        .DoFadeOut
    End With
    With cmdCloseBtn
        nKeyVaule = GetINI("Skin", "Closebtn_1", SkinPath)
        .ImgUp = LoadPicture(nFolderPath & nKeyVaule)
        nKeyVaule = GetINI("Skin", "Closebtn_2", SkinPath)
        .ImgOver = LoadPicture(nFolderPath & nKeyVaule)
        nKeyVaule = GetINI("Skin", "Closebtn_3", SkinPath)
        .ImgDown = LoadPicture(nFolderPath & nKeyVaule)
        .DoFadeOut
    End With
    nKeyVaule = GetINI("Skin", "head_1", SkinPath)
        picHead_1.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "head_2", SkinPath)
        picHead_2.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "head_3", SkinPath)
        picHead_3.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "left_1", SkinPath)
        picLeft_1.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "left_2", SkinPath)
        picLeft_2.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "left_3", SkinPath)
        picLeft_3.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "bottom_1", SkinPath)
        picBottom_1.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "bottom_2", SkinPath)
        picBottom_2.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "bottom_3", SkinPath)
        picBottom_3.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "right_1", SkinPath)
        picRight_1.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "right_2", SkinPath)
        picRight_2.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "right_3", SkinPath)
        picRight_3.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "frmbtns", SkinPath)
        picFormBtn.Picture = LoadPicture(nFolderPath & nKeyVaule)
    nKeyVaule = GetINI("Skin", "CaptionColor", SkinPath)
        lblCaption.ForeColor = CLng(nKeyVaule)
    nKeyVaule = GetINI("Skin", "FontName", SkinPath)
        myfont.Name = nKeyVaule
    nKeyVaule = GetINI("Skin", "FontSize", SkinPath)
        myfont.Size = nKeyVaule
    nKeyVaule = GetINI("Skin", "FormColor", SkinPath)
        Form.BackColor = nKeyVaule
    nKeyVaule = GetINI("Skin", "FontBold", SkinPath)
    If UCase(nKeyVaule) = "TRUE" Then
        myfont.Bold = True
        Else
        myfont.Bold = False
    End If
    Set lblCaption.Font = myfont
    Call DoLayOut(SkinPath, Form)
    Call ShowWindow(UserControl.ContainerHwnd, SW_MINIMIZE)
    Call UserControl_Resize
    Call ShowWindow(UserControl.ContainerHwnd, SW_NORMAL)
    RaiseEvent SkinChanged
    Debug.Print "Skin Changed: " & SkinPath
    Exit Function
err:
    MsgBox "Failed to load skin -> " & vbCrLf & SkinPath: Resume Next
End Function

Public Function LimitSize(ByVal nWidth As Long, ByVal nHeight As Long) ' ÏÞÖÆ´°¿Ú´óÐ¡ Limit the size of the form
On Error Resume Next
    MinSizeX = nWidth
    MinSizeY = nHeight
    OldWindowProc = GetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC)
    Call SetWindowLong(UserControl.ContainerHwnd, GWL_WNDPROC, AddressOf WndMessage)
End Function

Public Function DrawVistorForm(TheForm As Form)
On Error Resume Next
'¼ÓÔØVistor´°Ìå   Draw Vistor Form
    With TheForm
        .AutoRedraw = True
        .ScaleMode = 3
        .Cls   'Çå³ýÉÏ´Î»æÍ¼ºÛ¼£    Clean The Form
        .PaintPicture picLeft_2.Picture, 0, picHead_1.Height, , TheForm.ScaleHeight
        .PaintPicture picRight_2.Picture, .ScaleWidth - picRight_1.ScaleWidth, picHead_1.Height, , TheForm.ScaleHeight
        .PaintPicture picBottom_2.Picture, 0, .ScaleHeight - picBottom_1.ScaleHeight, TheForm.ScaleWidth
        .PaintPicture picLeft_1.Picture, 0, picHead_1.ScaleHeight
        .PaintPicture picLeft_3.Picture, 0, .ScaleHeight - picLeft_3.ScaleHeight
        .PaintPicture picRight_1.Picture, .ScaleWidth - picRight_1.ScaleWidth, picHead_1.Height
        .PaintPicture picRight_3.Picture, .ScaleWidth - picRight_1.ScaleWidth, .ScaleHeight - picRight_3.ScaleHeight
        .PaintPicture picBottom_3.Picture, .ScaleWidth - picBottom_3.ScaleWidth, .ScaleHeight - picBottom_3.ScaleHeight
        .PaintPicture picBottom_1.Picture, 0, .ScaleHeight - picBottom_1.ScaleHeight
        '.Refresh
    End With
    '»æÖÆÔ²½Ç´°Ìå  Draw Rounded-rectangular Form
    Call RgnForm(TheForm, userRgn1, userRgn2)
End Function

Private Sub cmdCloseBtn_Click()
'¹Ø±Õ´°Ìå   Close the Form
   RaiseEvent FormBtnClick(3)
   Call SendMessage(UserControl.ContainerHwnd, WM_CLOSE, 1, 1)
End Sub

Private Sub cmdMaxBtn_Click()
'×î´ó»¯´°Ìå   Maximize the Form
   RaiseEvent FormBtnClick(2)
   Call DoMaximize
End Sub

Private Sub cmdMinBtn_Click()
'×îÐ¡»¯´°Ìå    Minimize the Form
   RaiseEvent FormBtnClick(1)
   Call ShowWindow(UserControl.ContainerHwnd, SW_MINIMIZE)
End Sub

Private Sub lblCaption_Click()
    RaiseEvent Click
End Sub

Private Sub lblCaption_DblClick()
   Call DoMaximize
   RaiseEvent DblClick
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
   Call DoDrag
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
   Call DoMaximize
   RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    UserControl.AutoRedraw = True
    picHead_2.AutoRedraw = True
    bDraggable = True
    Set myfont = lblCaption.Font
    MinSizeX = 550
    MinSizeY = 200
    userFrmBtnT = 1
    userFrmBtnR = 101
    userRgn1 = 10
    userRgn2 = 8
    intSizerBottom = 6
    intSizerRight = 6
    Call UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
   Call DoDrag
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Vistor")
   Call PropBag.WriteProperty("Icon", imgIcon.Picture, LoadPicture())
   Call PropBag.WriteProperty("AlphaBlend", cmdMinBtn.UseAlphaBlend, True)
   Call PropBag.WriteProperty("EnableMax", bEnMax, True)
   Call PropBag.WriteProperty("CloseOnly", bCloseOnly, False)
   Call PropBag.WriteProperty("CaptionColor", lblCaption.ForeColor, &H0&)
   Call PropBag.WriteProperty("Font", myfont, lblCaption.Font)
   Call PropBag.WriteProperty("Shadow", bShadow, True)
   Call PropBag.WriteProperty("Transparency", intTransparency, 0)
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   lblCaption.Caption = PropBag.ReadProperty("Caption", "Vistor")
   Set imgIcon.Picture = PropBag.ReadProperty("Icon", LoadPicture())
   cmdMinBtn.UseAlphaBlend = PropBag.ReadProperty("AlphaBlend", True)
   cmdMaxBtn.UseAlphaBlend = PropBag.ReadProperty("AlphaBlend", True)
   cmdCloseBtn.UseAlphaBlend = PropBag.ReadProperty("AlphaBlend", True)
   bCloseOnly = PropBag.ReadProperty("CloseOnly", False)
   lblCaption.ForeColor = PropBag.ReadProperty("CaptionColor", &H0&)
   Set myfont = PropBag.ReadProperty("Font", lblCaption.Font)
   Set lblCaption.Font = myfont
   bShadow = PropBag.ReadProperty("Shadow", True)
   intTransparency = PropBag.ReadProperty("Transparency", 0)
    rtn = GetWindowLong(UserControl.ContainerHwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong UserControl.ContainerHwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes UserControl.ContainerHwnd, 0, (255 - intTransparency), LWA_ALPHA
    If bShadow Then
        SetClassLong UserControl.ContainerHwnd, GCL_STYLE, GetClassLong(UserControl.ContainerHwnd, GCL_STYLE) Or CS_DROPSHADOW
    End If
    If bCloseOnly Then
      cmdMinBtn.Visible = False
      cmdMaxBtn.Visible = False
    Else
      cmdMinBtn.Visible = True
      bEnMax = PropBag.ReadProperty("EnableMax", True)
      cmdMaxBtn.Visible = bEnMax
    End If
End Sub

Public Property Get Icon() As StdPicture
    Set Icon = imgIcon.Picture
End Property

Public Property Set Icon(img As StdPicture)
    Set imgIcon.Picture = img
    PropertyChanged "Icon"
End Property

Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(strCaption As String)
    lblCaption.Caption = strCaption
    PropertyChanged "Caption"
End Property

Public Property Get AlphaBlend() As Boolean
    AlphaBlend = cmdMinBtn.UseAlphaBlend
End Property

Public Property Let AlphaBlend(bBlend As Boolean)
On Error Resume Next
    cmdMinBtn.UseAlphaBlend = bBlend
    cmdMaxBtn.UseAlphaBlend = bBlend
    cmdCloseBtn.UseAlphaBlend = bBlend
    PropertyChanged "AlphaBlend"
End Property

Public Property Get SizerBottom() As Integer
    SizerBottom = intSizerBottom
End Property
Public Property Get SizerRight() As Integer
    SizerRight = intSizerRight
End Property

Public Property Get CloseOnly() As Boolean
    CloseOnly = bCloseOnly
End Property

Public Property Let CloseOnly(bClose As Boolean)
On Error Resume Next
    bCloseOnly = bClose
    If bClose Then
       cmdMinBtn.Visible = False
       cmdMaxBtn.Visible = False
    Else
       cmdMinBtn.Visible = True
       If bEnMax Then cmdMinBtn.Visible = True Else cmdMinBtn.Visible = False: PropertyChanged "EnableMax"
    End If
    PropertyChanged "CloseOnly"
End Property

Public Property Get EnableMaximize() As Boolean
    EnableMaximize = bEnMax
End Property

Public Property Let EnableMaximize(bEnableMax As Boolean)
On Error Resume Next
    If bCloseOnly Then Exit Property
    bEnMax = bEnableMax
    cmdMaxBtn.Visible = bEnableMax
    PropertyChanged "EnableMax"
End Property

Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = lblCaption.ForeColor
End Property

Public Property Let CaptionColor(Color As OLE_COLOR)
On Error Resume Next
    lblCaption.ForeColor = Color
    PropertyChanged "CaptionColor"
End Property

Public Property Get Shadow() As Boolean
    Shadow = bShadow
End Property

Public Property Let Shadow(bValue As Boolean)
On Error Resume Next
    bShadow = bValue
    PropertyChanged "Shadow"
End Property

Public Property Get Transparency() As Integer
    Transparency = intTransparency
End Property

Public Property Let Transparency(value As Integer)
On Error Resume Next
    If value > 250 Or value < 0 Then Exit Property
    intTransparency = value
    rtn = GetWindowLong(UserControl.ContainerHwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong UserControl.ContainerHwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes UserControl.ContainerHwnd, 0, (255 - value), LWA_ALPHA
    PropertyChanged "Transparency"
End Property

Public Property Get Font() As StdFont
  Dim tf  As New StdFont
  tf.Bold = myfont.Bold
  tf.Italic = myfont.Italic
  tf.Name = myfont.Name
  tf.Size = myfont.Size
  tf.Strikethrough = myfont.Strikethrough
  tf.Underline = myfont.Underline
  Set Font = tf
End Property

Public Property Set Font(ByVal newFont As StdFont)
On Error Resume Next
Dim tf  As New StdFont
  Set tf = newFont
  myfont.Bold = tf.Bold
  myfont.Italic = tf.Italic
  myfont.Name = tf.Name
  myfont.Size = tf.Size
  myfont.Strikethrough = tf.Strikethrough
  myfont.Underline = tf.Underline
  Set lblCaption.Font = tf
  PropertyChanged "Font"
End Property

Public Property Get Draggable() As Boolean
    Draggable = bDraggable
End Property

Private Sub UserControl_Resize()
On Error Resume Next

   DeleteObject outRgn
   'ÖØ¶¨Òå¿Ø¼þ´óÐ¡¼°Î»ÖÃ   Resize and reposition all the controls
   With UserControl
      .Cls
      picHead_2.Left = .picHead_1.Width
      picHead_3.Left = .ScaleWidth - picHead_3.Width
      picHead_2.Width = .ScaleWidth - picHead_1.Width - picHead_3.Width
      .Height = .picHead_1.Height * Screen.TwipsPerPixelY
      .AutoRedraw = True
      .PaintPicture picHead_2.Picture, 0, 0, .ScaleWidth
      .PaintPicture picHead_1.Picture, 0, 0
      .PaintPicture picHead_3.Picture, .ScaleWidth - picHead_3.ScaleWidth, 0
      .picFormBtn.Top = userFrmBtnT
      .picFormBtn.Left = .ScaleWidth - userFrmBtnR
      .lblCaption.Top = (.ScaleHeight - .lblCaption.Height) / 2
      .imgIcon.Top = (.ScaleHeight - .imgIcon.Height) / 2
   End With
   
End Sub

