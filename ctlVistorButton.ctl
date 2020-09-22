VERSION 5.00
Begin VB.UserControl ctlVistorButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox PicOver 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   1680
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   2
      Top             =   2160
      Width           =   405
   End
   Begin VB.PictureBox PicUp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   1680
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   1
      Top             =   1800
      Width           =   405
   End
   Begin VB.PictureBox PicDown 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00AEAEAE&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   1680
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   0
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "ctlVistorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************
'*
'*                         Thomas John (2003)
'*                        thomas.john@swing.be
'*
'**********************************************************************

'variables
Private bCapture As Boolean
Private lngRep As Long
Private EtatBut As Long
Private TransOK As Boolean
Private TransparanceSz As Long
Private DessusSz As Boolean
Private bPanel As Boolean
'
'--- API AlphaBlend ------------------
'
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
'
Private Const AC_SRC_OVER = &H0
'
Private Const AC_SRC_ALPHA = &H1
'
'---------------------------------------------------
Private WithEvents MinSz As Minuteur
Attribute MinSz.VB_VarHelpID = -1
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseOver(x As Single, y As Single)
Event MouseOut()

Public Function DoFadeOut()
    BitBlt UserControl.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hDc, 0, 0, vbSrcCopy
    UserControl.Refresh
End Function

Public Function DoSetCapture()
     bCapture = True
     lngRep = SetCapture(UserControl.hwnd)
End Function

Public Function DoReleaseCapture()
     bCapture = False
     ReleaseCapture
End Function

Private Sub lblCaption_Click()
    ReleaseCapture
    RaiseEvent Click
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    MinSz.Actif = False
    EtatBut = 1
    BitBlt UserControl.hDc, 0, 0, PicDown.ScaleWidth, PicDown.ScaleHeight, PicDown.hDc, 0, 0, vbSrcCopy
    UserControl.Refresh
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    DoEvents
    If bCapture = False Then
        bCapture = True
        lngRep = SetCapture(UserControl.hwnd)
        DessusSz = True
        If TransOK = True Then
            MinSz.Actif = True
        Else
            BitBlt UserControl.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, PicOver.hDc, 0, 0, vbSrcCopy
            UserControl.Refresh
        End If
    End If
    If x < 0 Or y < 0 Or x > UserControl.Width Or y > UserControl.Height Then
        DessusSz = False
        If EtatBut = 1 Then
            BitBlt UserControl.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, PicOver.hDc, 0, 0, vbSrcCopy
            UserControl.Refresh
        Else
            bCapture = False
            lngRep = ReleaseCapture
            If TransOK = True Then
                MinSz.Actif = True
            Else
                BitBlt UserControl.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hDc, 0, 0, vbSrcCopy
                UserControl.Refresh
            End If
        End If
        RaiseEvent MouseOut
    Else
        DessusSz = True
        If EtatBut = 1 Then
            BitBlt UserControl.hDc, 0, 0, PicDown.ScaleWidth, PicDown.ScaleHeight, PicDown.hDc, 0, 0, vbSrcCopy
            UserControl.Refresh
        Else
            If TransOK = True Then
                MinSz.Actif = True
            Else
                BitBlt UserControl.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, PicOver.hDc, 0, 0, vbSrcCopy
                UserControl.Refresh
            End If

        End If
        RaiseEvent MouseOver(x, y)
    End If

End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    EtatBut = 0
    bCapture = False
    lngRep = ReleaseCapture
    '
    If DessusSz = False Then
        '
        MinSz.Actif = True
        '
    Else
        '
        MinSz.Actif = False
        TransparanceSz = 0
        '
        BitBlt UserControl.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hDc, 0, 0, vbSrcCopy
        '
        UserControl.Refresh
        '
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)
    '
End Sub


Private Sub MinSz_Action()
On Error Resume Next

    Dim lBlend As Long
    Dim bf As BLENDFUNCTION
    '
    If DessusSz = True Then
        '
        TransparanceSz = TransparanceSz + 10
        If TransparanceSz >= 128 Then
            '
            MinSz.Actif = False
            TransparanceSz = 128
            '
        End If
        '
        bf.BlendOp = 0
        bf.BlendFlags = 0
        bf.SourceConstantAlpha = TransparanceSz
        bf.AlphaFormat = 0
        CopyMemory lBlend, bf, 4
        '
        AlphaBlend UserControl.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, PicOver.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, lBlend
        '
    Else
        '
        TransparanceSz = TransparanceSz - 10
        '
        If TransparanceSz <= 0 Then
            '
            MinSz.Actif = False
            TransparanceSz = 0
            '
        End If
        '
        bf.BlendOp = 0
        bf.BlendFlags = 0
        bf.SourceConstantAlpha = 128 - TransparanceSz
        bf.AlphaFormat = 0
        CopyMemory lBlend, bf, 4
        '
        AlphaBlend UserControl.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, lBlend
        '
    End If
    '
    UserControl.Refresh
    '
End Sub

'
Private Sub UserControl_Initialize()
    '
    PicUp.Visible = False
    PicDown.Visible = False
    PicOver.Visible = False
    '
    EtatBut = 0
    TransparanceSz = 0
    DessusSz = False
    '
    Set MinSz = New Minuteur
    MinSz.Intervalle = 40
    DoFadeOut
    UserControl_Resize
    '
End Sub
'
Private Sub UserControl_InitProperties()
    '
    TransOK = False
    '
End Sub
'
Private Sub UserControl_DblClick()
    
    UserControl_MouseDown 1, 0, 0, 0
    
End Sub
'
Private Sub UserControl_Click()
    '
    ReleaseCapture
    'RaiseEvent Click
    '
End Sub
'
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    If bPanel = False Then
    MinSz.Actif = False
    EtatBut = 1
    BitBlt UserControl.hDc, 0, 0, PicDown.ScaleWidth, PicDown.ScaleHeight, PicDown.hDc, 0, 0, vbSrcCopy
    UserControl.Refresh
    End If
    RaiseEvent MouseDown(Button, Shift, x, y)
    '
End Sub
'
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    DoEvents
    If bCapture = False Then
        '
        bCapture = True
        lngRep = SetCapture(UserControl.hwnd)
        '
        DessusSz = True
        If TransOK = True Then
            '
            MinSz.Actif = True
            '
        Else
            'UserControl.Picture = PicOver.Picture
            BitBlt UserControl.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, PicOver.hDc, 0, 0, vbSrcCopy
            '
            UserControl.Refresh
            '
        End If
        '
    End If
    '
    If x < 0 Or y < 0 Or x > UserControl.ScaleWidth Or y > UserControl.ScaleHeight Then
        '
        DessusSz = False
        If EtatBut = 1 Then
            '
            'UserControl.Picture = PicOver.Picture
            BitBlt UserControl.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, PicOver.hDc, 0, 0, vbSrcCopy
            '
            UserControl.Refresh
            '
        Else
            '
            bCapture = False
            lngRep = ReleaseCapture
            If TransOK = True Then
                '
                MinSz.Actif = True
                '
            Else
                '
                'UserControl.Picture = PicUp.Picture
                BitBlt UserControl.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hDc, 0, 0, vbSrcCopy
                '
                UserControl.Refresh
                '
            End If
            '
        End If
        '
        RaiseEvent MouseOut
        '
    Else
        '
        '
        DessusSz = True
        '
        If EtatBut = 1 Then
            '
            'UserControl.Picture = PicDown.Picture
            BitBlt UserControl.hDc, 0, 0, PicDown.ScaleWidth, PicDown.ScaleHeight, PicDown.hDc, 0, 0, vbSrcCopy
            '
            UserControl.Refresh
            '
        Else
            '
            If TransOK = True Then
                '
                MinSz.Actif = True
                '
            Else
                '
                BitBlt UserControl.hDc, 0, 0, PicOver.ScaleWidth, PicOver.ScaleHeight, PicOver.hDc, 0, 0, vbSrcCopy
                UserControl.Refresh
                '
            End If
            '
        End If
        '
        RaiseEvent MouseOver(x, y)
        '
    End If
    '
End Sub
'
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '
    EtatBut = 0
    bCapture = False
    lngRep = ReleaseCapture
    '
    If DessusSz = False Then
        '
        MinSz.Actif = True
        '
    Else
        If bPanel = False Then '
        MinSz.Actif = False
        TransparanceSz = 0
        'UserControl.Picture = PicUp.Picture
        BitBlt UserControl.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hDc, 0, 0, vbSrcCopy
        UserControl.Refresh
        End If
    End If
    '
    RaiseEvent Click
    RaiseEvent MouseUp(Button, Shift, x, y)
    '
End Sub
'
Private Sub UserControl_Resize()
    '
    UserControl.Width = PicUp.Width
    UserControl.Height = PicUp.Height
    lblCaption.Left = UserControl.Width / 2 - lblCaption.Width / 2
    lblCaption.Top = UserControl.Height / 2 - lblCaption.Height / 2 + 10
    '
End Sub
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    With PropBag
        '
        PicUp.Picture = .ReadProperty("ImgUp")
        PicDown.Picture = .ReadProperty("ImgDown")
        PicOver.Picture = .ReadProperty("ImgOver")
        lblCaption.Caption = .ReadProperty("Caption", "ChariButton")
        TransOK = .ReadProperty("UseAlphaBlend", False)
        bPanel = .ReadProperty("Panel", False)
        '
    End With
    '
    'UserControl.Picture = PicUp.Picture
    BitBlt UserControl.hDc, 0, 0, PicUp.ScaleWidth, PicUp.ScaleHeight, PicUp.hDc, 0, 0, vbSrcCopy
    '
    UserControl_Resize
    '
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    With PropBag
        '
        .WriteProperty "ImgUp", PicUp.Picture
        .WriteProperty "ImgDown", PicDown.Picture
        .WriteProperty "ImgOver", PicOver.Picture
        .WriteProperty "Caption", lblCaption.Caption, "ChariButton"
        .WriteProperty "UseAlphaBlend", TransOK, False
        .WriteProperty "Panel", bPanel, False
        '
    End With
    '
End Sub
'
'*******************************************************************************************************
'* PROPRIETES
'*******************************************************************************************************
'
'UTILISATION DE LA TRANSPARANCE OU PAS
Public Property Let UseAlphaBlend(Valeur As Boolean)
    '
    TransOK = Valeur
    '
    PropertyChanged "UseAlphaBlend"
    '
End Property
'
Public Property Get UseAlphaBlend() As Boolean
    '
    UseAlphaBlend = TransOK
    '
End Property

Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(strCaption As String)
    lblCaption.Caption = strCaption
    PropertyChanged "Caption"
    lblCaption.Left = UserControl.Width / 2 - lblCaption.Width / 2
End Property

Public Property Get Panel() As Boolean
    Panel = bPanel
End Property

Public Property Let Panel(value As Boolean)
    bPanel = value
    PropertyChanged "Panel"
End Property

'IMAGE BOUTTON NORMAL
Public Property Let ImgUp(Valeur As StdPicture)
    '
    PicUp.Picture = Valeur
    PropertyChanged "ImgUp"
    UserControl_Resize
    '
End Property
'
Public Property Set ImgUp(Valeur As StdPicture)
    '
    PicUp.Picture = Valeur
    PropertyChanged "ImgUp"
    UserControl_Resize
    '
End Property
'
Public Property Get ImgUp() As StdPicture
    '
    Set ImgUp = PicUp.Picture
    '
End Property
'
'IMAGE BOUTTON PRESSE
Public Property Let ImgDown(Valeur As StdPicture)
    '
    PicDown.Picture = Valeur
    PropertyChanged "ImgDown"
    UserControl_Resize
    '
End Property
'
Public Property Set ImgDown(Valeur As StdPicture)
    '
    PicDown.Picture = Valeur
    PropertyChanged "ImgDown"
    UserControl_Resize
    '
End Property
'
Public Property Get ImgDown() As StdPicture
    '
    Set ImgDown = PicDown.Picture
    '
End Property
'
Public Property Let ImgOver(Valeur As StdPicture)
    '
    PicOver.Picture = Valeur
    PropertyChanged "ImgOver"
    UserControl_Resize
    '
End Property
'
Public Property Set ImgOver(Valeur As StdPicture)
    '
    PicOver.Picture = Valeur
    PropertyChanged "ImgOver"
    UserControl_Resize
    '
End Property
'
Public Property Get ImgOver() As StdPicture
    '
    Set ImgOver = PicOver.Picture
    '
End Property
