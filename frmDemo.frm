VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDemo 
   BorderStyle     =   0  'None
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin Vistor3.ctlVistorForm ctlVistorForm 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   794
      Icon            =   "frmDemo.frx":0A02
   End
   Begin VB.PictureBox picSide 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   480
      ScaleHeight     =   221
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   12
      Top             =   1200
      Width           =   1515
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   1800
         Width           =   1500
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "Clear &List"
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   1440
         Width           =   1500
      End
      Begin VB.CommandButton cmdMore 
         Caption         =   "&More..."
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   1500
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About Vistor 3"
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   1500
      End
      Begin VB.CommandButton cmdImprove 
         Caption         =   "&Improvements"
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   1500
      End
      Begin VB.CommandButton cmdSkinInfo 
         Caption         =   "&Skin Information"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1500
      End
   End
   Begin VB.ListBox lstMsg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      ItemData        =   "frmDemo.frx":1414
      Left            =   2040
      List            =   "frmDemo.frx":1416
      TabIndex        =   11
      Top             =   1200
      Width           =   6615
   End
   Begin VB.ComboBox cboSkinFile 
      Height          =   330
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   675
      Width           =   2655
   End
   Begin ComctlLib.Slider sldTransparency 
      Height          =   255
      Left            =   5070
      TabIndex        =   0
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   327682
      SmallChange     =   10
      Max             =   250
      SelStart        =   15
      TickFrequency   =   10
      Value           =   15
   End
   Begin VB.Timer tmrDraggable 
      Interval        =   100
      Left            =   0
      Top             =   960
   End
   Begin VB.Label lblOpacity 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transparency"
      Height          =   210
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblSkin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skin"
      Height          =   210
      Left            =   480
      TabIndex        =   10
      Top             =   720
      Width           =   330
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      Height          =   210
      Left            =   8400
      TabIndex        =   3
      Top             =   720
      Width           =   210
   End
   Begin VB.Image imgResizer 
      Height          =   255
      Left            =   0
      MousePointer    =   8  'Size NW SE
      Stretch         =   -1  'True
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgRightResizer 
      Height          =   255
      Left            =   240
      MousePointer    =   9  'Size W E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   210
   End
   Begin VB.Image imgButtomResizer 
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      Stretch         =   -1  'True
      Top             =   720
      Width           =   465
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  '==========================================================================
  '                                                             This code is written by Charon (2008).
  '                                        If you have any problems using this Control, please contact me.
  '                                                         My E-mial Address: astrophsyics@126.com
  '==========================================================================
  
Option Explicit

Private Sub cboSkinFile_Click()
    Call ctlVistorForm.LoadSkin(App.Path & "\Skins\" & cboSkinFile.Text & "\VistorSkin.ini", Me)
End Sub

Private Sub cmdAbout_Click()
    lstMsg.Clear
    lstMsg.AddItem ""
    lstMsg.AddItem " This is the third version of Vistor."
    lstMsg.AddItem " Vistor is originally written to imitate Windows Vista User Interface."
    lstMsg.AddItem " Because somebody asked me to extend the functions, so i decided to develop Vistor 3."
    lstMsg.AddItem " In this version, you can find something very different from the previous versions,"
    lstMsg.AddItem " For instance, you can maximize the form, as well as set optional transparency."
    lstMsg.AddItem " You can even change the skin into your own. To customize your own skin, you only"
    lstMsg.AddItem " need to prepare some skin ScreenShots and cut them into several parts, then edit the"
    lstMsg.AddItem " 'VistorSkin.ini' file. If you have some problems using this code, you may contact me at"
    lstMsg.AddItem " astrophysics@126.com. Any comments are welcome."
    lstMsg.AddItem " Last but not least, don't forget to vote for me~!"
    lstMsg.AddItem ""
End Sub

Private Sub cmdClean_Click()
     lstMsg.Clear
End Sub

Private Sub cmdClose_Click()
       Call SendMessage(Me.hwnd, WM_CLOSE, 1, 1)
       'Call SendMessage(frmMore.hwnd, WM_CLOSE, 1, 1)
End Sub

Private Sub cmdImprove_Click()
    lstMsg.Clear
    lstMsg.AddItem ""
    lstMsg.AddItem " What's new? What has been improved?"
    lstMsg.AddItem " 1 - Skinnable Form (Allow users to customize their own skins.)"
    lstMsg.AddItem " 2 - Sizable Form"
    lstMsg.AddItem " 3 - Maximizable Form"
    lstMsg.AddItem "      (Here, notice that when a form is maximized, it won't cover the taskbar.)"
    lstMsg.AddItem " 4 - Optional Transparency (0 ~ 250)"
    lstMsg.AddItem " 5 - Work as a Usercontrol, more easy to use"
    lstMsg.AddItem " 6 - Fix some other bugs"
    lstMsg.AddItem ""
End Sub

Private Sub cmdMore_Click()
    frmMore.Show , Me
End Sub

Private Sub cmdSkinInfo_Click()
On Error Resume Next
Dim strPath As String
    strPath = App.Path & "\Skins\" & cboSkinFile.Text & "\VistorSkin.ini"
    lstMsg.AddItem ""
    lstMsg.AddItem " Following text is the Information of Skin " & Me.cboSkinFile.Text
    lstMsg.AddItem " Name=" & GetINI("SkinInfo", "Name", strPath)
    lstMsg.AddItem " Contact=" & GetINI("SkinInfo", "Contact", strPath)
    lstMsg.AddItem " Version=" & GetINI("SkinInfo", "Version", strPath)
    lstMsg.AddItem " Author=" & GetINI("SkinInfo", "Author", strPath)
    lstMsg.AddItem " Description=" & GetINI("SkinInfo", "Description", strPath)
    lstMsg.AddItem " Copyright=" & GetINI("SkinInfo", "Copyright", strPath)
    lstMsg.AddItem ""
End Sub

Private Sub ctlVistorForm_DblClick()
    lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Double Click."
End Sub

Private Sub ctlVistorForm_FormBtnClick(FormButton As Integer)
Select Case FormButton
    Case 1: lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Minimize Button Clicked."
    Case 2: lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Maximize Button Clicked."
    Case 3: lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Close Button Clicked."
End Select
End Sub

Private Sub ctlVistorForm_SkinChanged()
    lblOpacity.Font = ctlVistorForm.Font
    lblValue.Font = ctlVistorForm.Font
    lblSkin.Font = ctlVistorForm.Font
    cboSkinFile.Font = ctlVistorForm.Font
    cmdSkinInfo.Font = ctlVistorForm.Font
    cmdAbout.Font = ctlVistorForm.Font
    cmdClose.Font = ctlVistorForm.Font
    cmdClean.Font = ctlVistorForm.Font
    cmdMore.Font = ctlVistorForm.Font
    cmdImprove.Font = ctlVistorForm.Font
    lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Skin Changed. [" & cboSkinFile.Text & "]"
End Sub

Private Sub Form_Load()
    Me.Caption = ctlVistorForm.Caption
    Call ctlVistorForm.LimitSize(610, 340)
    lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Starting..."
    imgButtomResizer.Left = 0
    imgButtomResizer.Height = ctlVistorForm.SizerBottom
    imgRightResizer.Width = ctlVistorForm.SizerRight
    imgRightResizer.Top = ctlVistorForm.Height + ctlVistorForm.Top
    imgResizer.Width = ctlVistorForm.SizerBottom
    imgResizer.Height = ctlVistorForm.SizerRight
    GetDirOp App.Path & "\Skins", cboSkinFile
    cboSkinFile.Text = "VistaLikeEx(Blue)"
    Call sldTransparency_Scroll
    lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Done. Welcome to Vistor v3.05."
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call ctlVistorForm.DrawVistorForm(Me)
    imgButtomResizer.Top = Me.ScaleHeight - imgButtomResizer.Height
    imgRightResizer.Left = Me.ScaleWidth - imgRightResizer.Width
    imgButtomResizer.Width = Me.ScaleWidth - imgButtomResizer.Left - imgRightResizer.Width
    imgRightResizer.Height = Me.ScaleHeight - imgRightResizer.Top - imgButtomResizer.Height
    With imgResizer
       .Top = Me.ScaleHeight - .Height
       .Left = Me.ScaleWidth - .Width
    End With
    lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Resized. [W=" & Me.ScaleWidth & " H=" & Me.ScaleHeight & "]"
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
     Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OldWindowProc)
End Sub

Private Sub imgButtomResizer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Not ctlVistorForm.Draggable Then imgButtomResizer.MousePointer = 0: Exit Sub
     imgButtomResizer.MousePointer = 7
     If Button <> 0 Then
         Me.Height = Me.Height + y
     End If
End Sub

Private Sub imgResizer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Not ctlVistorForm.Draggable Then imgResizer.MousePointer = 0: Exit Sub
     imgResizer.MousePointer = 8
     If Button <> 0 Then
         Me.Height = Me.Height + y
         Me.Width = Me.Width + x
     End If
End Sub

Private Sub imgRightResizer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Not ctlVistorForm.Draggable Then imgRightResizer.MousePointer = 0: Exit Sub
     imgRightResizer.MousePointer = 9
     If Button <> 0 Then
          Me.Width = Me.Width + x
     End If
End Sub

Private Sub sldTransparency_Change()
    Call sldTransparency_Scroll
    lstMsg.AddItem Format(Now, "HH:MM:SS") & "  Transrarency Changed. [" & sldTransparency.value & "; " & Int((sldTransparency.value / 255) * 100) & "%]"
End Sub

Private Sub sldTransparency_Scroll()
Dim rtn As Long

    lblValue.Caption = sldTransparency.value
    ctlVistorForm.Transparency = sldTransparency.value
    
End Sub

Private Sub tmrDraggable_Timer()
If ctlVistorForm.Draggable = True Then
    ctlVistorForm.Caption = "Vistor Skinnable Form v3.05 Demo - Draggable"
    Me.Caption = "Vistor Skinnable Form v3.05 Demo - Draggable"
Else
    ctlVistorForm.Caption = "Vistor Skinnable Form v3.05 Demo - Undraggable"
    Me.Caption = "Vistor Skinnable Form v3.05 Demo - Undraggable"
End If
End Sub
