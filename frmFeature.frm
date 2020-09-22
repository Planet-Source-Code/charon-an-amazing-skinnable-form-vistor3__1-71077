VERSION 5.00
Begin VB.Form frmMore 
   Caption         =   "Test"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000016&
      Height          =   3495
      Left            =   1680
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   0
      Top             =   1320
      Width           =   8055
      Begin Vistor3.ctlVistorForm ctlVistorForm 
         CausesValidation=   0   'False
         Height          =   450
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8010
         _extentx        =   14129
         _extenty        =   794
         caption         =   "I'm in a PictureBox!  Try to drag me!"
         icon            =   "frmFeature.frx":0000
         captioncolor    =   16777215
         transparency    =   150
         enablemax       =   0   'False
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "This is a PictureBox."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Absolutely No Code!
