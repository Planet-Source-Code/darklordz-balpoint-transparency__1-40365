VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  TransParency (Win2K/XP Only!)"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMkT 
      Height          =   3090
      Left            =   45
      TabIndex        =   0
      Top             =   -45
      Width           =   4590
      Begin VB.Timer tmrTransparent 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4050
         Top             =   2565
      End
      Begin MSComctlLib.Slider Slider 
         Height          =   255
         Left            =   165
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   825
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Max             =   100
         SelStart        =   100
         TickStyle       =   3
         Value           =   100
      End
      Begin VB.Label lblInfo 
         Caption         =   "Info #"
         Height          =   390
         Left            =   180
         TabIndex        =   5
         Top             =   2490
         Width           =   4140
      End
      Begin VB.Label lblPerc 
         Caption         =   "Percentage: #"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   570
         Width           =   4170
      End
      Begin VB.Label lblRGB 
         Caption         =   "RGB Value: #"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   420
         Width           =   4170
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status: #"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   4170
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###############################################
'# BALPOINT TRANSPARENCY! V1.0                 #
'# =========================================== #
'# USES THE USER32.DLL FROM WINDOWS XP/2K      #
'# NOT COMPATIBLE WITH OTHER WINDOWS!          #
'# =========================================== #
'# THIS CODE IS FREEWARE AND MAY BE USED       #
'# FREELY. BUT U MUST LEAVE THE COMMENTS!      #
'# =========================================== #
'# ALL RIGHTS RESERVED, © BalPoint.com - 2002  #
'###############################################
Private Sub Form_Load()
'   SET INFORMATION LABEL CAPTION
    lblInfo.Caption = "© Balpoint.com - 2002. All Rights Reserved." & vbCrLf & _
                      "Visit Us online @ http://www.balpoint.com"
End Sub
Private Sub Slider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'   ENABLE TIMER ON MOUSE DOWN OF THE SLIDER
    tmrTransparent.Enabled = True
End Sub
Private Sub Slider_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   DISABLE TIMER ON MOUSE DOWN OF THE SLIDER
    tmrTransparent.Enabled = False
End Sub
Private Sub tmrTransparent_Timer()
'   MAKE TRANSPARENT
    MakeTransparent Me.hWnd, (Slider.Value * 2.55)
'   SET CAPTIONS
    lblStatus.Caption = "Status: Form Transparent (" & isTransparent(Me.hWnd) & ")"
    lblRGB.Caption = "RGB Value: " & (Slider.Value * 2.55) & ")"
    lblPerc.Caption = "Percentage: " & Slider.Value & "%"
'   IF RGB VALUE = 255 RESET STATUS!
    If Slider.Value = 100 Then
        MakeOpaque Me.hWnd
        lblStatus.Caption = "Status: Form Transparent (" & isTransparent(Me.hWnd) & ")"
    End If
End Sub
