VERSION 5.00
Begin VB.Form frmABOUT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Telnet"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "License:"
      Height          =   615
      Left            =   50
      TabIndex        =   7
      Top             =   2770
      Width           =   4575
      Begin VB.Label lbl_regged_to 
         Alignment       =   1  'Right Justify
         Caption         =   "UNREGISTERED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "This software is licensed to:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Version:"
      Height          =   600
      Left            =   50
      TabIndex        =   4
      Top             =   3375
      Width           =   3255
      Begin VB.Label lblver 
         Caption         =   "<buid>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Simple Telnet for Windows:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame frameabout 
      Caption         =   "About Information:"
      Height          =   2775
      Left            =   50
      TabIndex        =   1
      Top             =   0
      Width           =   4575
      Begin VB.Label Label5 
         Caption         =   "Telnet"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   490
      Left            =   3410
      TabIndex        =   0
      Top             =   3465
      Width           =   1215
   End
End
Attribute VB_Name = "frmABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Hide
End Sub
Private Sub Form_Load()
lblver.Caption = App.Major & "." & App.Revision
If (frmMAIN.is_regged = True) Then
lbl_regged_to.ForeColor = vbBlack
lbl_regged_to.Caption = GetSetting(appname:="simptel", section:="registration", Key:="user", Default:="UNREGISTERED")
Else
lbl_regged_to.ForeColor = vbRed
lbl_regged_to.Caption = "UNREGISTERED"
End If
End Sub
