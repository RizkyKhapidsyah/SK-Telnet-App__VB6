VERSION 5.00
Begin VB.Form frmREGISTER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register your copy"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdREGISTER 
      Caption         =   "Register"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtCODE 
      Height          =   285
      Left            =   120
      MaxLength       =   9
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtNAME 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Your registration code:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Your Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmREGISTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCANCEL_Click()
txtNAME.Text = ""
txtCODE.Text = ""
Me.Hide
End Sub
Private Sub cmdREGISTER_Click()
If (func.is_valid_reg(username:=CStr(txtNAME.Text), regcode:=Val(txtCODE.Text)) = True) Then
    SaveSetting "simptel", "registration", "user", txtNAME.Text
    SaveSetting "simptel", "registration", "code", txtCODE.Text
    MsgBox "Thanks for registering Simple Telnet for Windows!" & vbCrLf & _
    "The program needs to be restarted for registration to complete" & vbCrLf & _
    "Press OK to quit now", vbInformation, "Registration"
    frmMAIN.menu_file_exit_Click
Else
    MsgBox "Wrong code!", vbExclamation, "Registration"
    Dim username As String, serial As Double
    username = GetSetting(appname:="simptel", section:="registration", Key:="user", Default:="")
    serial = GetSetting(appname:="simptel", section:="registration", Key:="code", Default:="0")
    If ((username <> "") And (serial <> 0)) Then
        DeleteSetting "simptel", "registration", "user"
        DeleteSetting "simptel", "registration", "code"
    End If
End If
End Sub
Private Sub txtCODE_Change()
If ((txtNAME.Text <> "") And (txtCODE.Text <> "")) Then
cmdREGISTER.Enabled = True
Else
cmdREGISTER.Enabled = False
End If
End Sub
Private Sub txtNAME_Change()
If ((txtNAME.Text <> "") And (txtCODE.Text <> "")) Then
cmdREGISTER.Enabled = True
Else
cmdREGISTER.Enabled = False
End If
End Sub
