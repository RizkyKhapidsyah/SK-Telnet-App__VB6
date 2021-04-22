VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMAIN 
   Caption         =   "Telnet"
   ClientHeight    =   4755
   ClientLeft      =   3390
   ClientTop       =   3795
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8775
   Begin MSWinsockLib.Winsock winsck 
      Left            =   5520
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "log.txt"
      Filter          =   ".txt"
      Orientation     =   2
   End
   Begin VB.TextBox txtCMD 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      MaxLength       =   512
      TabIndex        =   1
      Top             =   4470
      Width           =   8775
   End
   Begin VB.TextBox txtSCREEN 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   32768
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
   End
   Begin VB.Menu menu_file 
      Caption         =   "&File"
      Begin VB.Menu menu_file_connect 
         Caption         =   "&Connect..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu menu_file_disconnect 
         Caption         =   "&Disconnect"
         Shortcut        =   {F2}
      End
      Begin VB.Menu menu_file_null_3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_file_save 
         Caption         =   "&Save log"
         Shortcut        =   {F3}
      End
      Begin VB.Menu menu_file_clear 
         Caption         =   "C&lear"
         Shortcut        =   {F5}
      End
      Begin VB.Menu menu_file_null_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_file_font 
         Caption         =   "&Font..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu menu_file_null_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu_file_echo 
         Caption         =   "&Echo?"
         Checked         =   -1  'True
         Shortcut        =   {F6}
      End
      Begin VB.Menu menu_file_secure 
         Caption         =   "Sec&ure?"
         Shortcut        =   {F7}
      End
      Begin VB.Menu menu_file_null_1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_file_exit 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu menu_edit 
      Caption         =   "&Edit"
      Begin VB.Menu menu_edit_cut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu menu_edit_copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu menu_edit_paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu menu_edit_null_1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_edit_del 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu menu_help 
      Caption         =   "&Help"
      Begin VB.Menu menu_help_reg 
         Caption         =   "&Register..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu menu_about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public item_got_focus As Byte
Public is_regged As Boolean
Private Sub Form_GotFocus()
txtCMD.SetFocus
End Sub
Private Sub Form_Load()
Dim form_left As Long, form_top As Long, form_height As Long, form_width As Long
form_left = GetSetting(appname:="simptel", section:="position", Key:="Left", Default:=0)
form_top = GetSetting(appname:="simptel", section:="position", Key:="Top", Default:=0)
form_height = GetSetting(appname:="simptel", section:="position", Key:="Height", Default:=5445)
form_width = GetSetting(appname:="simptel", section:="position", Key:="Width", Default:=8895)
If (form_top = 0) Then
Me.Top = (Screen.Height / 2) - (Me.Height / 2)
Else
Me.Top = form_top
End If
If (form_left = 0) Then
Me.Left = (Screen.Width / 2) - (Me.Width / 2)
Else
Me.Left = form_left
End If
Me.Height = form_height
Me.Width = form_width
Dim username As String, serial As Double
username = GetSetting(appname:="simptel", section:="registration", Key:="user", Default:="")
serial = GetSetting(appname:="simptel", section:="registration", Key:="code", Default:="0")
If ((username <> "") And (serial <> 0)) Then
    If (func.is_valid_reg(username:=CStr(username), regcode:=Val(serial)) = True) Then
        menu_help_reg.Enabled = False
        is_regged = True
        Else
        menu_help_reg.Enabled = True
        is_regged = False
        DeleteSetting "simptel", "registration", "user"
        DeleteSetting "simptel", "registration", "code"
    End If
    Else
    menu_help_reg.Enabled = True
    is_regged = False
    End If
txtSCREEN.FontName = GetSetting(appname:="simptel", section:="fonts", Key:="fontname", Default:="Lucida Console")
txtSCREEN.FontSize = GetSetting(appname:="simptel", section:="fonts", Key:="fontsize", Default:=9)
menu_file_echo.Checked = GetSetting(appname:="simptel", section:="options", Key:="Echo", Default:=1)
menu_file_secure.Checked = GetSetting(appname:="simptel", section:="options", Key:="Secure", Default:=0)
End Sub
Private Sub Form_Resize()
If (ScaleWidth < 3000) Then
frmMAIN.Enabled = False
frmMAIN.Width = 3115
Else
frmMAIN.Enabled = True
End If
If (ScaleHeight < 1605) Then
frmMAIN.Enabled = False
frmMAIN.Height = 2290
Else
frmMAIN.Enabled = True
End If
txtSCREEN.Move 0, 0, ScaleWidth, ScaleHeight - 300
txtCMD.Move 0, txtSCREEN.Height + 25, ScaleWidth, 265
End Sub
Private Sub Form_Unload(Cancel As Integer)
menu_file_exit_Click
End Sub
Private Sub menu_about_Click()
frmABOUT.Show 1
End Sub
Private Sub menu_edit_copy_Click()
Select Case item_got_focus
Case Is = 1
Clipboard.SetText (txtSCREEN.SelText)
Case Is = 2
Clipboard.SetText (txtCMD.Text)
Case Else
MsgBox "BUG DETECTED!"
End Select
End Sub
Private Sub menu_edit_cut_Click()
Clipboard.SetText (txtCMD.SelText)
txtCMD.SelText = ""
End Sub
Private Sub menu_edit_del_Click()
If (txtCMD.SelLength = 0) Then
txtCMD.SelLength = 1
txtCMD.SelText = ""
Else
txtCMD.SelText = ""
End If
End Sub
Private Sub menu_edit_paste_Click()
txtCMD.SetFocus
txtCMD.SelText = Clipboard.GetText
End Sub
Private Sub menu_file_clear_Click()
txtSCREEN.Text = ""
End Sub
Private Sub menu_file_connect_Click()
frmCONNECT2.Show 1
End Sub
Private Sub menu_file_disconnect_Click()
If (winsck.State <> 0) Then
winsck.Close
func.send_to_buffer ("Connection to " & winsck.RemoteHost & ":" & winsck.RemotePort & " closed.")
End If
frmMAIN.Caption = "Telnet"
End Sub
Private Sub menu_file_echo_Click()
If (menu_file_echo.Checked = False) Then
menu_file_echo.Checked = True
func.send_to_buffer ("ECHO is now ON")
Else
menu_file_echo.Checked = False
func.send_to_buffer ("ECHO is now OFF")
End If
End Sub
Public Sub menu_file_exit_Click()
'Saving form positioning setting
SaveSetting "simptel", "position", "Left", frmMAIN.Left
SaveSetting "simptel", "position", "Top", frmMAIN.Top
SaveSetting "simptel", "position", "Height", frmMAIN.Height
SaveSetting "simptel", "position", "Width", frmMAIN.Width
SaveSetting "simptel", "options", "Echo", menu_file_echo.Checked
SaveSetting "simptel", "options", "Secure", menu_file_secure.Checked
End
End Sub
Private Sub menu_file_font_Click()
With CommonDialog1
.FontName = txtSCREEN.Font
.FontSize = txtSCREEN.FontSize
.Flags = &H400 + &H3 + &H4000 + &H10000 + &H2000
.Min = 9
.Max = 16
End With
On Error GoTo ERR_FONT
CommonDialog1.ShowFont
With txtSCREEN
.Font = CommonDialog1.FontName
.FontSize = CommonDialog1.FontSize
SaveSetting "simptel", "fonts", "fontname", txtSCREEN.FontName
SaveSetting "simptel", "fonts", "fontsize", txtSCREEN.FontSize
End With
Exit Sub
ERR_FONT:
Exit Sub
End Sub
Private Sub menu_file_save_Click()
Dim save_path As String, logsize As Long
On Error GoTo ERR_HANDLER
CommonDialog1.ShowSave
save_path = CommonDialog1.FileName
If (save_path <> "") Then
Open save_path For Binary As #1
Put #1, , txtSCREEN.Text
Close #1
logsize = Len(txtSCREEN.Text)
func.send_to_buffer ("Log written to: " & CommonDialog1.FileName)
func.send_to_buffer ("Written " & logsize & " bytes to disk")
End If
ERR_HANDLER:
Exit Sub
End Sub
Private Sub menu_file_secure_Click()
If (menu_file_secure.Checked = True) Then
menu_file_secure.Checked = False
func.send_to_buffer ("SECURE mode is now OFF")
txtCMD.PasswordChar = ""
Else
menu_file_secure.Checked = True
func.send_to_buffer ("SECURE mode is now ON")
txtCMD.PasswordChar = " "
End If
End Sub

Private Sub menu_help_reg_Click()
frmREGISTER.Show 1

End Sub

Private Sub txtCMD_Change()
If (menu_file_secure.Checked = True) Then
txtCMD.PasswordChar = " "
Else
txtCMD.PasswordChar = ""
End If
End Sub
Private Sub txtCMD_GotFocus()
item_got_focus = 2
menu_edit_cut.Enabled = True
menu_edit_del.Enabled = True
End Sub

Private Sub txtCMD_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (winsck.State = 0) Then
        func.send_to_buffer ("You need to connect first!")
        txtCMD.Text = ""
        Else
        Dim string_to_send As String
        string_to_send = txtCMD.Text
        If (string_to_send <> "") Then
            winsck.SendData string_to_send & vbCrLf
            func.send_to_buffer_norm (txtCMD.Text)
            txtCMD.Text = ""
            txtCMD.SetFocus
            Else
            func.send_to_buffer ("Nothing to send!")
            txtCMD.Text = ""
            txtCMD.SetFocus
        End If
    End If
End If
End Sub

Private Sub txtSCREEN_Change()
txtSCREEN.SelStart = Len(txtSCREEN.Text)
End Sub
Private Sub txtSCREEN_GotFocus()
item_got_focus = 1
menu_edit_cut.Enabled = False
menu_edit_del.Enabled = False
End Sub
Private Sub winsck_Close()
func.send_to_buffer ("Disconnected from: " & winsck.RemoteHost & ":" & winsck.RemotePort)
frmMAIN.Caption = "Telnet"
winsck.Close
End Sub
Private Sub winsck_Connect()
func.send_to_buffer ("Succeeded connection to: " & winsck.RemoteHost & ":" & winsck.RemotePort)
txtCMD.SetFocus
frmMAIN.Caption = "Connected to: " & winsck.RemoteHost & ":" & winsck.RemotePort
End Sub
Private Sub winsck_DataArrival(ByVal bytesTotal As Long)
Dim data_received As String
winsck.GetData data_received, vbString
func.send_to_buffer_getdata (data_received)
End Sub
Private Sub winsck_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Select Case number
Case Is = 7
func.send_to_buffer ("ERROR: Out of memory.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 380
func.send_to_buffer ("ERROR: The property value is invalid.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 394
func.send_to_buffer ("ERROR: The property can't be read.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 383
func.send_to_buffer ("ERROR: The property is read-only.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 1004
func.send_to_buffer ("ERROR: The operation was canceled.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10014
func.send_to_buffer ("ERROR: The requested address is a broadcast address, but flag is not set.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10035
func.send_to_buffer ("ERROR: Socket is non-blocking and the specified operation will block.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10036
func.send_to_buffer ("ERROR: A blocking Winsock operation in progress.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10037
func.send_to_buffer ("ERROR: The operation is completed. No blocking operation in progress")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10038
func.send_to_buffer ("ERROR: The descriptor is not a socket.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10040
func.send_to_buffer ("ERROR: The datagram is too large to fit into the buffer and is truncated.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10043
func.send_to_buffer ("ERROR: The specified port is not supported.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10048
func.send_to_buffer ("ERROR: Address in use.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10049
func.send_to_buffer ("ERROR: Address not available from the local machine.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10050
func.send_to_buffer ("ERROR: Network subsystem failed.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10051
func.send_to_buffer ("ERROR: The network cannot be reached from this host at this time.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10052
func.send_to_buffer ("ERROR: Connection has timed out when SO_KEEPALIVE is set.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10053
func.send_to_buffer ("ERROR: Connection is aborted due to timeout or other failure.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10054
func.send_to_buffer ("ERROR: The connection is reset by remote side.!")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10055
func.send_to_buffer ("ERROR: No buffer space is available.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10056
func.send_to_buffer ("ERROR: Socket is already connected.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10057
func.send_to_buffer ("ERROR: Socket is not connected.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10058
func.send_to_buffer ("ERROR: Socket has been shut down.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10060
func.send_to_buffer ("ERROR: Socket has been shut down.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10061
func.send_to_buffer ("ERROR: Connection is forcefully rejected.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 10093
func.send_to_buffer ("ERROR: WinsockInit should be called first.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 11001
func.send_to_buffer ("ERROR: Authoritative answer: Host not found.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 11002
func.send_to_buffer ("ERROR: Non-Authoritative answer: Host not found.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 11003
func.send_to_buffer ("ERROR: Non-recoverable errors.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Is = 11004
func.send_to_buffer ("ERROR: Valid name, no data record of requested type.")
winsck.Close
frmMAIN.Caption = "Telnet"
Case Else
send_to_buffer ("ERROR: UNKNOWN! E-Mail the log from this telnet session to vladkoacs@yahoo.com and tell me what happened!")
winsck.Close
frmMAIN.Caption = "Telnet"
End Select
End Sub
