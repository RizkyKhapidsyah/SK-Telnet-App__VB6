VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCONNECT2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect..."
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCONNECT2.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCANCEL 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtPORT 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   6
      Top             =   890
      Width           =   1335
   End
   Begin VB.TextBox txtHOST 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      MaxLength       =   64
      TabIndex        =   4
      Top             =   290
      Width           =   2655
   End
   Begin VB.CommandButton cmdDEL 
      Caption         =   "Del"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "Add"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1335
   End
   Begin MSComctlLib.ListView conn_list 
      Height          =   2130
      Left            =   50
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   50
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3757
      View            =   2
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdCONNECT 
      Caption         =   "Connect"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1325
   End
   Begin VB.Label Label2 
      Caption         =   "port number:"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   645
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "hostname:"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   45
      Width           =   2655
   End
End
Attribute VB_Name = "frmCONNECT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdADD_Click()
If (txtHOST.Text = "" Or Val(txtPORT.Text) = 0) Then
Exit Sub
End If
conn_list.Enabled = True
cmdDEL.Enabled = True
Dim connections As Variant, cnt As Long
connections = GetAllSettings(appname:="simptel", section:="connections")
On Error GoTo ERROR_HANDLER
For cnt = LBound(connections, 1) To UBound(connections, 1)
If (connections(cnt, 0) = txtHOST.Text) Then
MsgBox "This host is already added!", vbCritical, "Error adding host"
Exit Sub
End If
Next cnt
SaveSetting "simptel", "connections", txtHOST.Text, Val(txtPORT.Text)
conn_list.ListItems.Add , , txtHOST.Text, 1, 1
cmdDEL.Enabled = True
Exit Sub
ERROR_HANDLER:
SaveSetting "simptel", "connections", txtHOST.Text, Val(txtPORT.Text)
conn_list.ListItems.Add , , txtHOST.Text, 1, 1
End Sub
Private Sub cmdCANCEL_Click()
'Cleaning up the listview so that next time we begin to load it is empty
Dim cnt As Long, cnt2 As Long
cnt2 = conn_list.ListItems.count
For cnt = 1 To conn_list.ListItems.count
conn_list.ListItems.Remove (cnt2)
cnt2 = cnt2 - 1
Next cnt
cmdDEL.Enabled = False
conn_list.Enabled = False
Me.Hide
End Sub
Private Sub cmdCONNECT_Click()
frmMAIN.winsck.Close
frmMAIN.winsck.RemoteHost = CStr(txtHOST.Text)
If (Val(txtPORT.Text) > 65535) Then
frmMAIN.winsck.RemotePort = (Val(txtPORT.Text) - 65535)
Else
frmMAIN.winsck.RemotePort = Val(txtPORT.Text)
End If
frmMAIN.winsck.Connect
Dim cnt As Long, cnt2 As Long
cnt2 = conn_list.ListItems.count
For cnt = 1 To conn_list.ListItems.count
conn_list.ListItems.Remove (cnt2)
cnt2 = cnt2 - 1
Next cnt
cmdDEL.Enabled = False
conn_list.Enabled = False
Me.Hide
func.send_to_buffer ("Attempting connection to: " & frmMAIN.winsck.RemoteHost & ":" & frmMAIN.winsck.RemotePort)
End Sub

Private Sub cmdDEL_Click()
DeleteSetting "simptel", "connections", conn_list.SelectedItem
conn_list.ListItems.Remove (conn_list.SelectedItem.Index)
If (conn_list.ListItems.count = 0) Then
txtHOST.Text = ""
txtPORT.Text = ""
cmdDEL.Enabled = False
conn_list.Enabled = False
Else
conn_list.ListItems.Item(conn_list.ListItems.count).Selected = True
txtHOST.Text = conn_list.SelectedItem
txtPORT.Text = Val(GetSetting(appname:="simptel", section:="connections", Key:=conn_list.SelectedItem, Default:=0))
End If
End Sub

Private Sub conn_list_Click()
txtHOST.Text = conn_list.SelectedItem
txtPORT.Text = Val(GetSetting(appname:="simptel", section:="connections", Key:=conn_list.SelectedItem, Default:="0"))
End Sub

Private Sub Form_Activate()
Dim connections As Variant, cnt As Long
connections = GetAllSettings(appname:="simptel", section:="connections")
On Error GoTo ERROR_HANDLER
For cnt = LBound(connections, 1) To UBound(connections, 1)
conn_list.ListItems.Add , , connections(cnt, 0), 1, 1
conn_list.Enabled = True
cmdDEL.Enabled = True
Next cnt
conn_list.ListItems.Item(1).Selected = True
txtHOST.Text = conn_list.SelectedItem
txtPORT.Text = Val(GetSetting(appname:="simptel", section:="connections", Key:=conn_list.SelectedItem, Default:=0))
ERROR_HANDLER:
Exit Sub
End Sub

Private Sub txtHOST_Change()
If ((txtHOST.Text = "") Or (Val(txtPORT.Text) = 0)) Then
cmdCONNECT.Enabled = False
Else
cmdCONNECT.Enabled = True
End If
End Sub

Private Sub txtHOST_GotFocus()
txtHOST.SelStart = 0
txtHOST.SelLength = Len(txtHOST.Text)
End Sub

Private Sub txtPORT_Change()
If ((txtHOST.Text = "") Or (Val(txtPORT.Text) = 0)) Then
cmdCONNECT.Enabled = False
Else
cmdCONNECT.Enabled = True
End If
End Sub

Private Sub txtPORT_GotFocus()
txtPORT.SelStart = 0
txtPORT.SelLength = Len(txtHOST.Text)
End Sub
