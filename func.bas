Attribute VB_Name = "func"
Public Function send_to_buffer(text_to_display As String)
If (Len(frmMAIN.txtSCREEN.Text) = 0) Then
frmMAIN.txtSCREEN.Text = "*** " & text_to_display & vbCrLf
Else
frmMAIN.txtSCREEN.Text = frmMAIN.txtSCREEN.Text & vbCrLf & "*** " & text_to_display & vbCrLf & vbCrLf
End If
End Function
Public Function send_to_buffer_norm(text_to_input As String)
If frmMAIN.menu_file_echo.Checked = True Then
If (Len(frmMAIN.txtSCREEN.Text) = 0) Then
frmMAIN.txtSCREEN.Text = "> " & text_to_input
Else
frmMAIN.txtSCREEN.Text = frmMAIN.txtSCREEN.Text & "> " & text_to_input & vbCrLf
End If
End If
End Function
Public Function send_to_buffer_getdata(text_to_show As String)
If (Len(frmMAIN.txtSCREEN.Text) = 0) Then
frmMAIN.txtSCREEN.Text = text_to_show & vbCrLf
Else
frmMAIN.txtSCREEN.Text = frmMAIN.txtSCREEN.Text & text_to_show
End If
End Function
Public Function is_valid_reg(username As String, regcode As Long) As Boolean
Dim final_code As Double
final_code = generate_key(username)
If (final_code = regcode) Then
is_valid_reg = 1
Else
is_valid_reg = 0
End If
End Function
Public Function generate_key(username As String) As Double
Dim answer As Boolean, pre_code As Double, count As Long, i As Byte, something As Double
i = Len(username)
pre_code = 0
For count = 1 To i
pre_code = pre_code + Asc(Mid(username, count, 1))
Next count
something = (pre_code * (pre_code \ 2)) * 4
generate_key = something
End Function
