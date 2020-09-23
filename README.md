<div align="center">

## POP3 Mail Read


</div>

### Description

POP3 protocol client latest release. Usage and full source code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tair Abdurman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tair-abdurman.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tair-abdurman-pop3-mail-read__1-3318/archive/master.zip)





### Source Code

```
'WS_POP3_Conn is winsock component variable
'pop3 session state after send LIST command
'full source code and usage paper you can find at
'http://www.tair.freeservers.com
Case 4
 WS_POP3_Conn.GetData inBuffer2, vbString
 inBuffer = inBuffer & inBuffer2
If def_mail = 0 Then
'Answer on LIST command
 If Right(inBuffer, 5) = CRLF_CRLF Then
 'OK LIST response terminated
  def_mail = def_mail + 1
  Tmp_log.Text = Tmp_log.Text & "Parsing List Response" & CRLF
  If Parse_LIST_Response(inBuffer) = 0 Then
   Tmp_log.Text = Tmp_log.Text & "FOUND: " & mail_count & " mail(s)" & CRLF
   outBuffer = "RETR " & def_mail & CRLF
   Tmp_log.Text = Tmp_log.Text & "RETR " & def_mail & " COMMAND SENT" & CRLF
   Cmd_First.Enabled = False
   Cmd_Prev.Enabled = False
   Cmd_Next.Enabled = False
   Cmd_Last.Enabled = False
   Cmd_GoTo.Enabled = False
  Else
   outBuffer = "QUIT" & CRLF
   Tmp_log.Text = Tmp_log.Text & "NO MAILS FOUND" & CRLF
   Tmp_log.Text = Tmp_log.Text & "QUIT COMMAND SENT" & CRLF
   Command_ID = 5
   Tmp_log.Text = Tmp_log.Text & "cid=5" & CRLF
  End If
  Tmp_log.Text = Tmp_log.Text & "ibuffer=" & inBuffer
  Tmp_log.Text = Tmp_log.Text & "obuffer=" & outBuffer
  inBuffer = ""
  WS_POP3_Conn.SendData outBuffer
  Tmp_log.SelStart = Len(Tmp_log.Text) - 1
  Tmp_log.Refresh
 'EOF OK LIST response terminated
 End If
'EOF Answer on LIST command
Else
 If def_mail < mail_count Then
'recive n mail
 If Right(inBuffer, 5) = CRLF_CRLF Then
 'OK n mail terminated
  zu = Parse_Mail(inBuffer, def_mail)
  def_mail = def_mail + 1
  outBuffer = "RETR " & def_mail & CRLF
 Tmp_log.Text = Tmp_log.Text & "RETR " & def_mail & " COMMAND SENT" & CRLF
  inBuffer = ""
  WS_POP3_Conn.SendData outBuffer
  Tmp_log.SelStart = Len(Tmp_log.Text) - 1
  Tmp_log.Refresh
  'ok n mail recived
  'EOF ok n mail recived
  'Else
  'fail n mail not recived
  'EOF fail n mail not recived
  'End If
 'EOF OK n mail terminated
 End If
'EOF recive n mail
 Else
'recive last mail
 If Right(inBuffer, 5) = CRLF_CRLF Then
 'OK last mail terminated
  'If Left(inBuffer, 1) = "+" Then
  'ok last mail recived no errors
  zu = Parse_Mail(inBuffer, def_mail)
  Tmp_log.Text = Tmp_log.Text & "cid=5" & CRLF
  Tmp_log.Text = Tmp_log.Text & "Get Last Mail" & CRLF
  Tmp_log.Text = Tmp_log.Text & "ibuffer=" & inBuffer
  Tmp_log.Text = Tmp_log.Text & "obuffer=" & outBuffer
  outBuffer = "QUIT" & CRLF
  Tmp_log.Text = Tmp_log.Text & "QUIT COMMAND SENT" & CRLF
  Command_ID = 5
  Tmp_log.Text = Tmp_log.Text & "cid=5" & CRLF
  inBuffer = ""
  If mail_count > 1 Then
   Cmd_First.Enabled = False
   Cmd_Prev.Enabled = False
   Cmd_Next.Enabled = True
   Cmd_Last.Enabled = True
   Cmd_GoTo.Enabled = True
  End If
  Lbl_Mail_Count.Caption = "of " & mail_count
  Lbl_Mail_Count.Refresh
  Load_Fields 1
  txt_Position.Text = "1"
  txt_Position.Refresh
  WS_POP3_Conn.SendData outBuffer
  Tmp_log.Text = Tmp_log.Text & "QUIT COMMAND SENT" & CRLF
  Tmp_log.SelStart = Len(Tmp_log.Text) - 1
  Tmp_log.Refresh
  'EOF ok last mail recived no errors
  'Else
  'last mail recived with errors
  ' MsgBox "last mail recived with errors."
  ' Command_ID = 5
  'EOF last mail recived with errors
  'End If
 'EOF OK last mail terminated
 End If
 'recive last mail
 End If
End If
```

