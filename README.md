<div align="center">

## Make your own chat room in 10 minutes\!


</div>

### Description

Ever want your own chat? Ever want your own rules? This code allows you to make your own chat room! Allows 2 people to chat from anywhere in the world from any internet provider! Perfect for quick and private communication!

NOTE: This program requires mswinsck.ocx
 
### More Info
 
There are alot of things to set up in order for this code to work correctly. Here are the complete list of things you will need to do:

1. Add the mswinsck.ocx to your project

2. Create a textbox and name it txtHost     This will be the box where the remote host is entered

3. Create a textbox and name it txtLocalP.    This will be the box where the local port is entered

4. Create a textbox and name it txtRemoteP    This will be the box where the remote port is entered

5. Create a textbox and name it txtNick     This will be the box you will enter your nickname (aka screenname)

6. Create a textbox and name it txtSend     This will be the box you type in to send stuff to the chatroom

7. Create a large textbox and name it txtMain  This will be the chatroom

8. Make txtMain multiline and also add vertical scrollbars

9. Create a command button and name it cmdC. Give it the caption "Connect"

10. Create a command button and name it cmdD. Give it the caption "Disconnect"

11. Create a command button and name it cmdSend. Give it the caption "Send". This will be the button that sends the text in txtSend to the chatroom.

12. Put a winsock control on the form and name it sckSend

13. Labels can be put so you can remember which text box is which. Also, it would be best to erase all the text in the text boxes (ie: Get rid of Text1 written in the box)

When 2 people have the program running, this is how you connect:

1. First, enter a nickname in the txtNick box. This is the name that will come before what you say in the chatroom.

2. In the txtHost textbox, you must put the other person's IP address or hostname.

3. You and the other person must both think up any number to be your local port (Just make sure they're different numbers. ie: My local port is 1000, my friend's local port is 2000)

4. After you've both made up your local host and entered it in the txtLocalP textbox, you must next enter the other person's local port in your txtRemoteP textbox. For example, My local port is 1000. and my remote port is 2000... My friend's local port would be 2000 and my friends's remote port would be 1000.

5. Both of you must now click the connect button.

7. Now just type text in the txtSend textbox and click the "Send" button. You will notice the text moved into the chatroom. This text is visible by both people! Congratulations!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chaz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chaz.md)
**Level**          |Unknown
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chaz-make-your-own-chat-room-in-10-minutes__1-1029/archive/master.zip)





### Source Code

```
Private Sub cmdC_Click()
   If Len(txtNick) < 1 Then 'make sure there is a nickname entered
     MsgBox "You must enter a nickname first!"
     txtNick.SetFocus 'put the cursor in the nickname textbox
     Exit Sub
   End If
   If Len(txtHost) < 1 Or Len(txtLocalP) < 1 Or Len(txtRemoteP) < 1 Then
    MsgBox "Please make sure a Host, a Local Port, and a Remote Port have been entered!"
    Exit Sub
   End If
   sckSend.RemoteHost = txtHost   'set the host
   sckSend.LocalPort = txtLocalP   'set the local port
   sckSend.RemotePort = txtRemoteP  'set the remote port
   sckSend.Bind 'Connect!
   cmdSend.Enabled = True 'Enable the send button
   txtNick.Enabled = False 'Make it so you can't change your nickname
   txtSend.SetFocus   'you have been connected. put the cursor in the send textbox
End Sub
Private Sub cmdD_Click()
'The disconnect button was pushed.
End
End Sub
Private Sub cmdSend_Click()
'The Send button was pushed
sckSend.SendData txtNick.Text & ": " & txtSend.Text & Chr$(13) & Chr$(10) 'Send whatever is wrtten in txtSend to the other person's chatroom.
txtMain.Text = txtMain.Text & txtNick.Text & ": " & txtSend.Text & Chr$(13) & Chr$(10) 'Put it in your chatroom
txtMain.SelStart = Len(txtMain) 'scroll that chatroom down
txtSend.Text = "" 'clear the send textbox
End Sub
Private Sub Form_Load()
sckSend.Protocol = sckUDPProtocol 'set protocol. For this type of chat, we are using UDP
cmdSend.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub sckSend_DataArrival(ByVal bytesTotal As Long)
'We have received data!
Dim TheData As String
On Error GoTo ClearChat
sckSend.GetData TheData, vbString 'extract the data
txtMain.Text = txtMain.Text & TheData 'add the data to our chatroom
txtMain.SelStart = Len(txtMain) 'scroll that chatroom down
Exit Sub
ClearChat:
MsgBox "Chat room ran out of memory and must be cleared!"
txtMain.Text = ""
End Sub
Private Sub sckSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "An error occurred in winsock!"
End
End Sub
```

