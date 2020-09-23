<div align="center">

## Very basically FTP\-Client


</div>

### Description

Connects with a FTP-Server and transfers a file to it. I use it to transfer my IP to my HP every time I connect to the net
 
### More Info
 
Server, Password

Place two winsock-elements on a form and paste the code into it

You have to know some basic things about FTP-connections, the data transferred to or from the server is send over a second connection, called data-connecton. The first connection is called control-connection, over it you send your requests and login. If you want to kinow more about it, read RFC 959 ( search with Yahoo for it)

The code doesn't function without some editing, you have to change the FTP-Server and Password, please read my comments in the code


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Philipp 'Schlaubi' Stern](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/philipp-schlaubi-stern.md)
**Level**          |Unknown
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/philipp-schlaubi-stern-very-basically-ftp-client__1-1612/archive/master.zip)





### Source Code

```
Dim x1 As Long
Dim x2 As Long
Dim last As String
Private Sub Command1_Click()
Winsock2.RemoteHost = "" 'Enter Server here
Winsock2.RemotePort = 21 ' Usually the port is 21, but if it's different, enter it here
Winsock2.Connect
Do Until Winsock2.State = sckConnected ' Wait until connected
DoEvents
Debug.Print Winsock2.State
Loop
Winsock2.SendData "USER " & vbCrLf 'Enter username behind USER
last = ""
Do Until last <> "" 'Wait until server responds
DoEvents
Loop
Winsock2.SendData "PASS " & vbCrLf 'Enter password behind PASS
last = ""
Do Until last <> "" 'Wait until server responds
DoEvents
Loop
Randomize
x1 = Int(10 * Rnd + 1) ' Find two random numbers to specify port the server connects to
Randomize
x2 = Int(41 * Rnd + 10)
Dim ip As String
ip = Winsock2.LocalIP
Do Until InStr(ip, ".") = 0 ' replace every "." in IP with a ","
  ip = Mid(ip, 1, InStr(ip, ".") - 1) & "," & Mid(ip, InStr(ip, ".") + 1)
Loop
Winsock2.SendData "PORT " & ip & "," & Trim(Str(x1)) & "," & Trim(Str(x2)) & vbCrLf 'Tell the server with which IP he has to connect and with which port
last = ""
Do Until last <> "" 'Wait until server responds
DoEvents
Loop
Winsock1.Close
Winsock1.LocalPort = x1 * 256 Or x2 ' Set port of second winsock-control to the port the server will connect to
' x1 is the most-significant byte of the port number, x1 is the least significant byte. To find the port, you have to move every bit 8 places to the right (or multiply with 256). Then compare every bit with the bits of x2, using OR
Winsock1.Listen 'Listen for the FTP-Server to connect
Winsock2.SendData "STOR ich.html" & vbCrLf 'Store a file, with RETR you can get a file, with LIST you get a list of all file on the server, all this information is sent through the data-connection (to change directory use CWD)
Do Until Winsock1.State = sckConnected 'Wait until the FTP-Server connects
DoEvents
Loop
Pause 1 'wait a little bit, because the server needs a moment (don't know how, but it only works so)
Winsock1.SendData "TEST" 'Send some data, the FTP-Server will store it in the file. Send only ASCII data, if you send Binary you have to tell it the server before, use TYPE to do this
Pause 1
Winsock1.Close ' Close data-connection
Pause 1
Winsock2.Close 'You don't have to close the connection here, you also can transfer another file
End Sub
Public Sub Pause(Seconds)
Dim Zeit As Long
Zeit = Timer
Do
DoEvents
Loop Until Zeit + Seconds <= Timer
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Winsock1.GetData data
Debug.Print data
Winsock1.Close ' You have to close the connection after the Server had send you data, he will establish it again, when he sends more
Winsock1.Listen
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Winsock2.GetData data
Debug.Print data
last = data 'Store data
End Sub
```

