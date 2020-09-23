VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl PCS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   InvisibleAtRuntime=   -1  'True
   Picture         =   "PCS.ctx":0000
   ScaleHeight     =   570
   ScaleWidth      =   750
   ToolboxBitmap   =   "PCS.ctx":1806
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "65.174.43.73"
      RemotePort      =   80
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "65.174.43.73"
      RemotePort      =   80
   End
End
Attribute VB_Name = "PCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This user control was created by Garren Fitzenreiter
' If you use this, please give me credit
' http://www.keithware.com
' 2005

Dim PhoneNumber As String
Dim Message As String
Dim Callback As String
Dim WholeData As String
Dim WholeData2 As String
Dim sTrackNumber As String
Public Event SockError(Description As String)
Public Event SockConnect()
Public Event MessageSent(TrackingNumber As String)
Public Event CouldntSendMessage()
Public Event ReceivingData()
Public Event CouldntTrackMessage()
Public Event GotTrackingStatus(TheStatus As String, SentTime As String, ReceiveTime As String, FromSource As String, ToDestination As String, MessageSent As String)
Const Cookie As String = "JSESSIONID=Btx13CFA3R14Qug6sNpo52GQhRbdDZKgg6Px1eHYhXgu11qfJvmo!1445420257!182751416!5070!7002; pcsCustomer=customer=yes"
Const Cookie2 As String = "JSESSIONID=B65Q21WYeCPgXPw9VgZg6UAlj2wUEFI7N1Tu9UWGU8ZETgG979c1!1166028299!182751433!5070!7002; pcsCustomer=customer=yes"
Sub SendMessage(PCSPhoneNumber As String, TextMessage As String, CallbackNumber As String)
WholeData = ""
Winsock1.Close
Winsock1.Connect
PhoneNumber = PCSPhoneNumber
Message = TextMessage
Callback = CallbackNumber
End Sub
Sub TrackMessage(TrackingNumber As String)
Winsock2.Close
Winsock2.Connect
sTrackNumber = TrackingNumber
End Sub
Sub Disconnect()
Winsock1.Close
Winsock2.Close
WholeData2 = ""
WholeData = ""
End Sub
Private Sub UserControl_Resize()
UserControl.Width = 765
UserControl.Height = 570
End Sub
Private Sub Winsock1_Close()
On Error GoTo ErrorSending
Winsock1.Close
RaiseEvent MessageSent(TrackNumber(WholeData))
Exit Sub
ErrorSending:
RaiseEvent CouldntSendMessage
End Sub
Private Sub Winsock1_Connect()
RaiseEvent SockConnect
Dim Content As String
Dim ConLength As Long
Content = "randomToken=&phoneNumber=" & PhoneNumber & "&message=" & Message & "&callBackNumber=" & Callback & "&x=31&y=10"
ConLength = Len(Content)
Winsock1.SendData "POST /textmessaging/composeconfirm HTTP/1.1" & vbCrLf & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, */*" & vbCrLf & "Referer: http://messaging.sprintpcs.com/textmessaging/compose" & vbCrLf & "Accept-Language: en-us" & vbCrLf & "Content-Type: application/x-www-form-urlencoded" & vbCrLf & "Accept-Encoding: gzip, deflate" & vbCrLf & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.1.4322)" & vbCrLf & "Host: messaging.sprintpcs.com" & vbCrLf & "Content-Length: " & ConLength & vbCrLf & "Connection: Keep-Alive" & vbCrLf & "Cache-Control: no-cache" & vbCrLf & "Cookie: " & Cookie & vbCrLf & vbCrLf & Content
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock1.GetData Data
RaiseEvent ReceivingData
WholeData = WholeData & Data
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
RaiseEvent SockError(Description)
WholeData = ""
End Sub
Function Get_Between(Start As Integer, Data As String, Before As String, After As String)
On Error GoTo ErrorHandler
Dim Before1, After1 As Integer
Before1 = InStr(Start, Data, Before)
After1 = InStr(Before1 + Len(Before), Data, After)
Get_Between = Mid(Data, Before1, After1 - Before1)
Get_Between = Right(Get_Between, Len(Get_Between) - Len(Before))
Exit Function
ErrorHandler:
Err.Raise 1, , "Cannot Find String"
End Function
Function TrackNumber(Data As String)
TrackNumber = Get_Between(1, Data, "<a href=" & Chr(34) & "trackresults?trackNumber=", Chr(34) & ">")
End Function
Private Sub Winsock2_Close()
On Error GoTo BadTrack
WholeData2 = Replace(WholeData2, Chr(9), "")
If InStr(1, WholeData2, "<li>The tracking number is invalid. Please ensure that you have the right number and try again.") Then WholeData2 = "": RaiseEvent CouldntTrackMessage: Exit Sub
If InStr(1, WholeData2, "Invalid fields are marked with a (<strong>!</strong>) below. <br/>One of the following situations has occurred:") Then WholeData2 = "": RaiseEvent CouldntTrackMessage: Exit Sub
Winsock2.Close
RaiseEvent GotTrackingStatus(GetTrackStatus(WholeData2), GetTheStatus(WholeData2, "Sent"), GetOtherStatus(WholeData2, "Received"), GetTheStatus(WholeData2, "From"), GetTheStatus(WholeData2, "To"), GetTheStatus(WholeData2, "Message"))
WholeData2 = ""
Exit Sub
BadTrack:
RaiseEvent CouldntTrackMessage
WholeData2 = ""
End Sub
Private Sub Winsock2_Connect()
Winsock2.SendData "GET http://messaging.sprintpcs.com/textmessaging/trackresults?trackNumber=" & sTrackNumber & " HTTP/1.1" & vbCrLf & "Accept: */*" & vbCrLf & "Accept-Language: en-us" & vbCrLf & "Accept-Encoding: gzip, deflate" & vbCrLf & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; .NET CLR 1.1.4322)" & vbCrLf & "Host: messaging.sprintpcs.com" & vbCrLf & "Connection: Keep-Alive" & vbCrLf & "Cookie: " & Cookie2 & vbCrLf & vbCrLf
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock2.GetData Data
WholeData2 = WholeData2 & Data
End Sub
Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock2.Close
RaiseEvent SockError(Description)
WholeData2 = ""
End Sub
Function GetTrackStatus(Data As String)
GetTrackStatus = Get_Between(1, Data, "<p>", "</p>")
End Function
Function GetTheStatus(Data As String, WhichStatus As String)
On Error Resume Next
GetTheStatus = Get_Between(1, Data, "<strong>" & WhichStatus & "</strong>:</td>" & Chr(10) & Chr(10) & "<td class=""right"">", "</td>")
End Function
Function GetOtherStatus(Data As String, WhichStatus As String)
On Error Resume Next
GetOtherStatus = Get_Between(1, Data, "<strong>" & WhichStatus & "</strong>:</td>" & Chr(10) & Chr(10) & " <td class=""right"">", "</td>")
End Function
