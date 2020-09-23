VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Centennial 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Centennial.ctx":0000
   ScaleHeight     =   510
   ScaleWidth      =   465
   ToolboxBitmap   =   "Centennial.ctx":0CCE
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1560
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "centennialwireless.com"
      RemotePort      =   80
   End
End
Attribute VB_Name = "Centennial"
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
Dim CallBack As String
Dim WholeData As String
Public Event SockError(Description As String)
Public Event SockConnect()
Public Event MessageSent()
Public Event CouldntSendMessage()
Public Event ReceivingData()
Private Sub UserControl_Resize()
UserControl.Width = 550
UserControl.Height = 550
End Sub
Private Sub Winsock1_Connect()
RaiseEvent SockConnect
Winsock1.SendData "GET /home/sms.php?deviceid=" & PhoneNumber & "&mess=" & Message & "&yournumber=" & CallBack & "&submit=Send+Message&carrier=CENTENNIAL+COMMUNICATIONS&clientname=CENTENNIAL&TOTAL=9" & vbCrLf & vbCrLf
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
Sub SendMessage(CentennialPhoneNumber As String, TextMessage As String, CallbackNumber As String)
WholeData = ""
Winsock1.Close
Winsock1.Connect
PhoneNumber = CentennialPhoneNumber
Message = TextMessage
Message = ReplaceChars(Message)
CallBack = CallbackNumber
End Sub
Sub Disconnect()
Winsock1.Close
Winsock2.Close
WholeData = ""
End Sub
Private Sub Winsock1_Close()
Winsock1.Close
If InStr(1, WholeData, "Your Message has been sent!") Then GotIt: Exit Sub
RaiseEvent CouldntSendMessage
WholeData = ""
End Sub
Private Sub GotIt()
RaiseEvent MessageSent
WholeData = ""
End Sub
Function ReplaceChars(Status As String)
On Error Resume Next
Dim YELLO As String
YELLO = Replace(Status, "%", "%25")
Status = Replace(YELLO, ":", "%3A")
Status = Replace(YELLO, "`", "%60")
Status = Replace(YELLO, "~", "%7E")
Status = Replace(YELLO, "!", "%21")
Status = Replace(YELLO, "#", "%23")
Status = Replace(YELLO, "$", "%24")
Status = Replace(YELLO, "^", "%5E")
Status = Replace(YELLO, "&", "%26")
Status = Replace(YELLO, "(", "%28")
Status = Replace(YELLO, ")", "%29")
Status = Replace(YELLO, "+", "%2B")
Status = Replace(YELLO, "{", "%7B")
Status = Replace(YELLO, "}", "%7D")
Status = Replace(YELLO, "[", "%5B")
Status = Replace(YELLO, "]", "%5D")
Status = Replace(YELLO, Chr(34), "%22")
Status = Replace(YELLO, "'", "%27")
Status = Replace(YELLO, "?", "%3F")
Status = Replace(YELLO, "/", "%2F")
Status = Replace(YELLO, "\", "%5C")
Status = Replace(YELLO, "<", "%3C")
Status = Replace(YELLO, ">", "%3E")
Status = Replace(YELLO, "|", "%7C")
Status = Replace(YELLO, ",", "%2C")
Status = Replace(YELLO, vbCrLf, "%3Cbr%3E")
ReplaceChars = Status
ReplaceChars = Replace(ReplaceChars, " ", "%20")
End Function
