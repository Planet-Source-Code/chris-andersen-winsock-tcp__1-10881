VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "iBBS Client"
   ClientHeight    =   7935
   ClientLeft      =   4050
   ClientTop       =   7380
   ClientWidth     =   12315
   LinkTopic       =   "MDIForm1"
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   3240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnudisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnuchat 
         Caption         =   "Chat"
      End
      Begin VB.Menu mnufiles 
         Caption         =   "Files"
      End
      Begin VB.Menu mnuim 
         Caption         =   "Instant Message"
      End
      Begin VB.Menu mnumb 
         Caption         =   "Message Forum"
      End
      Begin VB.Menu mnumail 
         Caption         =   "Check Mailbox"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()


sckClient.RemoteHost = Form1.Text1.Text
sckClient.RemotePort = "1001"
sckClient.Connect



End Sub

Private Sub mnuchat_Click()

Load chatform
chatform.Height = 8415
chatform.Width = 6705
chatform.Show

End Sub

Private Sub mnufiles_Click()

Load fileform
fileform.Height = 8010
fileform.Width = 4350
fileform.Show

End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)

Dim strSendCode As String
Dim vntArray As Variant
Dim strText As String
Dim nItems As Integer
Dim n As Integer

sckClient.GetData strSendCode, vbString

' split function will be used to parse items contained in a string,
' and delimitted by ||
' The Split function returns a variant array containing each parsed item
' as an element in the array

' use split function to parse it
vntArray = Split(strSendCode, "||")

' how many items were parsed?
nItems = UBound(vntArray)
  'Text1.Text = strsendcode
  
Select Case vntArray(0)
    Case "connect1"
        If vntArray(1) = "logonyes" Then
            MDIForm1.Show
            Unload Form1
        Else:
            MsgBox ("Login Incorrect!")
        End If
    Case "imcode2"
        MsgBox ("<" & vntArray(2) & ">" & vntArray(1))
    Case "chatcode2"
        Dim strHandle As String
        Dim strmessage As String
        
        strHandle = vntArray(1)
        strmessage = vntArray(2)
        
        If chatform.Tag <> "Closed" Then
            chatform.Text1.Text = chatform.Text1.Text & "<" & vntArray(1) & ">" & vntArray(2) & vbCrLf
        End If
    Case "filelistcode2"
        nItems = UBound(vntArray)

        ' display each parsed item
       
        fileform.ListView1.ListItems.Clear
        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "\/")
            
         
            Set itm = fileform.ListView1.ListItems.Add(, , vntarray2(1))
            itm.SubItems(1) = vntarray2(0)
            
        Next n
        Set itm = Nothing
        
    Case "mbmessagescode2"
        
        nItems = UBound(vntArray)

        ' display each parsed item
       
        ListView2.ListItems.Clear
        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "\/")
            
            Text1.Text = Text1.Text & vbCrLf & vntArray(n)
            Set itm = ListView2.ListItems.Add(, , vntarray2(0))
            itm.SubItems(1) = vntarray2(1)
            itm.SubItems(2) = vntarray2(2)
            itm.SubItems(3) = vntarray2(3)
            itm.SubItems(4) = vntarray2(4)
        Next n
       
        Set itm = Nothing
        
    Case "imcode2"
        
        
End Select

End Sub

