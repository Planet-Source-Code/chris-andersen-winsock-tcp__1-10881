VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ActiveSkin.ocx"
Begin VB.Form frmMain 
   Caption         =   "iBBS Server"
   ClientHeight    =   8070
   ClientLeft      =   2385
   ClientTop       =   2025
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   Begin VB.CommandButton Command2 
      Caption         =   "&Close iBBS Server"
      Height          =   495
      Left            =   4860
      TabIndex        =   7
      Top             =   1500
      Width           =   1635
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   6360
      OleObjectBlob   =   "frmMain.frx":0442
      TabIndex        =   6
      Top             =   3660
      Width           =   480
   End
   Begin MSWinsockLib.Winsock ftpSock1 
      Index           =   0
      Left            =   5820
      Top             =   3660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4860
      Top             =   3660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start/Stop Server"
      Height          =   495
      Left            =   4860
      TabIndex        =   3
      ToolTipText     =   "Turn off/on iBBS Server"
      Top             =   120
      Width           =   1635
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Server &Options"
      Height          =   495
      Left            =   4860
      TabIndex        =   2
      Top             =   780
      Width           =   1635
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   5340
      Top             =   3660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame frStats 
      Caption         =   "Server Statistics"
      Height          =   3795
      Left            =   4440
      TabIndex        =   1
      Top             =   4200
      Width           =   2415
      Begin VB.Label lblusernm 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbluser 
         Caption         =   "Users Online:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.TextBox txtClientOutput 
      Height          =   7875
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4275
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSendCode As String
Dim lngSock As Long
Dim lngUsers As Long
Dim blnServerOn As Boolean
Dim blnServerPaused As Boolean
Dim blnSendDone As Boolean
Dim strMode As String
Dim strSubject As String
Dim strFrom As String
Dim strDate As String
Dim strReply As String
Dim strID As String
Dim db2 As Database
Dim rs2 As Recordset
Dim strMBString As String
Dim strFileStatus As String
Dim lngFtp As Long
Dim intBuffer As Integer
Dim lngBytesXfer As Long
Dim fFile As Long

Private Sub Command1_Click()

'turn server on and off
If blnServerOn = True Then
    sckServer(0).Close
    Dim X As Long
    For X = 1 To sckServer.UBound
        sckServer(X).Close
    Next X
    blnServerOn = False
    txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Server Status: Off"
Else
    sckServer(0).Listen
    blnServerOn = True
    txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Server Status: On"
End If


End Sub

Private Sub Command2_Click()
End


End Sub

Private Sub Form_Load()

'setup for icon minimize to system tray
With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = "iBBS Server" & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
End With


blnServerOn = True
blnServerPaused = False
lngSock = 0
lngFtp = 0
intBuffer = 2048
sckServer(lngSock).LocalPort = 1001
sckServer(lngSock).Listen
txtClientOutput.Text = "iBBS Server Version: " & App.Major & "." & App.Minor & " Revision " & App.Revision
txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Server Status: On"
lblusernm.Caption = "0"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Msg As Long
Msg = X

'for double clicking icon in system tray to restore or show menu
If Msg = WM_LBUTTONDBLCLK Then
    Call mnuShow_Click
ElseIf Msg = WM_RBUTTONDOWN Then
    PopupMenu mnuPopup
End If



End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    Me.Hide
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmMain = Nothing
Shell_NotifyIcon NIM_DELETE, IconData

End Sub

Private Sub sckServer_Close(Index As Integer)

Dim strUsers As String
Dim Z As Long

txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Disconnected:IP " & sckServer(Index).RemoteHostIP
'remove socket from collection

For Each colitem In colUsers
    colUsers.Remove CStr(Index)
Next colitem

'Having some problems with this right now
'strUsers = "userlistcode2||"
'i = 0
'create user list string to send to client
'For Each colitem In colUsers
'    i = i + 1
'    strUsers = strUsers & colUsers.Item(i) & "||"
'Next colitem

'send new user list to clients
'For Z = 1 To sckServer.UBound
'    If sckServer(Z).State <> 7 Then
'    Else
'        sckServer(Z).SendData strUsers
'        DoEvents
'    End If
'Next Z

sckServer(Index).Close
lngUsers = lngUsers - 1
If lngUsers < 0 Then lngUsers = 0
lblusernm.Caption = lngUsers

End Sub


Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim X As Long

'Check to see if any open socks are being used
'and use unused ones for new connections to save memory
For X = 1 To sckServer.UBound
    If sckServer(X).State <> 7 Then
        sckServer(X).Close
        txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Incoming Request:IP " & sckServer(Index).RemoteHostIP & " ID: " & requestID
        sckServer(X).Accept requestID
        txtClientOutput.Text = txtClientOutput.Text & vbCrLf & X
        'send ok connect to client in order to receive login info
        sckServer(X).SendData "connect1||ok"
        GoTo exitconnect
    End If
Next

'If all open socks are being used, create a new one.
lngSock = lngSock + 1
Load sckServer(lngSock)
txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Incoming Request:IP " & sckServer(Index).RemoteHostIP & " ID: " & requestID
txtClientOutput.Text = txtClientOutput.Text & vbCrLf & lngSock
sckServer(lngSock).Accept requestID
sckServer(X).SendData "connect1||ok"

exitconnect:

End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim lngInd As Long

sckServer(Index).GetData strSendCode, vbString
txtClientOutput.Text = txtClientOutput.Text & vbCrLf & strSendCode
'check send code

'process sendcodes and send back results to client(s)
lngInd = Index

CheckSendCode strSendCode, lngInd


End Sub

Private Sub CheckSendCode(strCode As String, lngIndex As Long)
'Parse the inputed sendcode using split function
Dim strHandle As String
Dim strPassword As String


Dim vntArray As Variant
Dim strText As String
Dim nItems As Integer
Dim n As Integer

Dim db As Database
Dim rs As Recordset

' split function will be used to parse items contained in a string,
' and delimitted by ||
' The Split function returns a variant array containing each parsed item
' as an element in the array

' use split function to parse it
vntArray = Split(strCode, "||")

' how many items were parsed?
nItems = UBound(vntArray)

'do a select case on the code type to determine clients request
'whether chat, im, mb, files, etc.. and take appropriate action
Select Case vntArray(0)
    Case "ibbslogin1"
        'login string
        strHandle = vntArray(1)
        strPassword = vntArray(2)
        txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Login attempt: " & strHandle & " " & strPassword
        'open user table in database and check user and password
        Set db = OpenDatabase(App.Path & "\users.mdb")
        Set rs = db.OpenRecordset("Select * from users where handle = '" & strHandle & "' and password ='" & strPassword & "'")
        If rs.RecordCount <> 0 Then
            'login passed
            blnLoginGood = True
            sckServer(lngIndex).SendData "connect1||logonyes"
            txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "User Connected:IP " & sckServer(Index).RemoteHostIP & " " & strID
            lngUsers = lngUsers + 1
            lblusernm = lngUsers
            'add user and socket to collection
            colUsers.Add strHandle, CStr(lngIndex)
            
            Dim W As Long
            Dim strUsersNew As String
            
            'Not working yet!
            'strUsersNew = "userlistcode2||"
            'send user list to clients
            'i = 0
            'For Each colitem In colUsers
            '    i = i + 1
            '    strUsersNew = strUsersNew & colUsers.Item(i) & "||"
            'Next colitem
            

            'send new user list to clients
            'For W = 1 To sckServer.UBound
            '    If sckServer(W).State <> 7 Then
            '    Else
            '        sckServer(W).SendData strUsersNew
            '        DoEvents
            '    End If
            'Next W
            
        Else
            'login failed
            blnLoginGood = False
            sckServer(lngIndex).SendData "connect1||logonno"
            txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "User Denied:IP " & sckServer(Index).RemoteHostIP & " " & strID
            'have client close connection
        End If
        
        
                
        Set rs = Nothing
        Set db = Nothing
        
    Case "chatcode1"
        'chatroom string
        Dim strSendChatText As String
        
        With colUsers(lngIndex)
        strSendChatText = "chatcode2||" & vntArray(2) & "||" & vntArray(1)
        txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Chat Text: " & strSendChatText
        For X = 1 To sckServer.UBound
            If sckServer(X).State <> 7 Then
            Else
                sckServer(X).SendData strSendChatText
                DoEvents
            End If
        Next X
        End With
    Case "filelistcode1"
        'file list request string
        Dim strSendList As String
        Dim strDir As String
        'do getfilelist function and return the list as a string
        
        With colUsers(CStr(lngIndex))
            strDir = vntArray(1)
            strSendList = GetFileList(strDir)
            'send list
            sckServer(lngIndex).SendData strSendList
        End With
        
    Case "getfilecode1"
        'Get file string
        Dim strFileName As String
        Dim BufferSize As Integer
        
        Dim lngFTPSock As Long
               
        txtClientOutput.Text = txtClientOutput.Text & vbCrLf & strCode
        
        With colUsers(CStr(lngIndex))
            strFileName = vntArray(1)
            lngBytesXfer = 0
            'Pass the users sock index and make an ftp winsock
            'connection to the user, this will cause problems with
            'firewalls, but I will deal with that later!
            lngFTPSock = MakeFTPConnection(lngIndex)
            fFile = FreeFile
        
            Open App.Path & "\files\" & strFileName For Binary Access Read As #fFile
        
            'loop through getfiledata getting each chunk of the file
            'and send it to the client until it is done
            Do Until FileLen(App.Path & "\files\" & strFileName) = lngBytesXfer
                    
                DoEvents
                strfiledata = GetFileData(FileLen(App.Path & "\files\" & strFileName), lngFTPSock, lngIndex)
                'txtClientOutput.Text = txtClientOutput.Text & vbCrLf & strfiledata
                ftpSock1(lngFTPSock).SendData strfiledata
            Loop
                   
            Close #fFile
        End With
        
    Case "imcode1"
        'instant message string
        Dim strIMMessage As String
        Dim strIMTo As String
        Dim strIMFrom As String
        Dim lngItem As Long
        Dim strIMSendtoClient As String
        
        strIMMessage = vntArray(1)
        strIMTo = vntArray(2)
        strIMFrom = vntArray(3)
        'check collection of users and find sock of selected user
        With colUsers(CStr(lngIndex))
        For Each colitem In colUsers
            i = i + 1
            DoEvents
            If colUsers(CStr(i)) = strIMTo Then
                lngItem = i
                'generate and send the string to the proper person
                strIMSendtoClient = "imcode2||" & strIMMessage & "||" & strIMFrom
                sckServer(lngItem).SendData strIMSendtoClient
            End If
        Next colitem
        End With
        
    Case "mbmessagescode1"
        'message board Get message list string
        Dim strSendMessageList As String
        
        With colUsers(lngIndex)
            'create list of messages
            strSendMessageList = GetMessageList()
            sckServer(lngIndex).SendData strSendMessageList
        End With
    Case "mbrepliescode1"
        'message board get replies code
        'incoming: sendcode1||messageid
        'outgoing: sendcode2||subject||handle||date||id
    Case "mbreadmessagecode1"
        'message board read message code
    Case "mbreadreplies"
        'message board read replies code
    Case "mailcode1"
        'email string
    Case Else
        'invalid sendcode. Do nothing
        txtClientOutput.Text = txtClientOutput.Text & vbCrLf & "Possbile Hack Attempt. IP: " & sckServer(lngIndex).RemoteHostIP & "  String: " & strCode
End Select

End Sub

Private Sub sckServer_SendComplete(Index As Integer)

'MsgBox "Done"
If strMode = "mb" Then
    blnSendDone = True
End If

End Sub

Private Sub Timer1_Timer()

'Keep checking states
For X = 0 To sckServer.UBound
    If sckServer(X).State = 8 Or sckServer(X).State = 9 Then
    sckServer(X).Close
    Exit Sub
    End If
Next

End Sub

Private Sub mnuExit_Click()

Unload Me
End

End Sub

Private Sub mnuShow_Click()

Me.WindowState = vbNormal
Shell_NotifyIcon NIM_DELETE, IconData
Me.Show

End Sub

Private Function GetFileList(strWorkingDir As String) As String
Dim hFile As Long
Dim fname As String
Dim WFD As WIN32_FIND_DATA
Dim dirList As String

    'Get the first file in the directory (it will usually return ".")
    hFile = FindFirstFile(App.Path & "\files\" & strWorkingDir & "*.*" + Chr$(0), WFD)
    
    'create string of files and their sizes to send
    dirList = "filelistcode2||"
    While FindNextFile(hFile, WFD)
        If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then
            dirList = dirList & WFD.nFileSizeLow & "\/" & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) & "||"
        End If
        DoEvents
    Wend

    GetFileList = dirList
    
End Function

Private Function GetMessageList() As String

Set db = OpenDatabase(App.Path & "\system\forum.mdb")
Set rs = db.OpenRecordset("Select * from tblQuestion")
strMBString = "mbmessagescode2||"
        
Do Until rs.EOF
    Set rs2 = db.OpenRecordset("Select * from tblResponse where questionid =" & rs!id)
    strSubject = rs!subject
    strFrom = rs!Handle
    strDate = rs!Date
    strID = rs!id
            
    If rs2.RecordCount <> 0 Then
        strReply = "Yes"
    Else:
        strReply = "No"
    End If
    Set rs2 = Nothing
           
    strMBString = strMBString & strSubject & "\/" & strFrom & "\/" & strDate & "\/" & strReply & "\/" & strID & "||"
    rs.MoveNext
Loop
Set rs = Nothing
Set db = Nothing

GetMessageList = strMBString


End Function

Private Function MakeFTPConnection(ftpIndex As Long) As Long
Dim Y As Long

'Check to see if any open socks are being used
'and use unused ones for new connections to save memory
For Y = 1 To ftpSock1.UBound
    If ftpSock1(Y).State <> 7 Then
        ftpSock1(Y).Close
        ftpSock1(Y).RemoteHost = sckServer(ftpIndex).RemoteHostIP
        ftpSock1(Y).RemotePort = "21"
        ftpSock1(Y).Connect
        
        MakeFTPConnection = Y
        GoTo exitftpconnect
    End If
Next

'If all open socks are being used, create a new one.
lngFtp = lngFtp + 1
Load ftpSock1(lngFtp)
ftpSock1(lngFtp).RemoteHost = sckServer(ftpIndex).RemoteHostIP
ftpSock1(lngFtp).RemotePort = "21"
ftpSock1(lngFtp).Connect
MakeFTPConnection = lngFtp

exitftpconnect:

End Function

Private Function GetFileData(ttlBytes As Long, lngsckFTP As Long, lngSock As Long)

Dim BlockSize As Integer
Dim DataToSend As String

BlockSize = intBuffer

'With colUsers(CStr(lngSock))
    'Determine the proper buffer size.
    If BlockSize > (ttlBytes - lngBytesXfer) Then
        BlockSize = (ttlBytes - lngBytesXfer)
    End If

    DataToSend = Space$(BlockSize) 'allocate space to store data.
    Get #fFile, , DataToSend 'get data chunk
    GetFileData = DataToSend
    
    lngBytesXfer = lngBytesXfer + BlockSize
    
'End With
    


End Function

