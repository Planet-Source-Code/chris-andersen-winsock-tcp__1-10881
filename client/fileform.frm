VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fileform 
   Caption         =   "File Section"
   ClientHeight    =   5085
   ClientLeft      =   7470
   ClientTop       =   3360
   ClientWidth     =   4230
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   4230
   Begin MSWinsockLib.Winsock ftpclient 
      Left            =   3780
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7335
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   12938
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "fileform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ftpClient_ConnectionRequest(ByVal requestID As Long)

ftpclient.Close
'ftp1.Listen
ftpclient.Accept requestID

End Sub

Private Sub ftpClient_DataArrival(ByVal bytesTotal As Long)

Dim data As String

ftpclient.GetData data

'MsgBox bytesTotal

Put #fFile, , data

If strFileLen = Loc(fFile) Then
    Close #fFile
    ftpclient.Close
    ftpclient.Listen
End If

End Sub

Private Sub Form_Load()

ftpclient.LocalPort = "21"
ftpclient.Listen

MDIForm1.sckClient.SendData "filelistcode1||"

End Sub

Private Sub ListView1_DblClick()

MDIForm1.sckClient.SendData "getfilecode1||" & fileform.ListView1.SelectedItem

fFile = FreeFile

strFileLen = fileform.ListView1.SelectedItem.SubItems(1)

Open App.Path & "\dl\" & fileform.ListView1.SelectedItem For Binary Access Write As #fFile


End Sub
