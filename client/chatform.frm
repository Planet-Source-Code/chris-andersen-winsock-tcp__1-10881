VERSION 5.00
Begin VB.Form chatform 
   Caption         =   "Chat Room"
   ClientHeight    =   7305
   ClientLeft      =   1935
   ClientTop       =   3225
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Send "
      Height          =   495
      Left            =   5220
      TabIndex        =   3
      Top             =   7380
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   7035
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   180
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   7035
      Left            =   4860
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   180
      Width           =   1635
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   7380
      Width           =   5055
   End
End
Attribute VB_Name = "chatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

MDIForm1.sckClient.SendData "chatcode1||" & Text3.Text & "||" & strHandle

End Sub

Private Sub Form_Paint()

chatform.Tag = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

chatform.Tag = "Closed"

'Set chatform = Nothing

End Sub

