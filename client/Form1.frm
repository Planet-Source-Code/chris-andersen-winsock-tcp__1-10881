VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "iBBS Client Logon"
   ClientHeight    =   2340
   ClientLeft      =   5430
   ClientTop       =   6435
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   5760
   Begin VB.CommandButton Command1 
      Caption         =   "Logon"
      Height          =   435
      Left            =   2880
      TabIndex        =   6
      Top             =   1740
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1740
      TabIndex        =   5
      Top             =   600
      Width           =   3435
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1740
      TabIndex        =   3
      Top             =   1080
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   120
      Width           =   3435
   End
   Begin VB.Label Label3 
      Caption         =   "Handle"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

MDIForm1.sckClient.SendData "ibbslogin1||" & Text3.Text & "||" & Text2.Text
strHandle = Text3.Text

End Sub

Private Sub Form_Load()
Text1.Text = "127.0.0.1"

Load MDIForm1

'MDIForm1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub
