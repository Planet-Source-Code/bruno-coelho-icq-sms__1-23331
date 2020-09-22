VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Send SMS thru ICQ"
   ClientHeight    =   5040
   ClientLeft      =   4935
   ClientTop       =   2145
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   5040
   Begin VB.Frame status 
      Caption         =   " Status "
      Height          =   435
      Left            =   15
      TabIndex        =   10
      Top             =   4065
      Width           =   4950
   End
   Begin VB.TextBox number 
      Height          =   300
      Left            =   2625
      TabIndex        =   8
      Top             =   1065
      Width           =   2325
   End
   Begin VB.TextBox prefix 
      Height          =   300
      Left            =   1065
      TabIndex        =   6
      Top             =   1065
      Width           =   540
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4365
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Send"
      Height          =   345
      Left            =   3675
      TabIndex        =   5
      Top             =   3585
      Width           =   1260
   End
   Begin VB.TextBox msg 
      Height          =   1860
      Left            =   60
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1635
      Width           =   4875
   End
   Begin VB.TextBox pass 
      Height          =   300
      Left            =   1080
      TabIndex        =   1
      Top             =   660
      Width           =   1920
   End
   Begin VB.TextBox user 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   210
      Width           =   1890
   End
   Begin VB.Label email 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bruno@escripovoa.pt"
      Height          =   195
      Left            =   3390
      TabIndex        =   14
      Top             =   4725
      Width           =   1530
   End
   Begin VB.Label copyrigth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyrigth 2001 - Bruno Coelho"
      Height          =   195
      Left            =   75
      TabIndex        =   13
      Top             =   4710
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "You have :"
      Height          =   195
      Left            =   45
      TabIndex        =   12
      Top             =   3645
      Width           =   780
   End
   Begin VB.Label words 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   345
      Left            =   885
      TabIndex        =   11
      Top             =   3600
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number :"
      Height          =   195
      Index           =   3
      Left            =   1860
      TabIndex        =   9
      Top             =   1125
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prefix :"
      Height          =   195
      Index           =   2
      Left            =   495
      TabIndex        =   7
      Top             =   1125
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icq # :"
      Height          =   195
      Index           =   0
      Left            =   510
      TabIndex        =   2
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
status.Caption = "opening registry page and say your are online"
'opens the registry page and say your are online
Inet1.OpenURL "http://web.icq.com/karma/dologin/1,,,00.html?uService=1&uLogin=" + user.Text + "&uPassword=" + pass.Text
status.Caption = "Sending the message to the phone number"
'send the message to the phone number you want
Inet1.OpenURL "http://web.icq.com/sms/send_history/1,,,00.html?target=msghistory&prefix=+" + prefix.Text + "&carrier=aaa&tophone=" + number.Text + "&msg=" + msg.Text
MsgBox "Message sent with success !", vbInformation, "SMSIcq"
status.Caption = ""
End Sub



Private Sub email_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
email.ForeColor = QBColor(9)
email.FontUnderline = True
End Sub

Private Sub email_Click()
isTemp = "mailto:bruno@escripovoa.pt"
lRet = ShellExecute(hWnd, "open", isTemp, vbNull, vbNull, 1)

End Sub

Private Sub Form_Load()
words.Caption = 150
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
email.ForeColor = QBColor(0)
email.FontUnderline = False

End Sub

Private Sub Label3_Click()

End Sub

Private Sub msg_Change()
words.Caption = 150 - Len(msg)
End Sub
