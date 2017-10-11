VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Login"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   12
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   2  '화면 가운데
   Begin YAM.CandyButton cmdLogin 
      Height          =   615
      Left            =   3350
      TabIndex        =   2
      Top             =   8500
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "로그인"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox txtID 
      Height          =   480
      IMEMode         =   3  '사용 못함
      Left            =   3120
      MaxLength       =   64
      TabIndex        =   0
      Text            =   "123"
      Top             =   7080
      Width           =   3735
   End
   Begin VB.TextBox txtPass 
      Height          =   480
      IMEMode         =   3  '사용 못함
      Left            =   3120
      MaxLength       =   64
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "123"
      Top             =   7680
      Width           =   3735
   End
   Begin YAM.CandyButton cmdExit 
      Height          =   615
      Left            =   5200
      TabIndex        =   3
      Top             =   8500
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "종료"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()

Dim recAddress As New ADODB.Recordset
recAddress.Open "SELECT * FROM login where ID = '" & txtID & "' and Pass = '" & txtPass & "'", adoConnect, adOpenStatic, adLockOptimistic

If recAddress.RecordCount = 0 Then
    MsgBox ("비밀번호가 일치하지 않습니다.")
Else
    recAddress.MoveFirst
    MsgBox "환영합니다!"
    l_ID = recAddress.Fields("ID")
    b_Login = True
    Unload frmLogin
    Load frmMain
    frmMain.Show
End If


End Sub

Private Sub cmdExit_Click()
Unload frmLogin
End Sub

Private Sub Form_Load()
Set adoConnect = New ADODB.Connection

With adoConnect
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path + "\YAM.mdb"
    .Open
End With

g_MainDate = Date
g_Date = Date
b_Login = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not b_Login Then
    If MsgBox("정말로 종료하시겠습니까?", vbYesNo, "종료") = vbNo Then
        Cancel = 1
    End If
End If
End Sub
