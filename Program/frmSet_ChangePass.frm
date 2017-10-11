VERSION 5.00
Begin VB.Form frmSet_ChangePass 
   BorderStyle     =   1  '단일 고정
   Caption         =   "비밀번호 변경"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSet_ChangePass.frx":0000
   ScaleHeight     =   4425
   ScaleWidth      =   4755
   StartUpPosition =   2  '화면 가운데
   Begin YAM.CandyButton cmdChange 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "변경하기"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   12632064
      ColorButtonDown =   8421376
      BorderBrightness=   0
      ColorBright     =   16777152
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox txtNowPass 
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox txtNewPass2 
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtNewPass1 
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  '사용 못함
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2200
      Width           =   3735
   End
   Begin YAM.CandyButton cmdBack 
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "뒤　로"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   12632064
      ColorButtonDown =   8421376
      BorderBrightness=   0
      ColorBright     =   16777152
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "현재 비밀번호"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "비밀번호 확인"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "새로운 비밀번호"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1845
      Width           =   3735
   End
End
Attribute VB_Name = "frmSet_ChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBack_Click()
Unload frmSet_ChangePass
End Sub

Private Sub cmdChange_Click()
Dim recAddress As New ADODB.Recordset
recAddress.Open "SELECT * FROM login where ID = '" & l_ID & "' and Pass = '" & txtNowPass & "'", adoConnect, adOpenStatic, adLockOptimistic

If recAddress.RecordCount = 0 Then
    MsgBox ("현재 비밀번호가 일치하지 않습니다.")
Else
    '일치 확인
    If txtNewPass1 = txtNewPass2 And Not txtNewPass1 = "" And Not txtNowPass = txtNewPass1 Then
        Dim query As String
        query = "UPDATE login SET Pass = '" & txtNewPass1 & "' WHERE Pass = '" & txtNowPass & "'"
        adoConnect.Execute query
        MsgBox ("비밀번호가 정상적으로 수정되었습니다.")
        Unload frmSet_ChangePass
    Else
        MsgBox ("새로운 비밀번호가 일치하지 않거나 중복된 비밀번호입니다.")
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSet_Main.Show
End Sub

