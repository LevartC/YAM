VERSION 5.00
Begin VB.Form frmSet_Main 
   BorderStyle     =   1  '단일 고정
   Caption         =   "설정"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetting.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   7350
   StartUpPosition =   2  '화면 가운데
   Begin YAM.CandyButton Command2 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY바다L"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "카테고리 추가 및 수정"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   8438015
      ColorButtonUp   =   33023
      ColorButtonDown =   16576
      BorderBrightness=   0
      ColorBright     =   8438015
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdChangePass 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY바다L"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "비밀번호 변경"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   8438015
      ColorButtonUp   =   33023
      ColorButtonDown =   16576
      BorderBrightness=   0
      ColorBright     =   8438015
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdBack 
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY바다L"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "뒤로"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   8438015
      ColorButtonUp   =   33023
      ColorButtonDown =   16576
      BorderBrightness=   0
      ColorBright     =   8438015
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "설　정"
      BeginProperty Font 
         Name            =   "한컴 바겐세일 M"
         Size            =   48
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   3
      Top             =   -120
      Width           =   2295
   End
End
Attribute VB_Name = "frmSet_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Unload frmSet_Main
End Sub

Private Sub Command2_Click()
frmSet_Main.Hide
frmSet_Category.Show
End Sub

Private Sub cmdChangePass_Click()
frmSet_Main.Hide
frmSet_ChangePass.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

