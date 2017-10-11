VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBudget 
   BorderStyle     =   1  '단일 고정
   Caption         =   "예산"
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBudget.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows 기본값
   Begin YAM.CandyButton cmdBack 
      Height          =   615
      Left            =   4200
      TabIndex        =   37
      Top             =   8280
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "뒤로"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "수　입"
      TabPicture(0)   =   "frmBudget.frx":3C8C0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "지　출"
      TabPicture(1)   =   "frmBudget.frx":3C8DC
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   -74880
         TabIndex        =   19
         Top             =   340
         Width           =   6975
         Begin VB.Frame Frame2 
            Height          =   855
            Left            =   360
            TabIndex        =   33
            Top             =   360
            Width           =   6255
            Begin VB.CommandButton Command6 
               Caption         =   ">"
               Height          =   495
               Left            =   5160
               TabIndex        =   35
               Top             =   240
               Width           =   855
            End
            Begin VB.CommandButton Command5 
               Caption         =   "<"
               Height          =   495
               Left            =   240
               TabIndex        =   34
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label2 
               Alignment       =   2  '가운데 맞춤
               Caption         =   "2012년 12월"
               Height          =   255
               Left            =   1560
               TabIndex        =   36
               Top             =   360
               Width           =   3135
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3375
            Left            =   360
            TabIndex        =   23
            Top             =   1200
            Width           =   6255
            Begin VB.TextBox Text1 
               Height          =   495
               Left            =   1680
               TabIndex        =   26
               Top             =   360
               Width           =   3855
            End
            Begin VB.TextBox Text2 
               Height          =   495
               Left            =   1680
               TabIndex        =   25
               Top             =   1440
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Height          =   495
               Left            =   1680
               TabIndex        =   24
               Top             =   2520
               Width           =   3855
            End
            Begin VB.Label Label5 
               Caption         =   "원"
               Height          =   255
               Index           =   2
               Left            =   5640
               TabIndex        =   32
               Top             =   2640
               Width           =   255
            End
            Begin VB.Label Label5 
               Caption         =   "원"
               Height          =   255
               Index           =   1
               Left            =   5640
               TabIndex        =   31
               Top             =   1560
               Width           =   255
            End
            Begin VB.Label Label5 
               Caption         =   "원"
               Height          =   255
               Index           =   0
               Left            =   5640
               TabIndex        =   30
               Top             =   480
               Width           =   255
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FF0000&
               X1              =   240
               X2              =   6000
               Y1              =   2280
               Y2              =   2280
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FF0000&
               Index           =   0
               X1              =   240
               X2              =   6000
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label Label3 
               Caption         =   "남은 예산"
               Height          =   495
               Index           =   2
               Left            =   240
               TabIndex        =   29
               Top             =   2640
               Width           =   1575
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "총 수입"
               Height          =   240
               Index           =   1
               Left            =   240
               TabIndex        =   28
               Top             =   1560
               Width           =   795
            End
            Begin VB.Label Label3 
               Caption         =   "총 예산"
               Height          =   495
               Index           =   0
               Left            =   240
               TabIndex        =   27
               Top             =   480
               Width           =   1575
            End
         End
         Begin VB.Frame Frame4 
            Height          =   855
            Left            =   360
            TabIndex        =   20
            Top             =   4800
            Width           =   6255
            Begin VB.CommandButton Command7 
               Caption         =   ">"
               Height          =   495
               Left            =   5040
               TabIndex        =   21
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label4 
               Caption         =   "예산목록 보기"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   360
               Width           =   1575
            End
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5895
         Left            =   120
         TabIndex        =   1
         Top             =   340
         Width           =   6975
         Begin VB.Frame Frame6 
            Height          =   855
            Left            =   360
            TabIndex        =   16
            Top             =   4800
            Width           =   6255
            Begin VB.CommandButton Command3 
               Caption         =   ">"
               Height          =   495
               Left            =   5040
               TabIndex        =   17
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "예산목록 보기"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame Frame7 
            Height          =   3375
            Left            =   360
            TabIndex        =   6
            Top             =   1200
            Width           =   6255
            Begin VB.TextBox Text4 
               Height          =   495
               Left            =   1680
               TabIndex        =   9
               Top             =   2520
               Width           =   3855
            End
            Begin VB.TextBox Text5 
               Height          =   495
               Left            =   1680
               TabIndex        =   8
               Top             =   1440
               Width           =   3855
            End
            Begin VB.TextBox Text6 
               Height          =   495
               Left            =   1680
               TabIndex        =   7
               Top             =   360
               Width           =   3855
            End
            Begin VB.Label Label3 
               Caption         =   "총 예산"
               Height          =   495
               Index           =   3
               Left            =   240
               TabIndex        =   15
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "총 지출"
               Height          =   495
               Index           =   4
               Left            =   240
               TabIndex        =   14
               Top             =   1560
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "남은 예산"
               Height          =   495
               Index           =   5
               Left            =   240
               TabIndex        =   13
               Top             =   2640
               Width           =   1575
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FF0000&
               Index           =   1
               X1              =   240
               X2              =   6000
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FF0000&
               X1              =   240
               X2              =   6000
               Y1              =   2280
               Y2              =   2280
            End
            Begin VB.Label Label5 
               Caption         =   "원"
               Height          =   255
               Index           =   3
               Left            =   5640
               TabIndex        =   12
               Top             =   480
               Width           =   255
            End
            Begin VB.Label Label5 
               Caption         =   "원"
               Height          =   255
               Index           =   4
               Left            =   5640
               TabIndex        =   11
               Top             =   1560
               Width           =   255
            End
            Begin VB.Label Label5 
               Caption         =   "원"
               Height          =   255
               Index           =   5
               Left            =   5640
               TabIndex        =   10
               Top             =   2640
               Width           =   255
            End
         End
         Begin VB.Frame Frame8 
            Height          =   855
            Left            =   360
            TabIndex        =   2
            Top             =   360
            Width           =   6255
            Begin VB.CommandButton Command4 
               Caption         =   "<"
               Height          =   495
               Left            =   240
               TabIndex        =   4
               Top             =   240
               Width           =   855
            End
            Begin VB.CommandButton Command9 
               Caption         =   ">"
               Height          =   495
               Left            =   5160
               TabIndex        =   3
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label7 
               Alignment       =   2  '가운데 맞춤
               Caption         =   "2012년 12월"
               Height          =   255
               Left            =   1560
               TabIndex        =   5
               Top             =   360
               Width           =   3135
            End
         End
      End
   End
   Begin YAM.CandyButton cmdAddBudget 
      Height          =   615
      Left            =   1320
      TabIndex        =   38
      Top             =   8280
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "예산 추가"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
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
Attribute VB_Name = "frmBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub
