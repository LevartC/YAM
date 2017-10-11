VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAccMove 
   BorderStyle     =   1  '단일 고정
   Caption         =   "자산이동"
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
   Picture         =   "frmAccMove.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   2  '화면 가운데
   Begin YAM.CandyButton cmdSave 
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   8760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "저　장"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16744703
      ColorButtonDown =   16711935
      BorderBrightness=   0
      ColorBright     =   16744703
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin TabDlg.SSTab sstAccMove 
      Height          =   7455
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "현금출금"
      TabPicture(0)   =   "frmAccMove.frx":12A2A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Label2"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "현금입금"
      TabPicture(1)   =   "frmAccMove.frx":12A46
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame4 
         Height          =   5535
         Left            =   240
         TabIndex        =   36
         Top             =   1780
         Width           =   6495
         Begin VB.TextBox txtQuantity_D 
            Alignment       =   1  '오른쪽 맞춤
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   13
            Top             =   3675
            Width           =   3255
         End
         Begin VB.CommandButton cmdLoadCal 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   5280
            TabIndex        =   11
            Top             =   2205
            Width           =   720
         End
         Begin VB.ComboBox cbxCategory_D 
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
            Left            =   2280
            TabIndex        =   12
            Top             =   2940
            Width           =   3255
         End
         Begin VB.ComboBox cbxType_D 
            Enabled         =   0   'False
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
            Left            =   2280
            TabIndex        =   10
            Top             =   1500
            Width           =   3255
         End
         Begin VB.TextBox txtComment_D 
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   1920
            TabIndex        =   14
            Top             =   4440
            Width           =   3975
         End
         Begin VB.Label lblDate2 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            Caption         =   "날      짜"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            TabIndex        =   47
            Top             =   2280
            Width           =   1110
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "현금잔액"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   46
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label lblCashBalance 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            Caption         =   "0 원"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   5370
            TabIndex        =   45
            Top             =   285
            Width           =   525
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "날　　짜"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   44
            Top             =   2280
            Width           =   1140
         End
         Begin VB.Line Line14 
            X1              =   240
            X2              =   6240
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line10 
            X1              =   240
            X2              =   6240
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "카테고리"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   43
            Top             =   3000
            Width           =   1140
         End
         Begin VB.Line Line9 
            X1              =   240
            X2              =   6240
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "금　　액"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   42
            Top             =   3735
            Width           =   1140
         End
         Begin VB.Line Line8 
            X1              =   240
            X2              =   6240
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label15 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5595
            TabIndex        =   41
            Top             =   3720
            Width           =   285
         End
         Begin VB.Line Line7 
            X1              =   240
            X2              =   6240
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label lblType_D 
            AutoSize        =   -1  'True
            Caption         =   "통　　장"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   40
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Label lblBC_D 
            AutoSize        =   -1  'True
            Caption         =   "통장잔액"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   39
            Top             =   900
            Width           =   1140
         End
         Begin VB.Label lblBCBalance_D 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            Caption         =   "0 원"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5370
            TabIndex        =   38
            Top             =   915
            Width           =   525
         End
         Begin VB.Line Line6 
            X1              =   240
            X2              =   6240
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "내　　용"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   37
            Top             =   4680
            Width           =   1140
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5535
         Left            =   -74760
         TabIndex        =   22
         Top             =   1780
         Width           =   6495
         Begin VB.TextBox txtCommision_W 
            Alignment       =   1  '오른쪽 맞춤
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   6
            Top             =   4320
            Width           =   3255
         End
         Begin VB.TextBox txtComment_W 
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1920
            TabIndex        =   7
            Top             =   4920
            Width           =   3975
         End
         Begin VB.ComboBox cbxType_W 
            Enabled         =   0   'False
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
            Left            =   2280
            TabIndex        =   2
            Top             =   1500
            Width           =   3255
         End
         Begin VB.ComboBox cbxCategory_W 
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
            Left            =   2280
            TabIndex        =   4
            Top             =   2940
            Width           =   3255
         End
         Begin VB.CommandButton cmdLoadCal 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   5280
            TabIndex        =   3
            Top             =   2205
            Width           =   720
         End
         Begin VB.TextBox txtQuantity_W 
            Alignment       =   1  '오른쪽 맞춤
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   5
            Top             =   3675
            Width           =   3255
         End
         Begin VB.Label Label25 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5595
            TabIndex        =   35
            Top             =   4365
            Width           =   285
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "수 수 료"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   34
            Top             =   4380
            Width           =   1035
         End
         Begin VB.Line Line13 
            X1              =   240
            X2              =   6240
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "내　　용"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   33
            Top             =   5040
            Width           =   1140
         End
         Begin VB.Line Line12 
            X1              =   240
            X2              =   6240
            Y1              =   4200
            Y2              =   4200
         End
         Begin VB.Label lblBCBalance_W 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            Caption         =   "0 원"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5370
            TabIndex        =   32
            Top             =   915
            Width           =   525
         End
         Begin VB.Label lblBC_W 
            AutoSize        =   -1  'True
            Caption         =   "통장잔액"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   31
            Top             =   900
            Width           =   1140
         End
         Begin VB.Label lblType_W 
            AutoSize        =   -1  'True
            Caption         =   "통　　장"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   30
            Top             =   1560
            Width           =   1140
         End
         Begin VB.Line Line11 
            X1              =   240
            X2              =   6240
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label Label9 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            Caption         =   "원"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5595
            TabIndex        =   29
            Top             =   3720
            Width           =   285
         End
         Begin VB.Line Line5 
            X1              =   240
            X2              =   6240
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "금　　액"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   28
            Top             =   3735
            Width           =   1140
         End
         Begin VB.Line Line4 
            X1              =   240
            X2              =   6240
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "카테고리"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   27
            Top             =   3000
            Width           =   1140
         End
         Begin VB.Line Line3 
            X1              =   240
            X2              =   6240
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line2 
            X1              =   240
            X2              =   6240
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "날　　짜"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   26
            Top             =   2280
            Width           =   1140
         End
         Begin VB.Label lblCashBalance 
            Alignment       =   1  '오른쪽 맞춤
            AutoSize        =   -1  'True
            Caption         =   "0 원"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   5370
            TabIndex        =   25
            Top             =   285
            Width           =   525
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "현금잔액"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   24
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label lblDate1 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            Caption         =   "날      짜"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            TabIndex        =   23
            Top             =   2280
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74760
         TabIndex        =   20
         Top             =   940
         Width           =   6495
         Begin VB.OptionButton optType_W 
            Caption         =   "통장"
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
            Index           =   0
            Left            =   1680
            TabIndex        =   0
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optType_W 
            Caption         =   "체크카드"
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
            Index           =   1
            Left            =   3600
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   240
         TabIndex        =   19
         Top             =   940
         Width           =   6495
         Begin VB.OptionButton optType_D 
            Caption         =   "체크카드"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   9
            Top             =   300
            Width           =   1575
         End
         Begin VB.OptionButton optType_D 
            Caption         =   "통장"
            BeginProperty Font 
               Name            =   "HY바다L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   8
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Label Label14 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   1  '단일 고정
         Caption         =   "체크/통장으로 현금을 입금했을 때 사용합니다."
         BeginProperty Font 
            Name            =   "HY바다L"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   21
         Top             =   465
         Width           =   6405
      End
      Begin VB.Label Label2 
         Alignment       =   2  '가운데 맞춤
         BorderStyle     =   1  '단일 고정
         Caption         =   "체크/통장에서 현금을 출금했을 때 사용합니다."
         BeginProperty Font 
            Name            =   "HY바다L"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74700
         TabIndex        =   18
         Top             =   465
         Width           =   6405
      End
   End
   Begin YAM.CandyButton cmdBack 
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      Top             =   8760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      ColorButtonHover=   16761087
      ColorButtonUp   =   16744703
      ColorButtonDown =   16711935
      BorderBrightness=   0
      ColorBright     =   16744703
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
End
Attribute VB_Name = "frmAccMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbxType_D_Click()
If optType_D(0).Value Then
    Call updateCBBalance(optType_D(0).Caption, cbxType_D.Text, lblBCBalance_D)
Else
    Call updateCBBalance(optType_D(1).Caption, cbxType_D.Text, lblBCBalance_D)
End If
End Sub

Private Sub cbxType_W_Click()
If optType_W(0).Value Then
    Call updateCBBalance(optType_W(0).Caption, cbxType_W.Text, lblBCBalance_W)
Else
    Call updateCBBalance(optType_W(1).Caption, cbxType_W.Text, lblBCBalance_W)
End If
End Sub

Private Sub cmdBack_Click()
Unload frmAccMove
End Sub

Private Sub cmdLoadCal_Click(Index As Integer)
frmCalendar.Show
End Sub

Private Sub cmdSave_Click()
Dim c_error As Boolean
Dim i As Integer
Dim s_type As String
Dim cat_ID As Integer
Dim bank_ID As Integer
Dim recTemp As New ADODB.Recordset
Dim query As String
Select Case sstAccMove.Tab
Case 0  '통장출금
    '입력 체크
    For i = 0 To 1
        If optType_W(i) = True Then
            s_type = optType_W(i).Caption
        End If
    Next
    If IsNumeric(txtQuantity_W.Text) = False Or txtQuantity_W = "" Then
        MsgBox ("금액을 정확히 입력하세요.")
        c_error = True
    End If
    If IsNumeric(txtCommision_W.Text) = False Or txtCommision_W = "" Then
        MsgBox ("수수료가 정확히 입력되지 않아 0으로 처리합니다.")
        c_error = True
    End If
    
    If c_error Then
        c_error = False
    Else
        '카테고리 이름을 이용해 ID를 추출
        recTemp.Open "SELECT ID FROM category_m WHERE m_name = '" & cbxCategory_W.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
        cat_ID = recTemp.Fields("ID")
        recTemp.Close
        Select Case s_type
        Case "통장"
            '통장 이름을 이용해 ID를 추출
            recTemp.Open "SELECT ID FROM bankbook WHERE m_name = '" & cbxType_W.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            bank_ID = recTemp.Fields("ID")
            recTemp.Close
            '쿼리 작성
            '내역 추가
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, bankbook_id, connect_id, commision, e_memo) values('" & _
                    Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_W & ", " & _
                    cat_ID & ", '출금', '" & s_type & "', " & bank_ID & ", '" & _
                    l_ID & "', " & txtCommision_W & ", '" & txtComment_W & "')"
            adoConnect.Execute query
            '카테고리 업데이트
            query = "update category_m set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '금액 갱신
            query = "update login set cash = cash + " & txtQuantity_W & " where ID = '" & l_ID & "'"
            adoConnect.Execute query
            query = "update bankbook set quantity = quantity - " & (CLng(txtQuantity_W) + CLng(txtCommision_W)) & " where m_name = '" & cbxType_W.Text & "'"
            adoConnect.Execute query
            
            MsgBox "성공적으로 저장되었습니다."
            Unload frmAccMove
        Case "체크카드"
            '카드 이름을 이용해 통장 ID를 추출
            recTemp.Open "SELECT bankbook_id FROM checkcard WHERE m_name = '" & cbxType_W.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            bank_ID = recTemp.Fields("bankbook_id")
            recTemp.Close
            '쿼리 작성
            '내역 추가
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, bankbook_id, connect_id, commision, e_memo) values('" & _
                    Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_W & ", " & _
                    cat_ID & ", '출금', '" & s_type & "', " & bank_ID & ", '" & _
                    l_ID & "', " & txtCommision_W & ", '" & txtComment_W & "')"
            adoConnect.Execute query
            '카테고리 업데이트
            query = "update category_m set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '금액 갱신
            query = "update login set cash = cash + " & txtQuantity_W & " where ID = '" & l_ID & "'"
            adoConnect.Execute query
            query = "update bankbook set quantity = quantity - " & (CLng(txtQuantity_W) + CLng(txtCommision_W)) & " where ID = " & bank_ID
            adoConnect.Execute query
            
            MsgBox "성공적으로 저장되었습니다."
            Unload frmAccMove
        End Select
    End If
Case 1  '통장입금
    '입력 체크
    For i = 0 To 1
        If optType_D(i) = True Then
            s_type = optType_D(i).Caption
        End If
    Next
    If IsNumeric(txtQuantity_D.Text) = False Or txtQuantity_D = "" Then
        MsgBox ("금액을 정확히 입력하세요.")
        c_error = True
    End If
    
    If c_error Then
        c_error = False
    Else
        '카테고리 이름을 이용해 ID를 추출
        recTemp.Open "SELECT ID FROM category_m WHERE m_name = '" & cbxCategory_D.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
        cat_ID = recTemp.Fields("ID")
        recTemp.Close
        Select Case s_type
        Case "통장"
            '통장 이름을 이용해 ID를 추출
            recTemp.Open "SELECT ID FROM bankbook WHERE m_name = '" & cbxType_D.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            bank_ID = recTemp.Fields("ID")
            recTemp.Close
            '쿼리 작성
            '내역 추가
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, bankbook_id, connect_id, e_memo) values('" & _
                    Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_D & ", " & _
                    cat_ID & ", '입금', '" & s_type & "', " & bank_ID & ", '" & _
                    l_ID & "', '" & txtComment_D & "')"
            adoConnect.Execute query
            '카테고리 업데이트
            query = "update category_m set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '금액 갱신
            query = "update login set cash = cash - " & txtQuantity_D & " where ID = '" & l_ID & "'"
            adoConnect.Execute query
            query = "update bankbook set quantity = quantity + " & txtQuantity_D & " where m_name = '" & cbxType_D.Text & "'"
            adoConnect.Execute query
            
            MsgBox "성공적으로 저장되었습니다."
            Unload frmAccMove
        Case "체크카드"
            '카드 이름을 이용해 통장 ID를 추출
            recTemp.Open "SELECT bankbook_id FROM checkcard WHERE m_name = '" & cbxType_D.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            bank_ID = recTemp.Fields("bankbook_id")
            recTemp.Close
            '쿼리 작성
            '내역 추가
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, bankbook_id, connect_id, e_memo) values('" & _
                    Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_D & ", " & _
                    cat_ID & ", '입금', '" & s_type & "', " & bank_ID & ", '" & _
                    l_ID & "', '" & txtComment_D & "')"
            adoConnect.Execute query
            '카테고리 업데이트
            query = "update category_m set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '금액 갱신
            query = "update login set cash = cash + " & txtQuantity_D & " where ID = '" & l_ID & "'"
            adoConnect.Execute query
            query = "update bankbook set quantity = quantity - " & txtQuantity_D & " where ID = " & bank_ID
            adoConnect.Execute query
            
            MsgBox "성공적으로 저장되었습니다."
            Unload frmAccMove
        End Select
    End If
Case 2
End Select
End Sub

Private Sub Form_Activate()
lblDate1 = g_Date
End Sub

Private Sub Form_Load()
'Date 넣기
g_Date = Date
lblDate1 = g_Date
lblDate2 = g_Date

Dim recCat_ID As New ADODB.Recordset
'이동카테고리 리스트 로드
recCat_ID.Open "SELECT * FROM category_m ORDER BY fou desc", adoConnect, adOpenStatic, adLockOptimistic
If recCat_ID.RecordCount > 0 Then
    recCat_ID.MoveFirst
cbxCategory_W.Text = recCat_ID.Fields("m_name")
cbxCategory_D.Text = recCat_ID.Fields("m_name")
End If
Do While recCat_ID.EOF = False
cbxCategory_W.AddItem recCat_ID.Fields("m_name")
cbxCategory_D.AddItem recCat_ID.Fields("m_name")
recCat_ID.MoveNext
Loop
recCat_ID.Close

Dim i As Integer
Dim recTemp As New ADODB.Recordset
'현금 잔액 갱신
recTemp.Open "SELECT Cash FROM login WHERE ID = '" & l_ID & "'", adoConnect, adOpenStatic, adLockOptimistic
lblCashBalance(0).Caption = recTemp.Fields("Cash") & " 원"
lblCashBalance(1).Caption = recTemp.Fields("Cash") & " 원"
recTemp.Close
sstAccMove.Tab = 1
Call updateBalance(optType_D(0).Caption)
sstAccMove.Tab = 0
Call updateBalance(optType_W(0).Caption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
g_Date = g_MainDate
frmMain.Show
End Sub


Private Sub optType_W_Click(Index As Integer)
Call updateBalance(optType_W(Index).Caption)
End Sub

Private Sub optType_D_Click(Index As Integer)
Call updateBalance(optType_D(Index).Caption)
End Sub

Private Function updateBalance(strType As String)

Dim recTemp As New ADODB.Recordset
Dim i_temp As Integer
Select Case sstAccMove.Tab
Case 0
    Select Case strType
    Case "통장"
        '통장 잔액 & 리스트 갱신
        recTemp.Open "SELECT m_name FROM bankbook ORDER BY ID", adoConnect, adOpenStatic, adLockOptimistic
        If recTemp.RecordCount > 0 Then
            cbxType_W.Clear
            cbxType_W.Enabled = True
            lblBC_W = "통장잔액"
            lblType_W = "통　　장"
            recTemp.MoveFirst
            cbxType_W.Text = recTemp.Fields("m_name")
            Do While recTemp.EOF = False
                cbxType_W.AddItem recTemp.Fields("m_name")
                recTemp.MoveNext
            Loop
            recTemp.Close
            recTemp.Open "SELECT quantity FROM bankbook WHERE m_name = '" & cbxType_W.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            lblBCBalance_W = recTemp.Fields("quantity") & " 원"
            recTemp.Close
        Else
            MsgBox ("통장이 없습니다.")
            cbxType_W.Enabled = False
        End If
    Case "체크카드"
        '체크 잔액 & 리스트 갱신
        recTemp.Open "SELECT * FROM checkcard ORDER BY ID", adoConnect, adOpenStatic, adLockOptimistic
        If recTemp.RecordCount > 0 Then
            cbxType_W.Clear
            lblBC_W = "체크잔액"
            lblType_W = "체크카드"
            cbxType_W.Enabled = True
            recTemp.MoveFirst
            i_temp = recTemp("bankbook_id")
            cbxType_W.Text = recTemp.Fields("m_name")
            Do While recTemp.EOF = False
                cbxType_W.AddItem recTemp.Fields("m_name")
                recTemp.MoveNext
            Loop
            recTemp.Close
            recTemp.Open "SELECT quantity FROM bankbook WHERE ID = " & i_temp, adoConnect, adOpenStatic, adLockOptimistic
            lblBCBalance_W = recTemp.Fields("quantity") & " 원"
            recTemp.Close
        Else
            MsgBox ("체크카드가 없습니다.")
            cbxType_W.Enabled = False
        End If
    End Select
Case 1
    Select Case strType
    Case "통장"
        '통장 잔액 & 리스트 갱신
        recTemp.Open "SELECT m_name FROM bankbook ORDER BY ID", adoConnect, adOpenStatic, adLockOptimistic
        If recTemp.RecordCount > 0 Then
            cbxType_D.Clear
            cbxType_D.Enabled = True
            lblBC_D = "통장잔액"
            lblType_D = "통　　장"
            recTemp.MoveFirst
            cbxType_D.Text = recTemp.Fields("m_name")
            Do While recTemp.EOF = False
                cbxType_D.AddItem recTemp.Fields("m_name")
                recTemp.MoveNext
            Loop
            recTemp.Close
            recTemp.Open "SELECT quantity FROM bankbook WHERE m_name = '" & cbxType_D.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            lblBCBalance_D = recTemp.Fields("quantity") & " 원"
            recTemp.Close
        Else
            MsgBox ("통장이 없습니다.")
            cbxType_D.Enabled = False
        End If
    Case "체크카드"
        '체크 잔액 & 리스트 갱신
        recTemp.Open "SELECT * FROM checkcard ORDER BY ID", adoConnect, adOpenStatic, adLockOptimistic
        If recTemp.RecordCount > 0 Then
            cbxType_D.Clear
            lblBC_D = "체크잔액"
            lblType_D = "체크카드"
            cbxType_D.Enabled = True
            recTemp.MoveFirst
            i_temp = recTemp("bankbook_id")
            cbxType_D.Text = recTemp.Fields("m_name")
            Do While recTemp.EOF = False
                cbxType_D.AddItem recTemp.Fields("m_name")
                recTemp.MoveNext
            Loop
            recTemp.Close
            recTemp.Open "SELECT quantity FROM bankbook WHERE ID = " & i_temp, adoConnect, adOpenStatic, adLockOptimistic
            lblBCBalance_D = recTemp.Fields("quantity") & " 원"
            recTemp.Close
        Else
            MsgBox ("체크카드가 없습니다.")
            cbxType_D.Enabled = False
        End If
    End Select
End Select

End Function


