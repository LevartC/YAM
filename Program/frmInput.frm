VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInput 
   BorderStyle     =   1  '���� ����
   Caption         =   "����� �Է�"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "����"
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
   Picture         =   "frmInput.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   2  'ȭ�� ���
   Begin YAM.CandyButton cmdSave 
      Height          =   495
      Left            =   720
      TabIndex        =   18
      Top             =   8640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY�ٴ�L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "������"
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
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   12938
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY�ٴ�L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "������"
      TabPicture(0)   =   "frmInput.frx":2D26C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Label2"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "������"
      TabPicture(1)   =   "frmInput.frx":2D288
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label19"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "HY�ٴ�L"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   480
         TabIndex        =   34
         Top             =   2400
         Width           =   6255
         Begin VB.ComboBox cbxType_I 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   13
            Top             =   900
            Width           =   3255
         End
         Begin VB.TextBox txtComment_I 
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   1920
            TabIndex        =   17
            Top             =   3750
            Width           =   3975
         End
         Begin VB.TextBox txtQuantity_I 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   16
            Top             =   3075
            Width           =   3255
         End
         Begin VB.CommandButton cmdLoadCal 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            TabIndex        =   14
            Top             =   1605
            Width           =   720
         End
         Begin VB.ComboBox cbxCategory_I 
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   15
            Top             =   2340
            Width           =   3255
         End
         Begin VB.Line Line11 
            X1              =   120
            X2              =   6120
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line10 
            X1              =   120
            X2              =   6120
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "�ܡ�����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "������¥"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   1680
            Width           =   1140
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "ī�װ�"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   2400
            Width           =   1140
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "�ݡ�����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   41
            Top             =   3135
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   3990
            Width           =   1140
         End
         Begin VB.Label lblType_I 
            AutoSize        =   -1  'True
            Caption         =   "�롡����"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label lblDate2 
            Alignment       =   2  '��� ����
            AutoSize        =   -1  'True
            Caption         =   "��      ¥"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            TabIndex        =   37
            Top             =   1680
            Width           =   1110
         End
         Begin VB.Label lblImportBalance 
            Alignment       =   1  '������ ����
            AutoSize        =   -1  'True
            Caption         =   "0 ��"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5370
            TabIndex        =   36
            Top             =   285
            Width           =   525
         End
         Begin VB.Line Line8 
            X1              =   120
            X2              =   6120
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line7 
            X1              =   120
            X2              =   6120
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   6120
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   6120
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Label4 
            Alignment       =   1  '������ ����
            AutoSize        =   -1  'True
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   3120
            Width           =   285
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   480
         TabIndex        =   31
         Top             =   1200
         Width           =   6255
         Begin VB.OptionButton optType_I 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   3720
            TabIndex        =   12
            Top             =   440
            Width           =   1575
         End
         Begin VB.OptionButton optType_I 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1560
            TabIndex        =   11
            Top             =   440
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4695
         Left            =   -74520
         TabIndex        =   22
         Top             =   2400
         Width           =   6255
         Begin VB.ComboBox cbxType_E 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   6
            Top             =   900
            Width           =   3255
         End
         Begin VB.ComboBox cbxCategory_E 
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   8
            Top             =   2340
            Width           =   3255
         End
         Begin VB.CommandButton cmdLoadCal 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            TabIndex        =   7
            Top             =   1605
            Width           =   720
         End
         Begin VB.TextBox txtQuantity_E 
            Alignment       =   1  '������ ����
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2280
            TabIndex        =   9
            Top             =   3120
            Width           =   3255
         End
         Begin VB.TextBox txtComment_E 
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   1920
            TabIndex        =   10
            Top             =   3750
            Width           =   3975
         End
         Begin VB.Label lblType_E 
            AutoSize        =   -1  'True
            Caption         =   "�롡����"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   38
            Top             =   960
            Width           =   1140
         End
         Begin VB.Line Line9 
            X1              =   120
            X2              =   6120
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Label Label13 
            Alignment       =   1  '������ ����
            AutoSize        =   -1  'True
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5595
            TabIndex        =   33
            Top             =   3120
            Width           =   285
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   29
            Top             =   3990
            Width           =   1140
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   6120
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "�ݡ�����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   3135
            Width           =   1140
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   6120
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "ī�װ�"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   2400
            Width           =   1140
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   6120
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6120
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "������¥"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   1680
            Width           =   1140
         End
         Begin VB.Label lblExpenseBalance 
            Alignment       =   1  '������ ����
            AutoSize        =   -1  'True
            Caption         =   "0 ��"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5370
            TabIndex        =   25
            Top             =   285
            Width           =   525
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "�ܡ�����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Alignment       =   2  '��� ����
            AutoSize        =   -1  'True
            Caption         =   "��      ¥"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
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
            Top             =   1680
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   -74520
         TabIndex        =   21
         Top             =   1200
         Width           =   6255
         Begin VB.OptionButton optType_E 
            Caption         =   "üũī��"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4200
            TabIndex        =   5
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton optType_E 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2640
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optType_E 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "HY�ٴ�L"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   960
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Label Label19 
         Alignment       =   2  '��� ����
         BorderStyle     =   1  '���� ����
         Caption         =   "���� ������ �ۼ��մϴ�."
         BeginProperty Font 
            Name            =   "HY�ٴ�L"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label Label2 
         Alignment       =   2  '��� ����
         BorderStyle     =   1  '���� ����
         Caption         =   "���� ������ �ۼ��մϴ�."
         BeginProperty Font 
            Name            =   "HY�ٴ�L"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   30
         Top             =   600
         Width           =   6255
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   495
      Left            =   480
      TabIndex        =   20
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ī���ݰ���"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin YAM.CandyButton cmdBack 
      Height          =   495
      Left            =   4200
      TabIndex        =   19
      Top             =   8640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY�ٴ�L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�ڡ���"
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
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbxType_E_Click()
Dim i As Integer
For i = 0 To 2
    If optType_E(i).Value = True Then
        Call updateCBBalance(optType_E(i).Caption, cbxType_E.Text, lblExpenseBalance)
    End If
Next
End Sub

Private Sub cbxType_I_Click()
Dim i As Integer
For i = 0 To 1
    If optType_I(i).Value = True Then
        Call updateCBBalance(optType_I(i).Caption, cbxType_I.Text, lblImportBalance)
    End If
Next
End Sub

Private Sub cmdLoadCal_Click(Index As Integer)
frmCalendar.Show
End Sub

Private Sub cmdSave_Click()
Dim c_error As Boolean
Dim i As Integer
Dim s_type As String
Dim recCat_ID As New ADODB.Recordset
Dim cat_ID As Integer
Dim bank_ID As Integer
If SSTab1.Tab = 0 Then  '����
    '�Է� üũ
    For i = 0 To 2
        If optType_E(i) = True Then
            s_type = optType_E(i).Caption
        End If
    Next
    If s_type = "" Then
        MsgBox ("���� ������ �Է��ϼ���.")
        c_error = True
    End If
    If IsNumeric(txtQuantity_E.Text) = False Or txtQuantity_E = "" Then
        MsgBox ("�ݾ��� ��Ȯ�� �Է��ϼ���.")
        c_error = True
    End If

    If c_error Then
        c_error = False
    Else
        'ī�װ� �̸��� �̿��� ID�� ����
        recCat_ID.Open "SELECT ID FROM category_e WHERE m_name = '" & cbxCategory_E.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
        cat_ID = recCat_ID.Fields("ID")
        recCat_ID.Close
        Select Case s_type
        Case "����"
            '���� �ۼ�
            '���� �߰�
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, connect_id, e_memo) values('" & _
            Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_E & ", " & _
            cat_ID & ", '" & "����" & "', '" & s_type & "', '" & l_ID & "', '" & txtComment_E & "')"
            adoConnect.Execute query
            'ī�װ� ������Ʈ
            query = "update category_e set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '����ݾ� ����
            query = "update login set cash = cash - " & txtQuantity_E & " where ID = '" & l_ID & "'"
            adoConnect.Execute query
            
            MsgBox "���������� ����Ǿ����ϴ�."
            Unload frmInput
        Case "����"
            '���� ID �ۼ�
            recCat_ID.Open "SELECT ID FROM bankbook WHERE m_name = '" & cbxType_E.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            bank_ID = recCat_ID.Fields("ID")
            '���� �ۼ�
            '���� �߰�
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, bankbook_id, connect_id, e_memo) values('" & _
            Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_E & ", " & _
            cat_ID & ", '" & "����" & "', '" & s_type & "', " & bank_ID & ", '" & l_ID & "', '" & txtComment_E & "')"
            adoConnect.Execute query
            'ī�װ� ������Ʈ
            query = "update category_e set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '����ݾ� ����
            query = "update bankbook set quantity = quantity - " & txtQuantity_E & " where ID = " & bank_ID
            adoConnect.Execute query
            
            MsgBox "���������� ����Ǿ����ϴ�."
            Unload frmInput
        Case "üũī��"
            '���� ID �ۼ�
            recCat_ID.Open "SELECT bankbook_id FROM checkcard WHERE m_name = '" & cbxType_E.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            bank_ID = recCat_ID.Fields("bankbook_id")
            '���� �ۼ�
            '���� �߰�
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, bankbook_id, connect_id, e_memo) values('" & _
            Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_E & ", " & _
            cat_ID & ", '" & "����" & "', '" & s_type & "', " & bank_ID & ", '" & l_ID & "', '" & txtComment_E & "')"
            adoConnect.Execute query
            'ī�װ� ������Ʈ
            query = "update category_e set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '����ݾ� ����
            query = "update bankbook set quantity = quantity - " & txtQuantity_E & " where ID = " & bank_ID
            adoConnect.Execute query
            
            MsgBox "���������� ����Ǿ����ϴ�."
            Unload frmInput
        End Select
    End If
Else    '����
    '�Է� üũ
    For i = 0 To 1
        If optType_I(i) = True Then
            s_type = optType_I(i).Caption
        End If
    Next
    If s_type = "" Then
        MsgBox ("���� ������ �Է��ϼ���.")
        c_error = True
    End If
    If IsNumeric(txtQuantity_I) = False Or txtQuantity_I = "" Then
        MsgBox ("�ݾ��� ��Ȯ�� �Է��ϼ���.")
        c_error = True
    End If
    
    If c_error Then
        c_error = False
    Else
        'ī�װ� �̸��� �̿��� ID�� ����
        recCat_ID.Open "SELECT ID FROM category_i WHERE m_name = '" & cbxCategory_I.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
        cat_ID = recCat_ID.Fields("ID")
        recCat_ID.Close
        Select Case s_type
        Case "����"
            'ī�װ� �̸��� �̿��� ID�� ����
            '���� �ۼ�
            '���� �߰�
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, connect_id, e_memo) values('" & _
            Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_I & ", " & _
            cat_ID & ", '" & "����" & "', '" & s_type & "', '" & l_ID & "', '" & txtComment_I & "')"
            adoConnect.Execute query
            'ī�װ� ����
            query = "update category_i set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '����ݾ� ����
            query = "update login set cash = cash + " & txtQuantity_I & " where ID = '" & l_ID & "'"
            adoConnect.Execute query
            
            MsgBox "���������� ����Ǿ����ϴ�."
            Unload frmInput
        Case "����"
            '���� ID �ۼ�
            recCat_ID.Open "SELECT ID FROM bankbook WHERE m_name = '" & cbxType_E.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
            bank_ID = recCat_ID.Fields("ID")
            '���� �ۼ�
            '���� �߰�
            query = "insert into uselog(m_date, quantity, category_id, div_type, c_type, bankbook_id, connect_id, e_memo) values('" & _
            Format(g_Date, "YYYYMMDD") & "', " & txtQuantity_I & ", " & _
            cat_ID & ", '" & "����" & "', '" & s_type & "', " & bank_ID & ", '" & l_ID & "', '" & txtComment_I & "')"
            adoConnect.Execute query
            'ī�װ� ������Ʈ
            query = "update category_i set fou = fou + 1 where ID = " & cat_ID
            adoConnect.Execute query
            '����ݾ� ����
            query = "update bankbook set quantity = quantity + " & txtQuantity_I & " where m_name = '" & cbxType_I.Text & "'"
            adoConnect.Execute query
            
            MsgBox "���������� ����Ǿ����ϴ�."
            Unload frmInput
        End Select
    End If
End If

End Sub

Private Sub cmdBack_Click()
Unload frmInput
End Sub

Private Sub Form_Activate()
lblDate1 = g_Date
lblDate2 = g_Date
End Sub

Private Sub Form_Load()
'Date �ֱ�
g_Date = Date
lblDate1 = g_Date
lblDate2 = g_Date

Dim recCat_ID As New ADODB.Recordset
'����ī�װ� ����Ʈ �ε�
recCat_ID.Open "SELECT * FROM category_e ORDER BY fou desc", adoConnect, adOpenStatic, adLockOptimistic
If recCat_ID.RecordCount > 0 Then
    recCat_ID.MoveFirst
cbxCategory_E.Text = recCat_ID.Fields("m_name")
End If
Do While recCat_ID.EOF = False
cbxCategory_E.AddItem recCat_ID.Fields("m_name")
recCat_ID.MoveNext
Loop
recCat_ID.Close

'����ī�װ� ����Ʈ �ε�
recCat_ID.Open "SELECT * FROM category_i ORDER BY fou desc", adoConnect, adOpenStatic, adLockOptimistic
If recCat_ID.RecordCount > 0 Then
    recCat_ID.MoveFirst
cbxCategory_I.Text = recCat_ID.Fields("m_name")
End If
Do While recCat_ID.EOF = False
cbxCategory_I.AddItem recCat_ID.Fields("m_name")
recCat_ID.MoveNext
Loop
recCat_ID.Close

'�ܾ� ����
SSTab1.Tab = 1
Call updateBalance(optType_E(0).Caption)
SSTab1.Tab = 0
Call updateBalance(optType_I(0).Caption)

End Sub



Private Sub Form_Unload(Cancel As Integer)
g_Date = g_MainDate
frmMain.Show
End Sub

Private Sub optType_I_Click(Index As Integer)
Call updateBalance(optType_I(Index).Caption)
End Sub

Private Sub optType_E_Click(Index As Integer)
Call updateBalance(optType_E(Index).Caption)
End Sub

Private Function updateBalance(str As String)
'�ܾ� ����
Dim recTemp As New ADODB.Recordset
If SSTab1.Tab = 0 Then  '���� ��
Select Case str
Case "����"
    lblType_E.Enabled = False
    cbxType_E.Enabled = False
    recTemp.Open "SELECT Cash FROM login WHERE ID = '" & l_ID & "'", adoConnect, adOpenStatic, adLockOptimistic
    lblExpenseBalance = recTemp.Fields("Cash") & " ��"
    recTemp.Close
Case "����"
    recTemp.Open "SELECT * FROM bankbook ORDER BY ID", adoConnect, adOpenStatic, adLockOptimistic
    If recTemp.RecordCount > 0 Then
        cbxType_E.Clear
        lblType_E.Enabled = True
        lblType_E.Caption = "�롡����"
        cbxType_E.Enabled = True
        recTemp.MoveFirst
        lblExpenseBalance = recTemp.Fields("quantity") & " ��"
        cbxType_E.Text = recTemp.Fields("m_name")
        Do While recTemp.EOF = False
            cbxType_E.AddItem recTemp.Fields("m_name")
            recTemp.MoveNext
        Loop
        recTemp.Close
    Else
        MsgBox ("������ �����ϴ�.")
        optType_E(0).Value = True
    End If
Case "üũī��"
    recTemp.Open "SELECT * FROM checkcard ORDER BY ID", adoConnect, adOpenStatic, adLockOptimistic
    If recTemp.RecordCount > 0 Then
        cbxType_E.Clear
        lblType_E.Enabled = True
        lblType_E.Caption = "üũī��"
        cbxType_E.Enabled = True
        recTemp.MoveFirst
        Dim i_temp As Integer
        i_temp = recTemp("bankbook_id")
        cbxType_E.Text = recTemp.Fields("m_name")
        Do While recTemp.EOF = False
            cbxType_E.AddItem recTemp.Fields("m_name")
            recTemp.MoveNext
        Loop
        recTemp.Close
        recTemp.Open "SELECT quantity FROM bankbook WHERE ID = " & i_temp, adoConnect, adOpenStatic, adLockOptimistic
        lblExpenseBalance = recTemp.Fields("quantity") & " ��"
        recTemp.Close
    Else
        MsgBox ("üũī�尡 �����ϴ�.")
        optType_E(0).Value = True
    End If
End Select
Else    '���� ��
Select Case str
Case "����"
    lblType_I.Enabled = False
    cbxType_I.Enabled = False
    recTemp.Open "SELECT Cash FROM login WHERE ID = '" & l_ID & "'", adoConnect, adOpenStatic, adLockOptimistic
    lblImportBalance = recTemp.Fields("Cash") & " ��"
    recTemp.Close
Case "����"
    recTemp.Open "SELECT * FROM bankbook ORDER BY ID", adoConnect, adOpenStatic, adLockOptimistic
    If recTemp.RecordCount > 0 Then
        cbxType_I.Clear
        lblType_I.Enabled = True
        lblType_I.Caption = "�롡����"
        cbxType_I.Enabled = True
        recTemp.MoveFirst
        lblImportBalance = recTemp.Fields("quantity") & " ��"
        cbxType_I.Text = recTemp.Fields("m_name")
        Do While recTemp.EOF = False
            cbxType_I.AddItem recTemp.Fields("m_name")
            recTemp.MoveNext
        Loop
        recTemp.Close
    Else
        MsgBox ("������ �����ϴ�.")
        optType_I(0).Value = True
    End If
End Select
End If

End Function


