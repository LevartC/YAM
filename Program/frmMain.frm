VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  '���� ����
   Caption         =   "YAM (Your Account Manager)"
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
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   2  'ȭ�� ���
   Begin YAM.CandyButton cmdDate 
      Height          =   1695
      Left            =   4320
      TabIndex        =   22
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   27.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   12640511
      ColorButtonDown =   65535
      BorderBrightness=   0
      ColorBright     =   8438015
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdAccMove 
      Height          =   735
      Left            =   2040
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�ڻ��̵�"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdInput 
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�Է�"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdSearch 
      Height          =   735
      Left            =   3840
      TabIndex        =   16
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��ȸ"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdOption 
      Height          =   735
      Left            =   5520
      TabIndex        =   17
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdAccManage 
      Height          =   735
      Left            =   360
      TabIndex        =   18
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�ڻ����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdSummary 
      Height          =   735
      Left            =   2040
      TabIndex        =   19
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdCreator 
      Height          =   735
      Left            =   3840
      TabIndex        =   20
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdExit 
      Height          =   735
      Left            =   5520
      TabIndex        =   21
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8438015
      ColorButtonDown =   12648447
      BorderBrightness=   0
      ColorBright     =   33023
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdYearMonth 
      Height          =   615
      Left            =   4320
      TabIndex        =   23
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   12640511
      ColorButtonDown =   65535
      BorderBrightness=   0
      ColorBright     =   8438015
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Image imgDayNext 
      Height          =   600
      Left            =   5760
      Picture         =   "frmMain.frx":E57D
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image imgDayPrev 
      Height          =   600
      Left            =   720
      Picture         =   "frmMain.frx":F8C8
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image imgMonthNext 
      Height          =   600
      Left            =   5760
      Picture         =   "frmMain.frx":10C1E
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   720
   End
   Begin VB.Image imgMonthPrev 
      Height          =   600
      Left            =   720
      Picture         =   "frmMain.frx":11F69
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   720
   End
   Begin VB.Label lblImport_D 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "0 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   13
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblExpense_D 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "0 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label lblImport_M 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "0 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3840
      TabIndex        =   11
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label lblExpense_M 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "0 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   10
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Top             =   5160
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   5160
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�̴��� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   7
      Top             =   3720
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�̴��� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Top             =   3720
      Width           =   1515
   End
   Begin VB.Label Label9 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "�� �ڻ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "0 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblCash 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "0 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblBankbook 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "0 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAccManage_Click()
frmMain.Hide
frmManage.Show
End Sub

Private Sub cmdAccMove_Click()
frmMain.Hide
frmAccMove.Show
End Sub

Private Sub cmdCreator_Click()
MsgBox ("�غ����Դϴ�.")
End Sub

Private Sub cmdDate_Click()
g_Date = g_MainDate
frmCalendar.Show
End Sub

Private Sub cmdOption_Click()
frmMain.Hide
frmSet_Main.Show
End Sub

Private Sub cmdSearch_Click()
frmMain.Hide
frmSearch.Show
End Sub

Private Sub cmdSummary_Click()
MsgBox ("�غ����Դϴ�.")
End Sub

Private Sub Form_Activate()
g_MainDate = g_Date
Call updateDate
End Sub

Private Sub cmdExit_Click()
Unload frmMain
End Sub

Private Sub cmdInput_Click()
frmMain.Hide
Load frmInput
frmInput.Show
End Sub

Private Sub Command1_Click()
frmCalendar.Show
End Sub

Private Function updateDate()

cmdYearMonth.Caption = Year(g_MainDate) & "�� " & Month(g_MainDate) & "��"
Select Case Weekday(g_MainDate)
Case 1
cmdDate.Caption = Day(g_MainDate) & "��(��)"
Case 2
cmdDate.Caption = Day(g_MainDate) & "��(��)"
Case 3
cmdDate.Caption = Day(g_MainDate) & "��(��)"
Case 4
cmdDate.Caption = Day(g_MainDate) & "��(�)"
Case 5
cmdDate.Caption = Day(g_MainDate) & "��(��)"
Case 6
cmdDate.Caption = Day(g_MainDate) & "��(��)"
Case 7
cmdDate.Caption = Day(g_MainDate) & "��(��)"
End Select

Dim temp As String
Dim recSum As New ADODB.Recordset
'���� ����
recSum.Open "SELECT sum(quantity) as q_sum FROM uselog WHERE div_type = '����' and m_date like '" & Format(g_MainDate, "YYYYMM") & "%'", adoConnect, adOpenStatic, adLockOptimistic
temp = recSum.Fields("q_sum") & " "
If temp = " " Then
    lblExpense_M = "0 ��"
Else
    lblExpense_M = temp & " ��"
End If
recSum.Close
'�Ϻ� ����
recSum.Open "SELECT sum(quantity) as q_sum FROM uselog WHERE div_type = '����' and m_date = '" & Format(g_MainDate, "YYYYMMDD") & "'", adoConnect, adOpenStatic, adLockOptimistic
temp = recSum.Fields("q_sum") & " "
If temp = " " Then
    lblExpense_D = "0 ��"
Else
    lblExpense_D = temp & " ��"
End If
recSum.Close
'���� ����
recSum.Open "SELECT sum(quantity) as q_sum FROM uselog WHERE div_type = '����' and m_date like '" & Format(g_MainDate, "YYYYMM") & "%'", adoConnect, adOpenStatic, adLockOptimistic
temp = recSum.Fields("q_sum") & " "
If temp = " " Then
    lblImport_M = "0 ��"
Else
    lblImport_M = temp & " ��"
End If
recSum.Close
'�Ϻ� ����
recSum.Open "SELECT sum(quantity) as q_sum FROM uselog WHERE div_type = '����' and m_date = '" & Format(g_MainDate, "YYYYMMDD") & "'", adoConnect, adOpenStatic, adLockOptimistic
temp = recSum.Fields("q_sum") & " "
If temp = " " Then
    lblImport_D = "0 ��"
Else
    lblImport_D = temp & " ��"
End If
recSum.Close

'���� �ܾ� ��ȸ
Dim recAddress As New ADODB.Recordset
Dim l_Cash, l_Bankbook As Long
recAddress.Open "SELECT Cash FROM login", adoConnect, adOpenStatic, adLockOptimistic
l_Cash = recAddress.Fields("Cash")
lblCash = l_Cash & " ��"
recAddress.Close

'���� �ܾ� ��ȸ
recAddress.Open "SELECT SUM(quantity) as q_sum, count(*) as q_cnt FROM bankbook", adoConnect, adOpenStatic, adLockOptimistic
If recAddress.Fields("q_cnt") > 0 Then
l_Bankbook = recAddress.Fields("q_sum")
lblBankbook = l_Bankbook & " ��"
End If
lblTotal = (l_Cash + l_Bankbook) & " ��"

End Function

Private Sub Form_Load()
g_Date = Date

Call updateDate

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("������ �����Ͻðڽ��ϱ�?", vbYesNo, "����") = vbNo Then
Cancel = 1
End If
End Sub

Private Sub imgDayNext_Click()
g_MainDate = g_MainDate + 1
Call updateDate
End Sub

Private Sub imgDayPrev_Click()
g_MainDate = g_MainDate - 1
Call updateDate
End Sub

Private Sub imgMonthNext_Click()

Dim m_day As Integer
Select Case Month(g_MainDate)
Case 1, 3, 5, 7, 8, 10, 12
    m_day = 31
Case 4, 6, 9, 11
    m_day = 30
Case 2
    Dim t_year As Integer
    t_year = Year(g_MainDate)
    If (t_year Mod 4 = 0 And t_year Mod 100 <> 0) Or t_year Mod 400 = 0 Then
        m_day = 29
    Else
        m_day = 28
    End If
End Select

g_MainDate = g_MainDate + m_day
Call updateDate

End Sub

Private Sub imgMonthPrev_Click()
Dim m_day As Integer
Select Case Month(g_MainDate)
Case 1, 2, 4, 6, 8, 9, 11
    m_day = 31
Case 5, 7, 10, 12
    m_day = 30
Case 3
    Dim t_year As Integer
    t_year = Year(g_MainDate)
    If (t_year Mod 4 = 0 And t_year Mod 100 <> 0) Or t_year Mod 400 = 0 Then
        m_day = 29
    Else
        m_day = 28
    End If
End Select

g_MainDate = g_MainDate - m_day
Call updateDate

End Sub


