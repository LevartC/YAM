VERSION 5.00
Begin VB.Form frmSearch_Update 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "���� ����"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "HY�ٴ�M"
      Size            =   14.25
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4860
   StartUpPosition =   2  'ȭ�� ���
   Begin YAM.CandyButton cmdUpdate 
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "�����ϱ�"
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
   Begin VB.TextBox txtComment 
      Height          =   765
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.ComboBox cbxCategory 
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin YAM.CandyButton cmdLoadCal 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ">"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   8421376
      ColorButtonUp   =   16777088
      ColorButtonDown =   12632064
      BorderBrightness=   0
      ColorBright     =   16776960
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.TextBox txtQuantity 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin YAM.CandyButton cmdBack 
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "�ڷ�"
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "HY�ٴ�L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Top             =   1440
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "HY�ٴ�L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "�ݡ�����"
      BeginProperty Font 
         Name            =   "HY�ٴ�L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "ī�װ�"
      BeginProperty Font 
         Name            =   "HY�ٴ�L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label lblDate 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "������¥"
      BeginProperty Font 
         Name            =   "HY�ٴ�L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "������¥"
      BeginProperty Font 
         Name            =   "HY�ٴ�L"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmSearch_Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
Unload frmSearch_Update
End Sub

Private Sub cmdLoadCal_Click()
frmCalendar.Show
End Sub

Private Sub cmdUpdate_Click()
Dim c_error As Boolean
'�Է� üũ
If IsNumeric(txtQuantity.Text) = False Or txtQuantity = "" Then
    MsgBox ("�ݾ��� ��Ȯ�� �Է��ϼ���.")
    c_error = True
End If

If c_error Then
    c_error = False
Else
    Dim log_ID As Long
    Dim recCat_ID As New ADODB.Recordset
    Dim cat_ID As String
    Dim query As String
    Dim strType As String
    
    log_ID = frmSearch.lvwList.SelectedItem.SubItems(5)
    
    Select Case frmSearch.lvwList.SelectedItem.SubItems(6)
    Case "����"
        strType = "category_e"
    Case "����"
        strType = "category_i"
    Case "���"
        strType = "category_m"
    Case "�Ա�"
        strType = "category_m"
    End Select
    'ī�װ� �̸��� �̿��� ID�� ����
    recCat_ID.Open "SELECT ID FROM " & strType & " WHERE m_name = '" & cbxCategory.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
    cat_ID = recCat_ID.Fields("ID")
    recCat_ID.Close
    '���� �ۼ�
    query = "update uselog set m_date = '" & Format(g_Date, "YYYYMMDD") & "', quantity = " & txtQuantity & ", category_id = " & cat_ID & _
            ", e_memo = '" & txtComment & "' where ID = " & log_ID
    adoConnect.Execute query
    MsgBox ("���������� �����Ǿ����ϴ�.")
    Unload frmSearch_Update
End If

End Sub

Private Sub Form_Activate()
lblDate = g_Date
End Sub

Private Sub Form_Load()
lblDate = g_Date

Dim recCat As New ADODB.Recordset
Dim strType As String

'�ʱ�ȭ
cbxCategory.Text = frmSearch.lvwList.SelectedItem.SubItems(1)
txtComment = frmSearch.lvwList.SelectedItem.SubItems(3)
txtQuantity = Mid(frmSearch.lvwList.SelectedItem.SubItems(2), 2)
Select Case frmSearch.lvwList.SelectedItem.SubItems(6)
Case "����"
    strType = "category_e"
Case "����"
    strType = "category_i"
Case "���"
    strType = "category_m"
Case "�Ա�"
    strType = "category_m"
End Select

recCat.Open "SELECT * FROM " & strType & " ORDER BY fou desc", adoConnect, adOpenStatic, adLockOptimistic
If recCat.RecordCount > 0 Then
    recCat.MoveFirst
End If
Do While Not recCat.EOF
cbxCategory.AddItem recCat.Fields("m_name")
recCat.MoveNext
Loop
recCat.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSearch.Show
b_Switch = True
End Sub

