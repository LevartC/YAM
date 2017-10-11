VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmManage 
   BorderStyle     =   1  '단일 고정
   Caption         =   "자산관리"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "HY바다L"
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
   Picture         =   "frmManage.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   2  '화면 가운데
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12938
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "현금"
      TabPicture(0)   =   "frmManage.frx":331DA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdChange"
      Tab(0).Control(1)=   "txtCash"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Label2"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "통장"
      TabPicture(1)   =   "frmManage.frx":331F6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAddBank"
      Tab(1).Control(1)=   "txtANumber"
      Tab(1).Control(2)=   "txtQuantity"
      Tab(1).Control(3)=   "txtName_B"
      Tab(1).Control(4)=   "lvwBankbook"
      Tab(1).Control(5)=   "cmdUpdateBank"
      Tab(1).Control(6)=   "cmdDeleteBank"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label6"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "체크카드"
      TabPicture(2)   =   "frmManage.frx":33212
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdDeleteCheck"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdUpdateCheck"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdAddCheck"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lvwCheckcard"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtName_C"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cbxBankbook"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin YAM.CandyButton cmdAddBank 
         Height          =   615
         Left            =   -74640
         TabIndex        =   6
         Top             =   6480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HY바다L"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "추　가"
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
      Begin VB.TextBox txtANumber 
         Height          =   525
         Left            =   -72480
         TabIndex        =   4
         Top             =   5040
         Width           =   3615
      End
      Begin VB.TextBox txtQuantity 
         Height          =   525
         Left            =   -72480
         TabIndex        =   5
         Top             =   5760
         Width           =   3615
      End
      Begin VB.TextBox txtName_B 
         Height          =   525
         Left            =   -72480
         TabIndex        =   3
         Top             =   4320
         Width           =   3615
      End
      Begin VB.ComboBox cbxBankbook 
         Height          =   405
         Left            =   2520
         TabIndex        =   10
         Top             =   5880
         Width           =   3615
      End
      Begin VB.TextBox txtName_C 
         Height          =   525
         Left            =   2520
         TabIndex        =   9
         Top             =   5160
         Width           =   3615
      End
      Begin ComctlLib.ListView lvwCheckcard 
         Height          =   4455
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   7858
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin ComctlLib.ListView lvwBankbook 
         Height          =   3615
         Left            =   -74640
         TabIndex        =   20
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6376
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin YAM.CandyButton cmdChange 
         Height          =   615
         Left            =   -74400
         TabIndex        =   2
         Top             =   2280
         Width           =   5775
         _ExtentX        =   10186
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
         Caption         =   "수　정　하　기"
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
      Begin VB.TextBox txtCash 
         Alignment       =   1  '오른쪽 맞춤
         Height          =   495
         Left            =   -72720
         TabIndex        =   1
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   -74640
         TabIndex        =   15
         Top             =   480
         Width           =   6255
         Begin VB.Label lblCash 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "0 원"
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "현　금"
            Height          =   285
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
      End
      Begin YAM.CandyButton cmdUpdateBank 
         Height          =   615
         Left            =   -72400
         TabIndex        =   7
         Top             =   6480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HY바다L"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수　정"
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
      Begin YAM.CandyButton cmdDeleteBank 
         Height          =   615
         Left            =   -70200
         TabIndex        =   8
         Top             =   6480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HY바다L"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "삭　제"
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
      Begin YAM.CandyButton cmdAddCheck 
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   6480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HY바다L"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "추　가"
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
      Begin YAM.CandyButton cmdUpdateCheck 
         Height          =   615
         Left            =   2600
         TabIndex        =   12
         Top             =   6480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HY바다L"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "수　정"
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
      Begin YAM.CandyButton cmdDeleteCheck 
         Height          =   615
         Left            =   4800
         TabIndex        =   13
         Top             =   6480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HY바다L"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "삭　제"
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "계좌 번호"
         Height          =   285
         Left            =   -74280
         TabIndex        =   26
         Top             =   5160
         Width           =   1230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "통장 잔액"
         Height          =   285
         Left            =   -74280
         TabIndex        =   25
         Top             =   5865
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "통장 이름"
         Height          =   285
         Left            =   -74280
         TabIndex        =   24
         Top             =   4485
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "연결 통장"
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   5880
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "카드 이름"
         Height          =   285
         Left            =   720
         TabIndex        =   22
         Top             =   5280
         Width           =   1230
      End
      Begin VB.Label Label3 
         Alignment       =   1  '오른쪽 맞춤
         AutoSize        =   -1  'True
         Caption         =   "원"
         Height          =   285
         Left            =   -68880
         TabIndex        =   19
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "금액 입력"
         Height          =   285
         Left            =   -74400
         TabIndex        =   18
         Top             =   1680
         Width           =   1230
      End
   End
   Begin YAM.CandyButton cmdBack 
      Height          =   615
      Left            =   4800
      TabIndex        =   14
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
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
Attribute VB_Name = "frmManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddBank_Click()

'입력 체크
Dim i As Integer
Dim c_error As Boolean
Dim query As String

If txtName_B = "" Then
    MsgBox "통장 이름을 입력하세요."
    c_error = True
End If
For i = 1 To lvwBankbook.ListItems.Count
    If txtName_B = lvwBankbook.ListItems(i) Then
        MsgBox ("중복된 이름입니다. 다른 이름을 입력하세요.")
        c_error = True
        txtName_B.SetFocus
    End If
Next
If Not IsNumeric(txtQuantity.Text) Or txtQuantity = "" Then
    MsgBox "통장 잔액을 정확하게 입력하세요."
    c_error = True
End If

If Not c_error Then
    query = "INSERT INTO bankbook(a_number, m_name, quantity) VALUES('" & _
    txtANumber & "', '" & txtName_B & "', " & txtQuantity & ")"
    adoConnect.Execute query
    MsgBox ("정상적으로 추가되었습니다.")
    Call updateBankbook
End If

End Sub

Private Sub cmdAddCheck_Click()

'입력 체크
Dim i As Integer
Dim c_error As Boolean
Dim query As String

If txtName_C = "" Then
    MsgBox "카드 이름을 입력하세요."
    c_error = True
End If
For i = 1 To lvwCheckcard.ListItems.Count
    If txtName_C = lvwCheckcard.ListItems(i) Then
        MsgBox ("중복된 이름입니다. 다른 이름을 입력하세요.")
        c_error = True
        txtName_C.SetFocus
    End If
Next

If Not c_error Then
    Dim recSelect As New ADODB.Recordset
    recSelect.Open "SELECT ID FROM bankbook where m_name = '" & cbxBankbook.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
    recSelect.MoveFirst
    query = "INSERT INTO Checkcard(m_name, bankbook_id) VALUES('" & _
    txtName_C & "', " & recSelect.Fields("ID") & ")"
    adoConnect.Execute query
    MsgBox ("정상적으로 추가되었습니다.")
    recSelect.Close
    Call updateCheckcard
End If

End Sub

Private Sub cmdBack_Click()
Unload frmManage
End Sub

Private Sub cmdChange_Click()

'입력 체크
Dim c_error As Boolean
Dim query As String
If Not IsNumeric(txtCash.Text) Or txtCash = "" Then
    MsgBox "금액을 정확하게 입력하세요."
    c_error = True
End If

If Not c_error Then
    query = "update login set cash = " & txtCash & " where ID = '" & l_ID & "'"
    adoConnect.Execute query
    MsgBox ("정상적으로 수정되었습니다.")
    Call updateCash
End If

End Sub

Private Sub cmdDeleteBank_Click()
If MsgBox("정말로 삭제하시겠습니까?", vbYesNo, "삭제") = vbYes Then
    '쿼리 작성
    Dim query As String
    query = "delete from bankbook where ID = " & lvwBankbook.SelectedItem.SubItems(3)
    adoConnect.Execute query
    MsgBox ("정상적으로 삭제되었습니다.")
    Call updateBankbook
End If
End Sub

Private Sub cmdDeleteCheck_Click()
If MsgBox("정말로 삭제하시겠습니까?", vbYesNo, "삭제") = vbYes Then
    '쿼리 작성
    Dim query As String
    query = "delete from checkcard where ID = " & lvwCheckcard.SelectedItem.SubItems(2)
    adoConnect.Execute query
    MsgBox ("정상적으로 삭제되었습니다.")
    Call updateCheckcard
End If
End Sub

Private Sub cmdUpdateBank_Click()

If MsgBox("수정하시겠습니까?", vbYesNo, "수정") = vbYes Then
    '입력 중복 체크
    Dim c_error As Boolean
    Dim query As String
    Dim i As Integer
    
    If txtName_B = "" Then
        MsgBox "통장 이름을 입력하세요."
        c_error = True
    End If
    For i = 1 To lvwBankbook.ListItems.Count
        If txtName_B = lvwBankbook.ListItems(i) And Not lvwBankbook.SelectedItem = lvwBankbook.ListItems(i) Then
            MsgBox ("중복된 이름입니다. 다른 이름을 입력하세요.")
            c_error = True
            txtName_B.SetFocus
        End If
    Next
    If Not IsNumeric(txtQuantity.Text) Or txtQuantity = "" Then
        MsgBox "통장 잔액을 정확하게 입력하세요."
        c_error = True
    End If
    
    If Not c_error Then
        query = "update bankbook set a_number = '" & txtANumber & "', m_name = '" & txtName_B & "', quantity = " & txtQuantity & " where ID = " & lvwBankbook.SelectedItem.SubItems(3)
        adoConnect.Execute query
        MsgBox ("정상적으로 수정되었습니다.")
        Call updateBankbook
    End If
End If

End Sub

Private Sub cmdUpdateCheck_Click()
If MsgBox("수정하시겠습니까?", vbYesNo, "수정") = vbYes Then
    '입력 중복 체크
    Dim c_error As Boolean
    Dim query As String
    Dim i As Integer
    
    If txtName_C = "" Then
        MsgBox "카드 이름을 입력하세요."
        c_error = True
    End If
    For i = 1 To lvwCheckcard.ListItems.Count
        If txtName_C = lvwCheckcard.ListItems(i) And Not lvwCheckcard.SelectedItem = lvwCheckcard.ListItems(i) Then
            MsgBox ("중복된 이름입니다. 다른 이름을 입력하세요.")
            c_error = True
            txtName_C.SetFocus
        End If
    Next
    
    If Not c_error Then
        Dim recSelect As New ADODB.Recordset
        recSelect.Open "SELECT ID FROM bankbook where m_name = '" & cbxBankbook.Text & "'", adoConnect, adOpenStatic, adLockOptimistic
        recSelect.MoveFirst
        query = "update checkcard set m_name = '" & txtName_C & "', bankbook_id = " & recSelect.Fields("ID") & " where ID = " & lvwCheckcard.SelectedItem.SubItems(2)
        adoConnect.Execute query
        MsgBox ("정상적으로 수정되었습니다.")
        recSelect.Close
        Call updateCheckcard
    End If
End If
End Sub

Private Sub Form_Load()

Call bankbookRefresh
Call updateCheckcard
Call updateBankbook
Call updateCash

End Sub

Private Function bankbookRefresh()
Dim recTemp As New ADODB.Recordset
recTemp.Open "SELECT m_name FROM bankbook", adoConnect, adOpenStatic, adLockOptimistic
If recTemp.RecordCount <> 0 Then
    cbxBankbook.Clear
    Do While recTemp.EOF = False
        cbxBankbook.AddItem recTemp.Fields("m_name")
        recTemp.MoveNext
    Loop
Else
    MsgBox "연결된 통장이 없으므로 카드를 만드실 수 없습니다."
    lvwCheckcard.Enabled = False
End If
recTemp.Close
End Function

Private Function updateCash()

Dim recSearch As New ADODB.Recordset
recSearch.Open "SELECT Cash FROM login WHERE ID = '" & l_ID & "'", adoConnect, adOpenStatic, adLockOptimistic

If recSearch.RecordCount = 0 Then
    MsgBox "현금 잔액 조회 실패."
    Unload frmManage
Else
    lblCash = recSearch.Fields("Cash") & " 원"
End If

End Function

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Function updateCheckcard()

With lvwCheckcard
.View = lvwReport
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "카드 이름", 1500
.ColumnHeaders.Add , , "연결통장 이름", 2000
.ColumnHeaders.Add , , "ID", 0
End With

Dim lstItem As ListItem
Dim recSelect As New ADODB.Recordset
Dim recBank As New ADODB.Recordset

recSelect.Open "SELECT * FROM checkcard", adoConnect, adOpenStatic, adLockOptimistic
If recSelect.RecordCount <> 0 Then
    lvwCheckcard.ListItems.Clear
    Do While Not recSelect.EOF
        Set lstItem = lvwCheckcard.ListItems.Add(, , recSelect.Fields("m_name"))
            recBank.Open "SELECT m_name FROM bankbook where ID = " & recSelect.Fields("bankbook_id"), adoConnect, adOpenStatic, adLockOptimistic
            If recBank.RecordCount <> 0 Then
                lstItem.SubItems(1) = recBank.Fields("m_name")
            End If
            lstItem.SubItems(2) = recSelect.Fields("ID")
            recBank.Close
        recSelect.MoveNext
    Loop
End If

txtName_C = ""
cbxBankbook.Text = ""

End Function

Private Function updateBankbook()

With lvwBankbook
.View = lvwReport
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "통장 이름", 1500
.ColumnHeaders.Add , , "계좌 번호", 2500
.ColumnHeaders.Add , , "통장 잔액", 1200
.ColumnHeaders.Add , , "ID", 0
End With

Dim lstItem As ListItem
Dim recSelect As New ADODB.Recordset

recSelect.Open "SELECT * FROM bankbook", adoConnect, adOpenStatic, adLockOptimistic
If recSelect.RecordCount <> 0 Then
    lvwBankbook.ListItems.Clear
    Do While Not recSelect.EOF
        Set lstItem = lvwBankbook.ListItems.Add(, , recSelect.Fields("m_name"))
            lstItem.SubItems(1) = recSelect.Fields("a_number")
            lstItem.SubItems(2) = recSelect.Fields("quantity")
            lstItem.SubItems(3) = recSelect.Fields("ID")
        recSelect.MoveNext
    Loop
End If

txtName_B = ""
txtANumber = ""
txtQuantity = ""

End Function

Private Sub lvwBankbook_ItemClick(ByVal Item As ComctlLib.ListItem)
txtName_B = lvwBankbook.SelectedItem
txtANumber = lvwBankbook.SelectedItem.SubItems(1)
txtQuantity = lvwBankbook.SelectedItem.SubItems(2)
End Sub

Private Sub lvwCheckcard_ItemClick(ByVal Item As ComctlLib.ListItem)
txtName_C = lvwCheckcard.SelectedItem
cbxBankbook = lvwCheckcard.SelectedItem.SubItems(1)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then
    Call bankbookRefresh
End If
End Sub

