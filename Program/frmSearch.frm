VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  '단일 고정
   Caption         =   "조회"
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
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   2  '화면 가운데
   Begin YAM.CandyButton cmdBack 
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "뒤로"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8454016
      ColorButtonDown =   16777215
      BorderBrightness=   0
      ColorBright     =   16777088
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdSearch 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   1750
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "검색"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   5
      Checked         =   0   'False
      ColorButtonHover=   49152
      ColorButtonUp   =   16761024
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin ComctlLib.ListView lvwList 
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8070
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtMonth 
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtYear 
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin YAM.CandyButton cmdDelete 
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "삭제"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8454016
      ColorButtonDown =   16777215
      BorderBrightness=   0
      ColorBright     =   16777088
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdUpdate 
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "수정"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   8454143
      ColorButtonUp   =   8454016
      ColorButtonDown =   16777215
      BorderBrightness=   0
      ColorBright     =   16777088
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label lblImport_M 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0 원"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4695
      TabIndex        =   7
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label lblExpense_M 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0 원"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "월간 지출"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   5
      Top             =   2400
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "월간 수입"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   1755
   End
   Begin VB.Image imgMonthNext 
      Height          =   600
      Left            =   6000
      Picture         =   "frmSearch.frx":11B40
      Stretch         =   -1  'True
      Top             =   960
      Width           =   720
   End
   Begin VB.Image imgMonthPrev 
      Height          =   600
      Left            =   240
      Picture         =   "frmSearch.frx":12E8B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "월"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   1005
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "년"
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   24
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1005
      Width           =   495
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_Date As Date

Private Sub cmdBack_Click()
Unload frmSearch
End Sub

Private Sub cmdSearch_Click()
m_Date = CDate(txtYear & "-" & txtMonth & "-01")
Call updateSearch
End Sub

Private Sub cmdDelete_Click()
If MsgBox("정말로 삭제하시겠습니까?", vbYesNo, "삭제") = vbYes Then
    '타입 지정
    Dim query As String
    Dim strID As String
    Dim strType As String
    Dim bank_ID As Long
    Dim lQty As Long
    strID = lvwList.SelectedItem.SubItems(5)
    Dim recSelect As New ADODB.Recordset
    recSelect.Open "SELECT * FROM uselog where ID = " & strID, adoConnect, adOpenStatic, adLockOptimistic

If recSelect.RecordCount <> 0 Then
    lQty = recSelect.Fields("quantity")
    If recSelect.Fields("bankbook_id") <> "" Then
        bank_ID = recSelect.Fields("bankbook_id")
    End If
    '속성 갱신
    Select Case recSelect.Fields("div_type")
    Case "지출"
        strType = "category_e"
        '자금내역 복구
        Select Case recSelect.Fields("c_type")
        Case "현금"
            query = "update login set Cash = Cash + " & lQty & " where ID = '" & l_ID & "'"
        Case "통장"
            query = "update bankbook set quantity = quantity + " & lQty & " where ID = " & bank_ID
        Case "체크카드"
            query = "update bankbook set quantity = quantity + " & lQty & " where ID = " & bank_ID
        End Select
        adoConnect.Execute query
    Case "수입"
        strType = "category_i"
        '자금내역 복구
        Select Case recSelect.Fields("c_type")
        Case "현금"
            query = "update login set Cash = Cash - " & lQty & " where ID = '" & l_ID & "'"
        Case "통장"
            query = "update bankbook set quantity = quantity - " & lQty & " where ID = " & bank_ID
        End Select
        adoConnect.Execute query
    Case "출금"
        strType = "category_m"
        '자금내역 복구
        query = "update bankbook set quantity = quantity + " & (lQty + CLng(recSelect.Fields("commision"))) & " where ID = " & bank_ID
        adoConnect.Execute query
        query = "update login set Cash = Cash - " & lQty & " where ID = '" & l_ID & "'"
        adoConnect.Execute query
    Case "입금"
        strType = "category_m"
        query = "update bankbook set quantity = quantity - " & lQty & " where ID = " & bank_ID
        adoConnect.Execute query
        query = "update login set Cash = Cash + " & lQty & " where ID = '" & l_ID & "'"
        adoConnect.Execute query
    End Select
    '카테고리 사용 빈도 제거
    query = "update " & strType & " set fou = fou - 1 where ID = " & recSelect.Fields("category_id")
    adoConnect.Execute query
    '삭제 쿼리 작성
    query = "delete from uselog where ID = " & strID
    adoConnect.Execute query
    MsgBox ("정상적으로 삭제되었습니다.")
    Call updateSearch
Else
    MsgBox "삭제할 아이템이 존재하지 않습니다."
End If

End If
End Sub


Private Sub cmdUpdate_Click()
g_Date = CDate(Mid(lvwList.SelectedItem, 1, 4) & "-" & Mid(lvwList.SelectedItem, 5, 2) & "-" & Mid(lvwList.SelectedItem, 7, 2))
frmSearch.Hide
frmSearch_Update.Show
End Sub

Private Sub Form_Activate()
If b_Switch Then
Call updateSearch
End If
End Sub

Private Sub Form_Load()
m_Date = Date
txtYear = Year(m_Date)
txtMonth = Month(m_Date)

MakeColumns
Call updateSearch

End Sub


Private Sub MakeColumns()

With lvwList
.View = lvwReport
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "날짜", 900
.ColumnHeaders.Add , , "카테고리", 900
.ColumnHeaders.Add , , "금액", 800
.ColumnHeaders.Add , , "내용", 1500
.ColumnHeaders.Add , , "수단", 700
.ColumnHeaders.Add , , "ID", 0
.ColumnHeaders.Add , , "유형", 0
.ColumnHeaders.Add , , "수수료", 400
End With
End Sub

Private Function updateSearch()

txtYear = Year(m_Date)
txtMonth = Month(m_Date)

Dim recSum As New ADODB.Recordset
'월별 지출
recSum.Open "SELECT sum(quantity) as q_sum FROM uselog WHERE div_type = '지출' and m_date like '" & Format(m_Date, "YYYYMM") & "%'", adoConnect, adOpenStatic, adLockOptimistic
If recSum.Fields("q_sum") <> "" Then
    lblExpense_M = recSum.Fields("q_sum") & " 원"
Else
    lblExpense_M = "0 원"
End If
recSum.Close
'월별 수입
recSum.Open "SELECT sum(quantity) as q_sum FROM uselog WHERE div_type = '수입' and m_date like '" & Format(m_Date, "YYYYMM") & "%'", adoConnect, adOpenStatic, adLockOptimistic
If recSum.Fields("q_sum") <> "" Then
    lblImport_M = recSum.Fields("q_sum") & " 원"
Else
    lblImport_M = "0 원"
End If
recSum.Close

Dim lstItem As ListItem
Dim recSelect As New ADODB.Recordset
Dim recCat_E As New ADODB.Recordset
Dim recCat_I As New ADODB.Recordset
Dim recCat_M As New ADODB.Recordset
Dim s_temp As String

recSelect.Open "SELECT * FROM uselog where m_date like '" & Format(m_Date, "YYYYMM") & "%' ORDER BY m_date desc", adoConnect, adOpenStatic, adLockOptimistic
recCat_E.Open "SELECT * FROM category_e", adoConnect, adOpenStatic, adLockOptimistic
recCat_I.Open "SELECT * FROM category_i", adoConnect, adOpenStatic, adLockOptimistic
recCat_M.Open "SELECT * FROM category_m", adoConnect, adOpenStatic, adLockOptimistic

lvwList.ListItems.Clear
If recSelect.RecordCount <> 0 Then
    lvwList.ListItems.Clear
    Do While Not recSelect.EOF
        Set lstItem = lvwList.ListItems.Add(, , recSelect.Fields("m_date"))
            Select Case recSelect.Fields("div_type")
            Case "지출"
                If recCat_E.RecordCount <> 0 Then
                    recCat_E.MoveFirst
                    Do While Not recCat_E.EOF
                        s_temp = recSelect.Fields("category_id")
                        If s_temp = recCat_E.Fields("ID") Then
                            lstItem.SubItems(1) = recCat_E.Fields("m_name")
                            recCat_E.MoveLast
                        End If
                        recCat_E.MoveNext
                    Loop
                    lstItem.SubItems(2) = "-" & recSelect.Fields("quantity")
                End If
            Case "수입"
                If recCat_I.RecordCount <> 0 Then
                    recCat_I.MoveFirst
                    Do While Not recCat_I.EOF
                        s_temp = recSelect.Fields("category_id")
                        If s_temp = recCat_I.Fields("ID") Then
                            lstItem.SubItems(1) = recCat_I.Fields("m_name")
                            recCat_I.MoveLast
                        End If
                        recCat_I.MoveNext
                    Loop
                    lstItem.SubItems(2) = "+" & recSelect.Fields("quantity")
                End If
            Case "출금"
                If recCat_M.RecordCount <> 0 Then
                    recCat_M.MoveFirst
                    Do While Not recCat_M.EOF
                        s_temp = recSelect.Fields("category_id")
                        If s_temp = recCat_M.Fields("ID") Then
                            lstItem.SubItems(1) = recCat_M.Fields("m_name")
                            recCat_M.MoveLast
                        End If
                        recCat_M.MoveNext
                    Loop
                    lstItem.SubItems(2) = "±" & recSelect.Fields("quantity")
                    lstItem.SubItems(7) = recSelect.Fields("commision")
                End If
            Case "입금"
                If recCat_M.RecordCount <> 0 Then
                    recCat_M.MoveFirst
                    Do While Not recCat_M.EOF
                        s_temp = recSelect.Fields("category_id")
                        If s_temp = recCat_M.Fields("ID") Then
                            lstItem.SubItems(1) = recCat_M.Fields("m_name")
                            recCat_M.MoveLast
                        End If
                        recCat_M.MoveNext
                    Loop
                    lstItem.SubItems(2) = "±" & recSelect.Fields("quantity")
                End If
            End Select
            lstItem.SubItems(3) = recSelect.Fields("e_memo")
            lstItem.SubItems(4) = recSelect.Fields("c_type")
            lstItem.SubItems(5) = recSelect.Fields("ID")
            lstItem.SubItems(6) = recSelect.Fields("div_type")
        recSelect.MoveNext
    Loop
End If
End Function
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

m_Date = m_Date + m_day

Call updateSearch

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

m_Date = m_Date - m_day

Call updateSearch

End Sub


Private Sub Form_Unload(Cancel As Integer)
g_Date = g_MainDate
frmMain.Show
End Sub

