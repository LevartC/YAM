VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSet_Category 
   BorderStyle     =   1  '단일 고정
   Caption         =   "카테고리 추가 및 수정"
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
   Picture         =   "frmCategorySet.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   7260
   StartUpPosition =   2  '화면 가운데
   Begin YAM.CandyButton cmdAdd 
      Height          =   615
      Left            =   4200
      TabIndex        =   6
      Top             =   5640
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "추　가"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16744703
      ColorButtonUp   =   16744576
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16761024
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdBack 
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   8520
      Width           =   6615
      _ExtentX        =   11668
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
      Caption         =   "뒤　　로"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   16744703
      ColorButtonUp   =   16744576
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16761024
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   840
      TabIndex        =   15
      Top             =   1200
      Width           =   5535
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "수입"
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
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "이동"
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
         Index           =   2
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FFFFFF&
         Caption         =   "지출"
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
         Left            =   480
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCatFOU 
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtCatName 
      BeginProperty Font 
         Name            =   "HY바다L"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   3600
      Width           =   2655
   End
   Begin ComctlLib.ListView lvwList 
      Height          =   5055
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8916
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin YAM.CandyButton cmdUpdate 
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   6480
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "수　정"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16744703
      ColorButtonUp   =   16744576
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16761024
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin YAM.CandyButton cmdDelete 
      Height          =   615
      Left            =   4200
      TabIndex        =   8
      Top             =   7320
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "삭　제"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   2
      Checked         =   0   'False
      ColorButtonHover=   16744703
      ColorButtonUp   =   16744576
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16761024
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '투명
      Caption         =   "수정하실 카테고리를 목록에서 클릭한 후 수정하세요."
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
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   6855
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사용 빈도"
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
      Left            =   4200
      TabIndex        =   13
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "빈도수가 클수록 가계부 작성시 상위에 표시됩니다."
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
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   6615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "지출 / 수입 카테고리를 추가 / 수정 합니다."
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
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   6615
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      BackStyle       =   0  '투명
      Caption         =   "카테고리 이름"
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
      Left            =   4200
      TabIndex        =   10
      Top             =   3240
      Width           =   2655
   End
End
Attribute VB_Name = "frmSet_Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
'중복 체크
Dim c_error As Boolean
c_error = False
If txtCatName = "" Then
    MsgBox ("카테고리 이름을 입력하세요.")
    c_error = True
    txtCatName.SetFocus
ElseIf txtCatFOU = "" Then
    txtCatFOU = "0"
End If
For i = 1 To lvwList.ListItems.Count
    If txtCatName = lvwList.ListItems(i) Then
        MsgBox ("중복된 이름입니다. 다른 이름을 입력하세요.")
        c_error = True
        txtCatName.SetFocus
    End If
Next

'타입 지정
Dim strType As String
If optType(0).Value = True Then
    strType = "category_e"
ElseIf optType(1).Value = True Then
    strType = "category_i"
Else
    strType = "category_m"
End If

'쿼리 작성
If Not c_error Then
    Dim query As String
    query = "insert into " & strType & "(m_name, fou) values('" & txtCatName & "', " & txtCatFOU & ")"
    adoConnect.Execute query
    MsgBox ("성공적으로 추가되었습니다.")
    Call optSelect
    txtCatName = ""
    txtCatFOU = ""
End If
End Sub

Private Sub cmdBack_Click()
Unload frmSet_Category
End Sub

Private Sub cmdDelete_Click()
If MsgBox("정말로 삭제하시겠습니까?", vbYesNo, "삭제") = vbYes Then
    '타입 지정
    Dim strType As String
    If optType(0).Value = True Then
        strType = "category_e"
    ElseIf optType(1).Value = True Then
        strType = "category_i"
    Else
        strType = "category_m"
    End If
    
    '쿼리 작성
    Dim query As String
    query = "delete from " & strType & " where m_name = '" & lvwList.SelectedItem & "'"
    adoConnect.Execute query
    MsgBox ("정상적으로 삭제되었습니다.")
    Call optSelect
    txtCatName = ""
    txtCatFOU = ""
End If
End Sub

Private Sub cmdUpdate_Click()
'중복 체크
Dim c_error As Boolean
c_error = False
If txtCatName = "" Then
    MsgBox ("카테고리 이름을 입력하세요.")
    c_error = True
    txtCatName.SetFocus
ElseIf txtCatFOU = "" Then
    txtCatFOU = "0"
End If
For i = 1 To lvwList.ListItems.Count
    If txtCatName = lvwList.ListItems(i) And Not lvwList.SelectedItem = lvwList.ListItems(i) Then
        MsgBox ("중복된 이름입니다. 다른 이름을 입력하세요.")
        c_error = True
        txtCatName.SetFocus
    End If
Next

'타입 지정
Dim strType As String
If optType(0).Value = True Then
    strType = "category_e"
ElseIf optType(1).Value = True Then
    strType = "category_i"
Else
    strType = "category_m"
End If

'쿼리 작성
If Not c_error Then
    Dim query As String
    query = "update " & strType & " set m_name = '" & txtCatName & "', fou = " & txtCatFOU & " where m_name = '" & lvwList.SelectedItem & "'"
    adoConnect.Execute query
    MsgBox ("정상적으로 수정되었습니다.")
    Call optSelect
    txtCatName = ""
    txtCatFOU = ""
End If
End Sub

Private Sub Form_Load()
MakeColumns

Call optSelect

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmSet_Main.Show
End Sub

Private Sub MakeColumns()

With lvwList
.View = lvwReport
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "이름", 1500
.ColumnHeaders.Add , , "사용 빈도", 1500
End With
End Sub



Private Sub lvwList_ItemClick(ByVal Item As ComctlLib.ListItem)

txtCatName = lvwList.SelectedItem
txtCatFOU = lvwList.SelectedItem.SubItems(1)
End Sub

Private Sub optType_Click(Index As Integer)
Call optSelect
End Sub

Private Function optSelect()

Dim lstItem As ListItem
Dim recSelect As New ADODB.Recordset
Dim s_temp As String

If optType(0).Value Then
    s_temp = "category_e"
ElseIf optType(1).Value Then
    s_temp = "category_i"
Else
    s_temp = "category_m"
End If

recSelect.Open "SELECT m_name, fou FROM " & s_temp & " ORDER BY fou desc", adoConnect, adOpenStatic, adLockOptimistic
If recSelect.RecordCount <> 0 Then
    lvwList.ListItems.Clear
    Do While Not recSelect.EOF
        Set lstItem = lvwList.ListItems.Add(, , recSelect.Fields("m_name"))
            lstItem.SubItems(1) = recSelect.Fields("fou")
        recSelect.MoveNext
    Loop
End If
    
End Function

