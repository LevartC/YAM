Attribute VB_Name = "Module"
Public adoConnect As ADODB.Connection

Public g_Date As Date
Public g_MainDate As Date
Public l_ID As String
Public b_Login As Boolean
Public b_Switch As Boolean

Public stackForm As Form

Public Function updateCBBalance(strType As String, strCBType As String, lblUpdate As Label)
Dim recTemp As New ADODB.Recordset
Dim i_temp As Integer
Select Case strType
Case "����"
    recTemp.Open "SELECT Cash FROM login WHERE ID = '" & l_ID & "'", adoConnect, adOpenStatic, adLockOptimistic
    lblUpdate = recTemp.Fields("Cash") & " ��"
    recTemp.Close
Case "����"
    recTemp.Open "SELECT quantity FROM bankbook WHERE m_name = '" & strCBType & "'", adoConnect, adOpenStatic, adLockOptimistic
    lblUpdate = recTemp.Fields("quantity") & " ��"
    recTemp.Close
Case "üũī��"
    recTemp.Open "SELECT * FROM checkcard where m_name = '" & strCBType & "'", adoConnect, adOpenStatic, adLockOptimistic
    i_temp = recTemp.Fields("bankbook_id")
    recTemp.Close
    recTemp.Open "SELECT quantity FROM bankbook WHERE ID = " & i_temp, adoConnect, adOpenStatic, adLockOptimistic
    lblUpdate = recTemp.Fields("quantity") & " ��"
    recTemp.Close
End Select
End Function
