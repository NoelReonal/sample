VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecReadWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sConnection As String
Public oConn As ADODB.Connection
Public oData As New RemantecDataAccess

Dim sSQL As String
   
Public Function Record_Cancel(cID As String, cTblName As String, IsCancel As Boolean) As Boolean
On Error GoTo HandleError
    Dim nCancel As Integer
    
    'If MsgBox("Are you sure you want to " & IIf(IsCancel, "CANCEL", "UNCANCEL") & " this record(s)?", vbYesNo, "Verify") = vbNo Then
    '    Exit Function
    'End If
            
    If oData.ConnectOpen(sConnection, oConn) Then
        nCancel = IIf(IsCancel, 0, 1)
        If nCancel = 0 Then
            sSQL = "UPDATE " & cTblName & " SET CANCEL = " & nCancel & ", DATE_CANCEL = Null WHERE ID=" & cID
        Else
            sSQL = "UPDATE " & cTblName & " SET CANCEL = " & nCancel & ", DATE_CANCEL = " & Now() & " WHERE ID=" & cID
        End If
        oConn.Execute sSQL
        oData.ConnectClose oConn, True
        Record_Cancel = True
        Exit Function
    End If
HandleError:
    Record_Cancel = False
    MsgBox "ERROR: Record_Cancel()"
    oData.ConnectClose oConn, False
End Function

Public Function Record_Delete(ByVal cID As String, cTblName As String) As Boolean
On Error GoTo HandleError

    'If MsgBox("Are you sure you want to DELETE this record(s)?", vbYesNo, "Verify") = vbNo Then
    '    Exit Function
    'End If
        
    If oData.ConnectOpen(sConnection, oConn) Then
        sSQL = "DELETE From " & cTblName & " WHERE ID=" & cID
        oConn.Execute sSQL
        oData.ConnectClose oConn, True
        Record_Delete = True
        Exit Function
    End If
HandleError:
    Record_Delete = False
    MsgBox "ERROR: Record_Delete()"
    oData.ConnectClose oConn, False
End Function

Public Function Record_Save(ByVal cID As String, ByVal cTblName As String, ByRef rs As ADODB.Recordset) As Boolean
On Error GoTo HandleError

    'If MsgBox("Are you sure you want to save this record(s)?", vbYesNo, "VERIFY") = vbNo Then
    '    Exit Function
    'End If
    
    If oData.ConnectOpen(sConnection, oConn) Then
        sSQL = "Select * From " & cTblName & " WHERE ID=" & cID
        If Not oData.RecordWrite(sSQL, rs, oConn) Then
            GoTo HandleError
        End If
        Record_Save = True
    End If
    Exit Function
HandleError:
    Record_Save = False
    MsgBox "ERROR: Record_Save()"
    oData.ConnectClose oConn, False
End Function


Public Function Record_Close(ByRef rs As ADODB.Recordset) As Boolean
On Error GoTo HandleError

    rs.Close
    Set rs = Nothing
    oData.ConnectClose oConn, True
    Record_Close = True
    Exit Function
HandleError:
    Record_Close = False
    MsgBox "ERROR: Record_Close()"
    oData.ConnectClose oConn, False
End Function

Public Sub ErrorMessage()
    oData.ErrorMessage oConn
    oData.ConnectClose oConn, False
End Sub

Public Function NextNo(ByVal cKey As String, Optional ByVal cFormat As String) As String
On Error GoTo HandleError
    Dim rs As ADODB.Recordset
    Dim nLastNo As Long
    Dim oData As New RemantecDataAccess
    
    sSQL = "Select * From M_SOURCECD where Code='" & cKey & "'"
    If oData.RecordWrite(sSQL, rs, oConn) Then
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveFirst
            If rs("LastNo").Value = 2147483647 Then
                MsgBox "WARNING: Database Full!", vbCritical, "DATABASE ERROR"
                GoTo HandleError
            Else
                rs("LastNo") = rs("LastNo") + 1
                rs.Update
                nLastNo = rs("LastNo")
            End If
        Else
            rs.AddNew
            rs("Code") = cKey
            rs("Name") = "SOURCE CODE : " & cKey
            rs("LastNo") = 1
            rs.Update
            nLastNo = rs("LastNo")
        End If
        rs.Close
        Set rs = Nothing
    End If
    If cFormat = "" Then
        Select Case nLastNo
            Case Is <= 9999999
                cFormat = "0000000"
            Case Is <= 99999999
                cFormat = "00000000"
            Case Else
                cFormat = "0000000000"
        End Select
    End If
    NextNo = cKey & "#" & Format(nLastNo, cFormat)
    Exit Function
HandleError:
    NextNo = ""
    MsgBox "ERROR: NextNo()"
End Function

Public Function NextID(cCmd As String) As Long
On Error GoTo HandleError
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oConn
    cmd.CommandText = cCmd
    cmd.CommandType = adCmdStoredProc
    Set param = cmd.CreateParameter("NEXT_ID", adBigInt, adParamOutput, , Null)
    cmd.Parameters.Append param
    cmd.Execute
    NextID = IIf(IsNull(cmd.Parameters("NEXT_ID").Value), 0, cmd.Parameters("NEXT_ID").Value)
    Exit Function
HandleError:
    MsgBox "ERROR: NextID()"
    NextID = 0
End Function

