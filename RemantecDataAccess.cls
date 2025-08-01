VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecDataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************
' File:           RemantecDataAccess.cls
' Object:         Access
' Purpose:        Generic data access object
' Properties:     ExecSQL() -- Execute SQL statements and/or stored procedures.
'                 GetResult() -- Get a recordset from the database
'                 ConnectOpen() -- Open a connection
'                 ConnectClose() -- Close a connection
'***********************************************************************************
Private Const m_sObjectName As String = "RemantecDataAccess.cls"

'***********************************************************************************
' ExecSQL()
' Purpose: Execute SQL statements and/or stored procedures.
' Inputs:  sConnection     -- ADO connection string.
'          sSQL            -- SQL statement.
' Outputs: lErrNo          -- Error number returned.
' Returns: A boolean variable indicates if the action is successfull.
' Modification History
' Date        Initials  Description
' 08/14/2004  EAR       Created.
'***********************************************************************************

Public Function ExecSQL(ByVal sConnection As String, _
                        ByVal sSQL As String, _
                        Optional ByRef lErrNo As Long) As Boolean
On Error GoTo ExecSQL_Err
    Dim sErrDesc As String
    Dim sErrorMsg As String
    Dim oConn As ADODB.Connection
    Dim blnTran As Boolean

    ' Connect to the database.
    Set oConn = New ADODB.Connection
    Set oConnection = oConn
    With oConn
        .Open sConnection ', "Noel", "theboss"
        .BeginTrans
        blnTran = True
        .Execute CommandText:=sSQL
        .CommitTrans
        .Close
    End With

    If Not oConn Is Nothing Then Set oConn = Nothing
    ExecSQL = True
    Exit Function
ExecSQL_Err:
    'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "ExecSQL"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
    
    'Something goes wrong, so roll back.
    If blnTran Then
        oConn.RollbackTrans
    End If
     'Clean up.
    oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    ExecSQL = False
End Function

'***********************************************************************************
' GetResult()
' Purpose: Get an ADO recordset from an SQL action
' Inputs:  sConnection     -- ADO connection string.
'          sSQL            -- SQL statement.
' Outputs: lRows           -- Rows returned.
'          lErrNo          -- Error number returned.
' Returns: An ADO Recordset as a result of an SQL action.
' Modification History
' Date        Initials  Description
' 08/14/2004  EAR       Created.
'***********************************************************************************

Public Function GetResult(ByVal sConnection As String, _
                          ByVal sSQL As String, _
                          ByRef lRows As Long, _
                          Optional ByRef lErrNo As Long) As ADODB.Recordset
On Error GoTo GetResult_Err
    Dim oConn As New ADODB.Connection
    Dim rsGetResult As New ADODB.Recordset
    Dim lErrDesc As String
    Dim lErrorMsg As String
    
    ' Connect to the database.
    oConn.Open sConnection ', "Noel", "theboss"
    
    ' Set up the ADO recordset properties.
    With rsGetResult
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .ActiveConnection = oConn
        .Open sSQL
        lRows = .RecordCount ' rows fetched.
    End With
    
    'Return the recordset to the client.
    Set GetResult = rsGetResult
    
    'Disconnect from the database.
    Set rsGetResult.ActiveConnection = Nothing
    oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    Exit Function
GetResult_Err:
    'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "GetResult"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"

    'Clean up.
    'oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    On Error Resume Next
    Set oConn = Nothing
    Set rsGetResult = Nothing
End Function

'***********************************************************************************
' ConnectOpen()
' Purpose: Create an open connection.
' Inputs:  sConnection     -- ADO connection string.
' Outputs: oConn           -- ADO connection object.
' Returns: A boolean variable indicates if the action is successfull.
' Modification History
' Date        Initials  Description
' 08/14/2004  EAR       Created.
'***********************************************************************************
Public Function ConnectOpen(ByVal sConnection As String, ByRef oConn As ADODB.Connection) As Boolean
On Error GoTo ConnectOpen_Err
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrorMsg As String
    Dim blnTran As Boolean

    ' Connect to the database.
    Set oConn = New ADODB.Connection
    With oConn
        .Open sConnection ', "Noel", "theboss"
        .BeginTrans
    End With
    blnTran = True
    ConnectOpen = True
    Exit Function
    
ConnectOpen_Err:
    'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "ConnectOpen"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
    
    'Something goes wrong, so roll back.
    If blnTran Then
        oConn.RollbackTrans
    End If
    'Clean up.
    'oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    ConnectOpen = False
End Function

'***********************************************************************************
' ConnectClose()
' Purpose: Close a connection.
' Inputs:  oConn           -- ADO connection object.
' Outputs: bSaveOk         -- Status of Last Add, edit, delete calls.
' Returns: A boolean variable indicates if the action is successfull.
' Modification History
' Date        Initials  Description
' 08/14/2004  EAR       Created.
'***********************************************************************************
Public Function ConnectClose(ByRef oConn As ADODB.Connection, ByVal bSaveOk As Boolean) As Boolean
On Error GoTo ConnectClose_Err
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrorMsg As String
   
    If bSaveOk Then
        oConn.CommitTrans
        oConn.Close
        If Not oConn Is Nothing Then Set oConn = Nothing
    Else
        GoTo ConnectClose_Err
    End If
    ConnectClose = True
    Exit Function
    
ConnectClose_Err:
    'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "ConnectClose"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
    
    'Something goes wrong, so roll back.
    If Not bSaveOk Then
        oConn.RollbackTrans
    End If
    'Clean up.
    oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    ConnectClose = False
End Function


Public Function OpenRecordset(ByRef oConn As ADODB.Connection, _
    ByRef rs As ADODB.Recordset, ByVal sSQL As String, _
    Optional CursorType As CursorTypeEnum, Optional LockType As LockTypeEnum, _
    Optional CursorLocationType As CursorLocationEnum) As Boolean
On Error GoTo ErrorHandler
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrorMsg As String
    
    ' create a recordset using the provided collection
    Set rs = New ADODB.Recordset
    rs.CursorLocation = IIf(CursorLocationType = 0, adUseServer, CursorLocationType)
    'rs.Open sSQL, cn, IIf(CursorType = 0, adOpenForwardOnly, CursorType), _
    '    IIf(LockType = 0, adLockReadOnly, LockType)
         
    rs.Open sSQL, oConn, IIf(CursorType = 0, adOpenForwardOnly, CursorType), _
        IIf(LockType = 0, adLockReadOnly, LockType)
    On Error GoTo 0
    
    OpenRecordset = True
    Exit Function
    
ErrorHandler:
    OpenRecordset = False
     'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    
    If Not rs Is Nothing Then Set rs = Nothing
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "OpenRecordset"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
End Function

Public Function RecordRead(ByVal sSQL As String, ByRef rs As ADODB.Recordset, ByVal cn As ADODB.Connection) As Boolean
    If OpenRecordset(cn, rs, sSQL, adOpenForwardOnly, adLockReadOnly) Then
        RecordRead = True
    End If
End Function

Public Function RecordWrite(ByVal sSQL As String, ByRef rs As ADODB.Recordset, ByVal cn As ADODB.Connection) As Boolean
    'If OpenRecordset(cn, rs, sSQL, adOpenDynamic, adLockOptimistic, adUseClient) Then
    If OpenRecordset(cn, rs, sSQL, adOpenKeyset, adLockOptimistic, adUseServer) Then
        RecordWrite = True
    End If
End Function

'Public Function GetLastNo(cKey As String, nLastNo As Long, ByVal cn As ADODB.Connection) As Boolean
'    Dim rs As ADODB.Recordset
'    Dim sSQL As String
'
'    sSQL = "Select * From LastNo where ID='" & cKey & "'"
'    If RecordWrite(sSQL, rs, cn) Then
'        If Not (rs.BOF And rs.EOF) Then
'            rs.MoveFirst
'            If rs("LastNo").Value = 2147483647 Then
'                MsgBox "WARNING: Database Full!", vbCritical, "DATABASE ERROR"
'                rs("LastNo") = 1
'                GetLastNo = False
'            Else
'                rs("LastNo") = rs("LastNo") + 1
'                GetLastNo = True
'            End If
'        Else
'            rs.AddNew
'            rs("ID") = cKey
'            rs("LastNo") = 1
'            GetLastNo = True
'        End If
'        rs.Update
'        nLastNo = rs("LastNo")
'        rs.Close
'        Set rs = Nothing
'    End If
'End Function

'Public Function CreateID(cKey As String, Optional ByVal cn As ADODB.Connection) As String
'    Dim nLastNo As Long
'
'    If pIsGUID Or cKey = "GUID" Then
'        CreateID = CreateGUID()
'    ElseIf GetLastNo(cKey, nLastNo, cn) Then
'        CreateID = cKey & Format(nLastNo, "0000000000")
'    Else
'        CreateID = ""
'    End If
'End Function

Public Function GetLastNo(cKey As String, nLastNo As Long, ByVal sConnection As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim cn As ADODB.Connection
    Dim lOk As Boolean
    
    If ConnectOpen(sConnection, cn) Then
        sSQL = "Select * From LastNo where ID='" & cKey & "'"
        If RecordWrite(sSQL, rs, cn) Then
            If Not (rs.BOF And rs.EOF) Then
                rs.MoveFirst
                If rs("LastNo").Value = 2147483647 Then
                    MsgBox "WARNING: Database Full!", vbCritical, "DATABASE ERROR"
                    rs("LastNo") = 1
                    lOk = False
                Else
                    rs("LastNo") = rs("LastNo") + 1
                    rs.Update
                    nLastNo = rs("LastNo")
                    lOk = True
                End If
            Else
                rs.AddNew
                rs("ID") = cKey
                rs("LastNo") = 1
                rs.Update
                nLastNo = rs("LastNo")
                lOk = True
            End If
            rs.Close
            Set rs = Nothing
        End If
        If lOk Then
            ConnectClose cn, True
        Else
            ConnectClose cn, False
        End If
        GetLastNo = lOk
    End If
End Function

Public Function CreateID(cKey As String, Optional ByVal sConnection As String) As String
    Dim nLastNo As Long

    If pIsGUID Or cKey = "GUID" Then
        CreateID = CreateGUID()
    ElseIf GetLastNo(cKey, nLastNo, sConnection) Then
        'CreateID = cKey & Format(nLastNo, "0000000000")
        '* * * by:elad 2005.09.02 - client request
        If pClientName = "GOLDTOWN" Then
            CreateID = cKey & Format(nLastNo, "000000")
        Else
            CreateID = cKey & Format(nLastNo, "0000000000")
        End If
        '* * *
    Else
        CreateID = ""
    End If
End Function

Public Function GetSysLastNo(cKey As String, ByVal sConnection As String, ByVal cFormat As String) As String
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    Dim cn As ADODB.Connection
    Dim lOk As Boolean
    Dim nLastNo As Long
    
    If ConnectOpen(sConnection, cn) Then
        sSQL = "Select * From SysLastNo where ID='" & cKey & "'"
        If RecordWrite(sSQL, rs, cn) Then
            If Not (rs.BOF And rs.EOF) Then
                rs.MoveFirst
                If rs("LastNo").Value = 2147483647 Then
                    MsgBox "WARNING: Database Full!", vbCritical, "DATABASE ERROR"
                    rs("LastNo") = 1
                    lOk = False
                Else
                    rs("LastNo") = rs("LastNo") + 1
                    rs.Update
                    nLastNo = rs("LastNo")
                    lOk = True
                End If
            Else
                rs.AddNew
                rs("ID") = cKey
                rs("LastNo") = 1
                rs.Update
                nLastNo = rs("LastNo")
                lOk = True
            End If
            rs.Close
            Set rs = Nothing
        End If
        If lOk Then
            ConnectClose cn, True
        Else
            ConnectClose cn, False
        End If
        If lOk Then
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
            GetSysLastNo = Format(nLastNo, cFormat)
        Else
            GetSysLastNo = Null
        End If
    End If
End Function


'***********************************************************************************
' ConnectOpenDbf()
' Purpose: Create an open connection.
' Inputs:  cDbfPath        -- Directory or Path of DBF files.
' Outputs: oConn           -- ADO connection object.
' Returns: A boolean variable indicates if the action is successfull.
' Modification History
' Date        Initials  Description
' 11/22/2004  EAR       Created.
'***********************************************************************************
Public Function ConnectOpenDBF(cDbfPath As String, ByRef oConn As ADODB.Connection) As Boolean
On Error GoTo ConnectOpen_Err
    
    Set oConn = New ADODB.Connection
    oConn.CursorLocation = adUseClient
    oConn.Open "Driver={Microsoft dBASE Driver (*.dbf)};" & _
                    "DriverID=277;" & _
                    "Dbq=" & cDbfPath & ";"
    On Error GoTo 0
    ConnectOpenDBF = True
    Exit Function
    
ConnectOpen_Err:
    ConnectOpenDBF = False
    'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    'Clean up.
    'oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "ConnectOpen"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
End Function

'***********************************************************************************
' ConnectOpenVFDbf()
' Purpose: Create an open connection.
' Inputs:  cDbfPath        -- Directory or Path of DBF files.
' Outputs: oConn           -- ADO connection object.
' Returns: A boolean variable indicates if the action is successfull.
' Modification History
' Date        Initials  Description
' 11/22/2004  EAR       Created.
'***********************************************************************************
Public Function ConnectOpenFoxProDBF(cDbfPath As String, ByRef oConn As ADODB.Connection, Optional ByRef IsDbf As Boolean) As Boolean
On Error GoTo ConnectOpen_Err
    Dim cSourceType As String
    
    If IsDbf Then
        cSourceType = "DBF"
    Else
        cSourceType = "DBC"
    End If
    Set oConn = New ADODB.Connection
    oConn.CursorLocation = adUseClient
    oConn.Open "Driver={Microsoft Visual FoxPro Driver};" & _
           "SourceType=" & cSourceType & ";" & _
           "SourceDB=" & cDbfPath & ";" & _
           "Exclusive=No"
    On Error GoTo 0
    ConnectOpenFoxProDBF = True
    Exit Function
    
ConnectOpen_Err:
    ConnectOpenFoxProDBF = False
    'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    'Clean up.
    'oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "ConnectOpen"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
End Function


Public Sub ErrorMessage(ByVal oConn As ADODB.Connection)
    Dim lErrNo As Long
    Dim sErrDesc As String
    Dim sErrorMsg As String
    
   'Get the error details.
    lErrNo = oConn.Errors(0).Number
    sErrDesc = oConn.Errors(0).Description
    
    'Write error details to event log for troubleshooting.
    sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
                "Description: " & sErrDesc & vbCr & _
                "Source: " & m_sObjectName & "OpenRecordset"
    MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
End Sub


Public Function NextID(cCmd As String, oConn As ADODB.Connection) As Long
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


Public Sub SetGenID(ByVal sConnection As String, cFldName As String, cFileName As String, cGenID As String)
On Error GoTo HandleError
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT " & cFldName & " From " & cFileName & " ORDER BY " & cFldName & " DESC"
    Set rs = GetResult(sConnection, sSQL, 0)
    If rs.RecordCount > 0 Then
        ExecSQL sConnection, "SET GENERATOR " & cGenID & " TO " & rs(cFldName).Value
    Else
        ExecSQL sConnection, "SET GENERATOR " & cGenID & " TO 0"
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
HandleError:
    MsgBox "ERROR: SetGenID()"
End Sub


Public Function GetNextID(ByVal sConnection As String, cCmd As String) As Long
On Error GoTo HandleError
    Dim oData As New RemantecDataAccess
    Dim oConn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim param As ADODB.Parameter
    
    If oData.ConnectOpen(sConnection, oConn) Then
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oConn
        cmd.CommandText = cCmd
        cmd.CommandType = adCmdStoredProc
        Set param = cmd.CreateParameter("NEXT_ID", adBigInt, adParamOutput, , Null)
        cmd.Parameters.Append param
        cmd.Execute
        GetNextID = IIf(IsNull(cmd.Parameters("NEXT_ID").Value), 0, cmd.Parameters("NEXT_ID").Value)
        oData.ConnectClose oConn, True
        Exit Function
    End If
HandleError:
    MsgBox "ERROR: GetNextID()"
    oData.ConnectClose oConn, False
End Function

Public Function IsTableExist(ByVal sConnection As String, ByVal sTable As String) As Boolean
On Error GoTo GetResult_Err
    Dim oConn As New ADODB.Connection
    Dim rsGetResult As New ADODB.Recordset
    
    ' Connect to the database.
    oConn.Open sConnection ', "Noel", "theboss"
    
    ' Set up the ADO recordset properties.
    With rsGetResult
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .ActiveConnection = oConn
        .Open "SELECT * FROM " & sTable
        'lRows = .RecordCount ' rows fetched.
    End With
    'Disconnect from the database.
    Set rsGetResult.ActiveConnection = Nothing
    oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    IsTableExist = True
    Exit Function
GetResult_Err:
    If Not oConn Is Nothing Then Set oConn = Nothing
    On Error Resume Next
    Set oConn = Nothing
    Set rsGetResult = Nothing
    IsTableExist = False
End Function

Public Function IsFieldExist(ByVal sConnection As String, ByVal sTable As String, ByVal sField As String) As Boolean
On Error GoTo GetResult_Err
    Dim oConn As New ADODB.Connection
    Dim rsGetResult As New ADODB.Recordset
    
    ' Connect to the database.
    oConn.Open sConnection ', "Noel", "theboss"
    
    ' Set up the ADO recordset properties.
    With rsGetResult
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .ActiveConnection = oConn
        .Open "SELECT " & sField & " FROM " & sTable
        'lRows = .RecordCount ' rows fetched.
    End With
    'Disconnect from the database.
    Set rsGetResult.ActiveConnection = Nothing
    oConn.Close
    If Not oConn Is Nothing Then Set oConn = Nothing
    IsFieldExist = True
    Exit Function
GetResult_Err:
    If Not oConn Is Nothing Then Set oConn = Nothing
    On Error Resume Next
    Set oConn = Nothing
    Set rsGetResult = Nothing
    IsFieldExist = False
End Function

Public Function NextNo(ByRef oConn As ADODB.Connection, ByVal cKey As String, Optional ByVal cFormat As String) As String
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
