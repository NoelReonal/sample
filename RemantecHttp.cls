VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecHttp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetResultaa(ByVal sConnection As String, _
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


Public Function GetResult(ByVal sConnection As String, _
                          ByVal sSQL As String, _
                          ByRef lRows As Long, _
                          Optional ByRef lErrNo As Long) As ADODB.Recordset
On Error GoTo GetResult_Err
    Dim lErrDesc As String
    Dim lErrorMsg As String
    Dim rsGetResult As New ADODB.Recordset
    Dim http As Object
    
    lErrNo = 0
    lRows = 0
    
    
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", "http://13.215.35.151:8082/xml/select/" & sSQL, False
    http.send

    If http.Status = 200 Then
        'Debug.Print http.responseText
        Set rsGetResult = rsXML(http.responseText)
        lRows = rsGetResult.RecordCount
    Else
        MsgBox "HTTP Error: " & http.Status
    End If

    'Return the recordset to the client.
    Set GetResult = rsGetResult

    Exit Function
GetResult_Err:
    'Get the error details.
    'lErrNo = oConn.Errors(0).Number
    'sErrDesc = oConn.Errors(0).Description
    
    'Write error details to event log for troubleshooting.
    'sErrorMsg = "Error #: " & CStr(lErrNo) & vbCr & _
    '            "Description: " & sErrDesc & vbCr & _
    '            "Source: " & m_sObjectName & "GetResult"
    'MsgBox sErrorMsg, vbCritical, "WARNING: Database Error!"
    Set rsGetResult = Nothing
End Function


Private Function rsXML(xml As String) As ADODB.Recordset
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim rs As ADODB.Recordset
    Dim rowNode As MSXML2.IXMLDOMNode
    Dim fieldNode As MSXML2.IXMLDOMNode

    ' Initialize objects
    Set xmlDoc = New MSXML2.DOMDocument60
    Set rs = New ADODB.Recordset

    ' Load the XML file
    If Not xmlDoc.loadXML(xml) Then
        MsgBox "Failed to load XML."
        Exit Function
    End If

    ' Get the first <row> to extract field names
    Set rowNode = xmlDoc.selectSingleNode("//row")
    If rowNode Is Nothing Then
        MsgBox "No <row> nodes found."
        Exit Function
    End If

    ' Dynamically create fields from first <row>'s child nodes
    For Each fieldNode In rowNode.childNodes
        If IsDecimal(fieldNode.Text) Then
            rs.fields.Append fieldNode.nodeName, adCurrency
        Else
            rs.fields.Append fieldNode.nodeName, adVarChar, 255
        End If
    Next

    rs.Open

    ' Loop through all <row> nodes
    For Each rowNode In xmlDoc.selectNodes("//row")
        rs.AddNew
        For Each fieldNode In rowNode.childNodes
            rs.fields(fieldNode.nodeName).value = fieldNode.Text
        Next
        rs.Update
    Next

    ' Now rs is your populated Recordset
    'MsgBox "Recordset loaded. Record count: " & rs.RecordCount
    ' Example: Accessing a field
    'rs.MoveFirst
    'MsgBox "First ACCOUNTNO: " & rs("ACCOUNTNO").Value
    
    'If rs.RecordCount > 0 Then
    '    rs.MoveFirst
    '    Do Until rs.EOF
    '        Debug.Print "ACCOUNTNO: " & rs("ACCOUNTNO").Value & " " & rs("ACCOUNTNAME").Value
    '        rs.MoveNext
    '    Loop
    'End If
    
    ' Clean up
    'rs.Close
    'Set rs = Nothing
    
    Set xmlDoc = Nothing
    Set rsXML = rs
End Function


Private Function IsDecimal(value As String) As Boolean
    On Error GoTo ErrHandler
    Dim dblValue As Double
    
    ' Attempt to convert the value to a Double
    dblValue = CDbl(value)
    
    ' It's a decimal if it contains a decimal point
    If InStr(1, value, ".") > 0 Then
        IsDecimal = True
    Else
        ' Check if it has any fractional part
        If dblValue <> Int(dblValue) Then
            IsDecimal = True
        Else
            IsDecimal = False
        End If
    End If
    Exit Function

ErrHandler:
    IsDecimal = False
End Function

