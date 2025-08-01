Attribute VB_Name = "mdlSupport"
Public Function GetKeyValue(frm As Form) As String
    Dim ctrl As Control
    Dim keyValue As New Collection
    Dim values As String
    
    values = ""
    For Each ctrl In frm.Controls
        On Error Resume Next
        Dim key As String, val As String

        'key = ctrl.Name
        key = ctrl.DataField
        Select Case True
            Case TypeOf ctrl Is TextBox
                val = ctrl.Text
            Case TypeOf ctrl Is ComboBox
                val = ctrl.Text
            Case TypeOf ctrl Is CheckBox
                val = CStr(ctrl.Value)
            Case TypeOf ctrl Is OptionButton
                val = CStr(ctrl.Value)
            Case Else
                val = ""
        End Select
        If key <> "" And val <> "" Then
            If Not IsNumeric(val) Then
                val = "'" & val & "'"
            End If
            If values = "" Then
                values = key & "=" & val
            Else
                values = values + ", " + key & "=" & val
            End If
        End If
    Next
    GetKeyValue = values
End Function



