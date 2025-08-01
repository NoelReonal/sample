VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecSecurityAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum enBtn
    btnAdd = 0
    btnBrowse
    btnDup
    btnEdit
    btnDel
    btnPrint
    btnSave
    btnExit
End Enum

Public canOpen As Boolean
Public canAdminister As Boolean
Public canRead As Boolean
Public canAdd As Boolean
Public canEdit As Boolean
Public canDel As Boolean
Public canCancel As Boolean
Public canAccessAll As Boolean

Public cObjectName As String
Public cObjectDesc As String
Public cObjectType As String
Public cLogInUser As String

Public Function GetPermissions(ByVal sConnection As String, ByVal cObjName As String, _
ByVal cLogUser As String) As Boolean

    
    
    PermissionValue False       'set all permission to .f.

    cLogInUser = cLogUser
    cObjectName = cObjName
    
    UserAccountPermissions sConnection      'get first users own permissions
    GetGroupAccounts sConnection            'then the users group
    GetPermissions = True
    If canAccessAll Or cLogInUser = "Admin" Then
        PermissionValue True    'set all permissions to .t. if security is disabled
    End If
    If Not canOpen Then
        MsgBox "You don't have necessary permission to open this object." & vbCrLf & _
               "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
    End If
End Function




Public Function InitPermissions(ByVal sConnection As String, ByVal cObjName As String, _
ByVal cLogUser As String, Optional ByVal cObjDesc As String, _
Optional ByVal cObjType As String) As Boolean

    PermissionValue False       'set all permission to .f.

    cLogInUser = cLogUser
    cObjectName = cObjName
    cObjectDesc = cObjDesc
    cObjectType = cObjType
    If cObjDesc = "" Then
        ' use for compatibility from the old system
        If Not GetObjectDescription(sConnection) Then
            Exit Function
        End If
    Else
        If Not SaveSysObject(sConnection) Then
            Exit Function
        End If
    End If
    
    UserAccountPermissions sConnection      'get first users own permissions
    GetGroupAccounts sConnection            'then the users group
    InitPermissions = True
            
    If canAccessAll Or cLogInUser = "Admin" Then
        PermissionValue True    'set all permissions to .t. if security is disabled
    End If
    If Not canOpen Then
        MsgBox "You don't have necessary permission to open this object." & vbCrLf & _
               "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
    End If
End Function

Public Function SetPermit(cAction As String) As Boolean
    If canAccessAll Or cLogInUser = "Admin" Then
        SetPermit = True
        Exit Function
    End If
    
    cAction = UCase(cAction)
    Select Case cAction
        Case "ADMINISTER"
            If Not canAdminister Then
                MsgBox "You don't have necessary permission to administer." & vbCrLf & _
                "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
                Exit Function
            End If
        Case "ADD"
            If Not canAdd Then
                MsgBox "You don't have necessary permission to add record(s)." & vbCrLf & _
                "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
                Exit Function
            End If
        Case "READ"
            If Not canRead Then
                MsgBox "You don't have necessary permission to read record(s)." & vbCrLf & _
                "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
                Exit Function
            End If
        Case "DUPLICATE"
            If Not canAdd Then
                MsgBox "You don't have necessary permission to add/duplicate record(s)." & vbCrLf & _
                "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
                Exit Function
            End If
        Case "EDIT"
            If Not canEdit Then
                MsgBox "You don't have necessary permission to modify record(s)." & vbCrLf & _
                "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
                Exit Function
            End If
        Case "DELETE"
            If Not canDel Then
                MsgBox "You don't have necessary permission to delete record(s)." & vbCrLf & _
                "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
                Exit Function
            End If
        Case "CANCEL"
            If Not canCancel Then
                MsgBox "You don't have necessary permission to cancel record(s)." & vbCrLf & _
                "Have your system administrator or the person who created this object establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectDesc & "]"
                Exit Function
            End If
    End Select
    SetPermit = True
End Function

Private Function GetObjectDescription(ByVal sConnection As String) As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    Dim lFound As Boolean
    
    Set rs = ObjectDescription(sConnection, cObjectName)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        cObjectDesc = rs("Description").Value
        cObjectType = rs("ObjType").Value
        lFound = True
    Else
        lFound = False
    End If
    rs.Close
    Set rs = Nothing
    GetObjectDescription = lFound
    If Not lFound Then
        MsgBox "Not a valid Object [" & cObjectName & "]" & "." & vbCrLf & _
               "Have your system administrator create this object and establish the appropriate permissions for you.", vbExclamation, "Access Denied   [" & cObjectName & "]"
    End If
End Function

Private Sub PermissionValue(lEnabled As Boolean)
    canOpen = lEnabled
    canAdminister = lEnabled
    canRead = lEnabled
    canAdd = lEnabled
    canEdit = lEnabled
    canDel = lEnabled
    canCancel = lEnabled
End Sub

Private Sub UserAccountPermissions(ByVal sConnection As String)
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    Set rs = UserPermissions(sConnection, cObjectName)
    If Not (rs.BOF And rs.EOF) Then
        canOpen = IIf(rs("canOpen"), True, canOpen)
        canAdminister = IIf(rs("canAdminister"), True, canAdminister)
        canRead = IIf(rs("canRead"), True, canRead)
        canAdd = IIf(rs("canAdd"), True, canAdd)
        canEdit = IIf(rs("canEdit"), True, canEdit)
        canDel = IIf(rs("canDelete"), True, canDel)
        canCancel = IIf(rs("canCancel"), True, canCancel)
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetGroupAccounts(ByVal sConnection As String)
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    Set rs = GroupAccounts(sConnection)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
        Do While Not rs.EOF
            GroupAccountPermissions sConnection, rs("GrpCode").Value
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GroupAccountPermissions(ByVal sConnection As String, ByVal cGroupName As String)
    Dim sSQL As String
    Dim rs As ADODB.Recordset
      
    Set rs = GroupPermissions(sConnection, cGroupName, cObjectName)
    If Not (rs.BOF And rs.EOF) Then
        canOpen = IIf(rs("canOpen"), True, canOpen)
        canAdminister = IIf(rs("canAdminister"), True, canAdminister)
        canRead = IIf(rs("canRead"), True, canRead)
        canAdd = IIf(rs("canAdd"), True, canAdd)
        canEdit = IIf(rs("canEdit"), True, canEdit)
        canDel = IIf(rs("canDelete"), True, canDel)
        canCancel = IIf(rs("canCancel"), True, canCancel)
    End If
    rs.Close
    Set rs = Nothing
End Sub


Private Function GroupPermissions(ByVal sConnection As String, cGroupName As String, cObjectName As String) As ADODB.Recordset
On Error GoTo HandleError
    Dim oData As New RemantecDataAccess
    Dim sSQL As String
    
    sSQL = "SELECT sysPermission.* " & _
           "FROM sysObject INNER JOIN sysPermission ON sysObject.ID = sysPermission.ObjID " & _
           "WHERE (((sysPermission.WrkGrpCode)='" & cGroupName & "') AND ((sysObject.Name)='" & cObjectName & "')) "
   'Calls the stored procedure or SQL statement to get a list of records.
    Set GroupPermissions = oData.GetResult(sConnection, sSQL, 0)
    Set oData = Nothing
    Exit Function
HandleError:
    Err.Raise Err.Number
End Function

Private Function GroupAccounts(ByVal sConnection As String) As ADODB.Recordset
On Error GoTo HandleError
    Dim oData As New RemantecDataAccess
    Dim sSQL As String
    
    sSQL = "SELECT sysGrpMembership.* " & _
           "From sysGrpMembership " & _
           "WHERE (((sysGrpMembership.UsrCode)='" & cLogInUser & "')) "
   'Calls the stored procedure or SQL statement to get a list of records.
    Set GroupAccounts = oData.GetResult(sConnection, sSQL, 0)
    Set oData = Nothing
    Exit Function
HandleError:
    Err.Raise Err.Number
End Function

Private Function UserPermissions(ByVal sConnection As String, cObjectName As String) As ADODB.Recordset
On Error GoTo HandleError
    Dim oData As New RemantecDataAccess
    Dim sSQL As String
    
    sSQL = "SELECT sysWorkGroup.Code AS UserCode, sysObject.Name, sysPermission.canOpen, sysPermission.canAdminister, sysPermission.canRead, sysPermission.canAdd, sysPermission.canEdit, sysPermission.canDelete, sysPermission.canCancel " & _
           "FROM sysWorkGroup INNER JOIN (sysObject INNER JOIN sysPermission ON sysObject.ID = sysPermission.ObjID) ON sysWorkGroup.Code = sysPermission.WrkGrpCode " & _
           "WHERE (((sysWorkGroup.Code)='" & cLogInUser & "') AND ((sysObject.Name)='" & cObjectName & "')) "
   'Calls the stored procedure or SQL statement to get a list of records.
    Set UserPermissions = oData.GetResult(sConnection, sSQL, 0)
    Set oData = Nothing
    Exit Function
HandleError:
    Err.Raise Err.Number
End Function

Private Function ObjectDescription(ByVal sConnection As String, cObjectName As String) As ADODB.Recordset
On Error GoTo HandleError
    Dim oData As New RemantecDataAccess
    Dim sSQL As String

    sSQL = "SELECT sysObject.Description, sysObject.ObjType From sysObject " & _
           "WHERE (((sysObject.Name)='" & cObjectName & "')) "
   'Calls the stored procedure or SQL statement to get a list of records.
    Set ObjectDescription = oData.GetResult(sConnection, sSQL, 0)
    Set oData = Nothing
    Exit Function
HandleError:
    Err.Raise Err.Number
End Function

Private Function SaveSysObject(ByVal sConnection As String) As Boolean
On Error GoTo HandleError
    Dim oConn As ADODB.Connection
    Dim oData As New RemantecDataAccess
    
    If oData.ConnectOpen(sConnection, oConn) Then
        If SaveSysObjectNow(oConn) Then
            oData.ConnectClose oConn, True
            SaveSysObject = True
            Exit Function
        End If
    End If
HandleError:
    oData.ConnectClose oConn, False
End Function

Private Function SaveSysObjectNow(ByRef oConn As ADODB.Connection) As Boolean
On Error GoTo HandleError
    Dim oData As New RemantecDataAccess
    Dim rs As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * From sysObject " & _
           "WHERE (((sysObject.Name)='" & cObjectName & "')) "
    If oData.RecordWrite(sSQL, rs, oConn) Then
        If rs.EOF Then
            rs.AddNew
            rs("ID") = CreateDTID()
            rs("Name") = UCase(cObjectName)
        Else
            rs.MoveFirst
        End If
        rs("Description") = cObjectDesc
        rs("ObjType") = cObjectType
        rs.Update
        rs.Close
        Set rs = Nothing
        SaveSysObjectNow = True
    End If
    Exit Function
HandleError:
    SaveSysObjectNow = False
End Function

