VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecLicenseKeyReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const INI_FILE = "SETUP.INI"

Public IsAdmin As Boolean

Dim mvarCompanyName As String
Dim mvarAddress As String
Dim mvarPhoneNo As String
Dim mvarFaxNo As String
Dim mvarContactName As String
Dim mvarContactTitle As String
Dim mvarSerialNo As String

'Dim mvarLicenseKey As String

Dim clsLicense As RemantecLicenseKey

'Private Function ReadIni(ByVal ApplicationKey As String) As Boolean
Private Function ReadIni() As Boolean
    Dim cFileName As String
    Dim n As Long
    Dim cValue As String
    Dim cResult As String
    
    mvarSerialNo = clsLicense.GetSerialNumber
    'mvarLicenseKey = clsLicense.ReadLicenseKey(ApplicationKey)
    
    cFileName = App.Path & "\" & INI_FILE
    If Dir(cFileName) = "" Then
        'ReadIni = WriteIni(ApplicationKey)
        ReadIni = WriteIni()
        Exit Function
    End If
    
    Open (cFileName) For Input As #1
    n = 1
    Do While n <> 0
        Line Input #1, cValue
        n = InStr(1, cValue, "=")
        If n <> 0 Then
            cResult = mID$(cValue, n + 1)
            Select Case UCase(Left$(cValue, n - 1))
                Case "COMPANY NAME"
                    mvarCompanyName = cResult
                Case "ADDRESS"
                    mvarAddress = cResult
                Case "PHONE NO"
                    mvarPhoneNo = cResult
                Case "FAX NO"
                    mvarFaxNo = cResult
                Case "CONTACT NAME"
                    mvarContactName = cResult
                Case "CONTACT TITLE"
                    mvarContactTitle = cResult
            End Select
        End If
    Loop
    Close #1
    
  
End Function


'Private Function WriteIni(ByVal ApplicationKey As String) As Boolean
Private Function WriteIni() As Boolean
On Error GoTo HandleError
    Dim cFileName As String
    
    'If Not clsLicense.SaveLicenseKey(ApplicationKey, mvarLicenseKey) Then
    '    Exit Function
    'End If
    
    cFileName = App.Path & "\" & INI_FILE
    Open (cFileName) For Output As #1
    Print #1, "COMPANY NAME=" & mvarCompanyName
    Print #1, "ADDRESS=" & mvarAddress
    Print #1, "PHONE NO=" & mvarPhoneNo
    Print #1, "FAX NO=" & mvarFaxNo
    Print #1, "CONTACT NAME=" & mvarContactName
    Print #1, "CONTACT TITLE=" & mvarContactTitle
    Print #1,
    Close #1
    
    WriteIni = True
    Exit Function
HandleError:
    WriteIni = False
End Function

Private Sub Class_Initialize()
    Set clsLicense = New RemantecLicenseKey
End Sub

Private Sub Class_Terminate()
    Set clsLicense = Nothing
End Sub


Public Function RegisterLicenseKey(Optional ByVal ApplicationKey As String, Optional ByVal IsLicense As Boolean) As Boolean
    Dim frm As New frmLicenseKeyRegistration
    
    frm.IsAdmin = IsAdmin
    ReadIni
    If ApplicationKey <> "" Then
        frm.cmbSystem = ApplicationKey
        frm.cmbSystem.Enabled = False
        'ReadIni ApplicationKey
        'frm.ValidLicenseKey = clsLicense.GetLicenseKey(ApplicationKey, mvarSerialNo)
    Else
        frm.cmbSystem.ListIndex = -1
        frm.cmbSystem.Enabled = True
    End If
    
    frm.txtName = mvarCompanyName
    frm.txtAddress = mvarAddress
    frm.txtPhoneNo = mvarPhoneNo
    frm.txtFaxNo = mvarFaxNo
    frm.txtContactName = mvarContactName
    frm.txtContactTitle = mvarContactTitle
    frm.lblSerialNo.Caption = mvarSerialNo
    'frm.mbxLicenseKey = mvarLicenseKey
    If IsLicense Then
        frm.mbxLicenseKey = clsLicense.GetLicenseKey(frm.cmbSystem.Text, mvarSerialNo)
    End If
    frm.Show vbModal
    If frm.SelectOk Then
    
        If Not clsLicense.SaveLicenseKey(frm.cmbSystem.Text, frm.mbxLicenseKey) Then
            Exit Function
        End If
    
        'lblLicenseTo.Caption = "License To: " & frm.txtName
        mvarCompanyName = frm.txtName
        mvarAddress = frm.txtAddress
        mvarPhoneNo = frm.txtPhoneNo
        mvarFaxNo = frm.txtFaxNo
        mvarContactName = frm.txtContactName
        mvarContactTitle = frm.txtContactTitle
        'mvarLicenseKey = frm.mbxLicenseKey
        'If WriteIni(ApplicationKey) Then
        If WriteIni() Then
            RegisterLicenseKey = True
        End If
    End If
End Function


Public Function IsLicenseCopy(ByVal ApplicationKey As String, Optional ByVal nMaxDays As Integer) As Boolean
    Dim frm As New frmLicenseKeyRegistration
    
    Do While True
        IsLicenseCopy = clsLicense.IsLicenseCopy(ApplicationKey, nMaxDays)
        If Not IsLicenseCopy Then
            If Not RegisterLicenseKey(ApplicationKey) Then
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
End Function
