VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecLicenseKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mvarMaxRecord As Integer
'Public mvarSectionName As String

Private Const APP_NAME = "REMANTEC"
'Private Const SECTION_NAME = "RPOS-TOUCH"
Private Const LICENSE_KEY = "ID"
Private Const REGISTRATION_KEY = "REG"
Private Const UPDATE_KEY = "UPD"
Private Const NODAYS_KEY = "ND"
Private Const VALIDATION_KEY = "VLD"
Private Const VALID_CHARACTERS = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ"
Private Const DEFAULT_FORMAT = "&&&&&&-&&&&&&-&&&&&&-&&&&&&"
Private Const MAX_FILENAME_LEN = 256
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)

Private Declare Function GetComputerName Lib "kernel32" Alias _
    "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long


'Public Function GetSerialNumber() As String
'    Dim ser As String
'    Dim o As New Remantec_SerialNumber
'
'    ser = o.getSerialNo()
'    GetSerialNumber = ser
'    'Debug.Print GetSerialNumber
'End Function


'  Returns a volume's serial number
'
Public Function GetSerialNumber() As String
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim i As Long
   Dim j As Long
   Dim sDrive As String

   sDrive = "C"
   Call GetVolumeInformation(sDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, i, j, s2, MAX_FILENAME_LEN)
   GetSerialNumber = Abs(ser)
End Function


'  Returns a volume's license key
'
Public Function GetLicenseKey(ByVal ApplicationKey As String, ByVal cSerialNo As String) As String
    Dim cComputerID As String
    If mvarMaxRecord > 0 Then
        cComputerID = cSerialNo & Trim(Str(mvarMaxRecord))
    Else
        cComputerID = cSerialNo
    End If
    GetLicenseKey = CreateKey(ApplicationKey, cComputerID)
End Function


Public Function SaveLicenseKey(ByVal ApplicationKey As String, ByVal cLicenseKey As String) As Boolean
    
    SaveLicenseKey = False
    If cLicenseKey = "000000-000000-000000-000000" Then
        SaveSettingLicenseKey ApplicationKey, cLicenseKey
        SaveLicenseKey = True
    Else
        If cLicenseKey = GetLicenseKey(ApplicationKey, GetSerialNumber()) Then
            SaveSettingLicenseKey ApplicationKey, cLicenseKey
            SaveLicenseKey = True
        End If
    End If
    
    If Not SaveLicenseKey Then
        MsgBox "Please register a valid license key." & vbCrLf & _
               "Contact your local distributor for a registered license key.", vbCritical, "INVALID KEY"
    End If

End Function


Public Function ReadLicenseKey(ByVal ApplicationKey As String) As String
    Dim values As Variant

    values = GetAllSettings(APP_NAME, ApplicationKey)
    
    If IsEmpty(values) Then
        ReadLicenseKey = "000000-000000-000000-000000"
    Else
        ReadLicenseKey = values(0, 1)
    End If
End Function


Public Sub DeleteLicenseKey(ByVal ApplicationKey As String)
On Error Resume Next
    DeleteSetting APP_NAME, ApplicationKey
End Sub


Public Function IsLicenseCopy(ByVal ApplicationKey As String, Optional ByVal nMaxDays As Integer) As Boolean
    Dim cLicenseKey As String
    Dim values As Variant
    Dim nDaysRemaining As Currency
    
   
    IsLicenseCopy = True
    values = GetAllSettings(APP_NAME, ApplicationKey)
    If IsEmpty(values) Then
        ' Set a demo copy only
        SaveLicenseKey ApplicationKey, "000000-000000-000000-000000"
        IsLicenseCopy = ValidDemo(ApplicationKey, nMaxDays)
    Else
        cLicenseKey = values(0, 1)
        If cLicenseKey = "000000-000000-000000-000000" Then
            If pVar.ExpiryDate <> 0 Then
                nDaysRemaining = pVar.ExpiryDate - pVar.CurrentDate
                If nDaysRemaining < 0 Then
                    nDaysRemaining = 0
                End If
                If (nDaysRemaining = 0) Then
                    MsgBox "Demo copy has expired." & vbCrLf & _
                       "Contact your local distributor for a registered license key.", vbCritical, "EXPIRED KEY"
                    IsLicenseCopy = False
                ElseIf (nDaysRemaining <= 30) Then
                    MsgBox "Demo copy will expire in " & Str(nDaysRemaining) & " day(s)." & vbCrLf & _
                       "Contact your local distributor for a registered license key.", vbCritical, "WARNING"
                    IsLicenseCopy = True
                End If
            Else
                IsLicenseCopy = ValidDemo(ApplicationKey, nMaxDays)
            End If
        ElseIf cLicenseKey <> GetLicenseKey(ApplicationKey, GetSerialNumber()) Then
            MsgBox "Please register a valid license key." & vbCrLf & _
                   "Contact your local distributor for a registered license key.", vbCritical, "INVALID KEY"
            IsLicenseCopy = False
        End If
    End If
End Function


Private Sub SaveSettingLicenseKey(ByVal ApplicationKey As String, ByVal cLicenseKey As String)
    Dim values As Variant
    Dim cRegKey As String
    Dim cCurDate As String
    Dim cNoDays As String
    Dim cValidKey As String
    
    values = GetAllSettings(APP_NAME, ApplicationKey)
    
    If IsEmpty(values) Then
        cLicenseKey = "000000-000000-000000-000000"
        cRegKey = Format(Now, "YYYYMMDDHHSS") & GetSerialNumber()
        cCurDate = Format(Date, "YYYYMMDD")
        cNoDays = "000"
        cValidKey = cCurDate & cNoDays
        cCurDate = Encrypt(cCurDate, True)
        cNoDays = Encrypt(cNoDays, True)
        SaveSetting APP_NAME, ApplicationKey, LICENSE_KEY, cLicenseKey
        SaveSetting APP_NAME, ApplicationKey, REGISTRATION_KEY, cRegKey
        SaveSetting APP_NAME, ApplicationKey, UPDATE_KEY, cCurDate
        SaveSetting APP_NAME, ApplicationKey, NODAYS_KEY, cNoDays
        SaveSetting APP_NAME, ApplicationKey, VALIDATION_KEY, cValidKey
    Else
        SaveSetting APP_NAME, ApplicationKey, LICENSE_KEY, cLicenseKey
    End If
End Sub

Private Function ValidDemo(ByVal ApplicationKey As String, nMaxDays As Integer) As Boolean
    Dim values As Variant
    Dim cCurDate As String
    Dim cNoDays As String
    Dim cValidKey As String
    Dim cNewDate As String
    Dim nDays As Integer
    Dim nDaysRemaining As Integer
    
    values = GetAllSettings(APP_NAME, ApplicationKey)
    If Not IsEmpty(values) Then
        cCurDate = values(2, 1)
        cNoDays = values(3, 1)
        cValidKey = values(4, 1)
        cCurDate = Encrypt(cCurDate, False)
        cNoDays = Encrypt(cNoDays, False)
        
        
        'MsgBox "Current Date : " & cCurDate & vbCrLf & _
               "No Days : " & cNoDays & vbCrLf & _
               "Valid Key 1 : " & mID$(cValidKey, 1, 8) & vbCrLf & _
                "Valid Key 2 : " & mID$(cValidKey, 9, 7)

        If cCurDate = mID$(cValidKey, 1, 8) And cNoDays = mID$(cValidKey, 9, 7) Then
            
            nDays = CInt(cNoDays)
            'nDays = GetTranDays(nDays)
            
            If nDays >= nMaxDays Then
                MsgBox "Demo copy has expired." & vbCrLf & _
                       "Contact your local distributor for a registered license key.", vbCritical, "EXPIRED KEY"
                ValidDemo = False
            Else
                cNewDate = Format(Date, "YYYYMMDD")
                If cCurDate <> cNewDate Then
                    cNoDays = Format(CInt(cNoDays) + 1, "0000000")
                    cCurDate = cNewDate
                    cValidKey = cCurDate & cNoDays
                    cCurDate = Encrypt(cCurDate, True)
                    cNoDays = Encrypt(cNoDays, True)
                    SaveSetting APP_NAME, ApplicationKey, UPDATE_KEY, cCurDate
                    SaveSetting APP_NAME, ApplicationKey, NODAYS_KEY, cNoDays
                    SaveSetting APP_NAME, ApplicationKey, VALIDATION_KEY, cValidKey
                End If
                ValidDemo = True
            
                nDaysRemaining = nMaxDays - nDays
                If (nDaysRemaining <= 30) Then
                    MsgBox "Demo copy will expire in " & Str(nDaysRemaining) & " day(s)." & vbCrLf & _
                       "Contact your local distributor for a registered license key.", vbCritical, "WARNING"
                End If
            
            End If
        Else
            MsgBox "Please register a valid license key." & vbCrLf & _
                   "Contact your local distributor for a registered license key.", vbCritical, "TAMPERED KEY"
            ValidDemo = False
        End If
    End If
End Function


Private Function CreateKey(ApplicationKey As String, UserName As String, Optional sFormat As String = DEFAULT_FORMAT, Optional ValidCharacters As String = VALID_CHARACTERS) As String
    ' for use in sFormat; use '&' to represent alpha-numeric characters
    Dim intTemp As Integer
    Dim strTextChar As String
    Dim strKeyChar As String
    Dim intEncryptedChar As String
    Dim strKey As String
    Dim i As Integer
    Dim strUserName As String
    
    ApplicationKey = UCase(ApplicationKey)
    strUserName = LCase(Trim(UserName))
    
    If Len(strUserName) = 0 Then
        Err.Raise vbError + 1001, , "Invalid Username"
        Exit Function
    End If
    
    'This is an altered simple encryption algorithm
    For i = 1 To CountAmpersands(sFormat)
        strTextChar = mID(strUserName, (i Mod Len(strUserName)) + 1, 1)
        strKeyChar = mID(ApplicationKey, (i Mod Len(ApplicationKey)) + 1, 1)
        intTemp = (((Asc(strKeyChar) * i) * Len(ApplicationKey) + 1) Mod Len(ValidCharacters) + 1)
        strTextChar = Chr(Asc(strTextChar) Xor intTemp)
        intTemp = (((Asc(strKeyChar) * i) * Len(UserName) + 1) Mod Len(ValidCharacters) + 1)
        strTextChar = Chr(Asc(strTextChar) Xor intTemp)
        intEncryptedChar = ((Asc(strTextChar) Xor Asc(strKeyChar)) Mod Len(ValidCharacters)) + 1
        strKey = strKey & mID(ValidCharacters, intEncryptedChar, 1)
    Next i
    
    CreateKey = Format(strKey, sFormat)
End Function

Private Function CountAmpersands(ByVal Format As String) As Integer
    'Counts the number of characters that need to be returned
    
    Dim i As Integer
    Dim intCount As Integer
    
    intCount = 0
    For i = 1 To Len(Format)
        If mID(Format, i, 1) = "&" Then
            intCount = intCount + 1
        End If
    Next i
    
    CountAmpersands = intCount
End Function

Private Function Encrypt(cString As String, lFlag As Boolean) As String
    Dim i As Long, n As Long, nStrLen As String, cResult As String, cValue As String
    Dim j As Long

    ' max len = 9
    nStrLen = Len(cString)
    cResult = ""

    For i = 1 To nStrLen
        n = Asc(mID$(cString, i, 1))
        If lFlag Then
            j = n + (i * i) + 101
            cValue = Chr$(j)
        Else
            j = n - (i * i) - 101
            cValue = Chr$(j)
        End If
        cResult = cResult & cValue
    Next

    Encrypt = cResult
End Function


Public Function DaysRemaining(ByVal ApplicationKey As String, nMaxDays As Integer) As Integer
    Dim values As Variant
    Dim cNoDays As String
    Dim o As New dbSettings
    
    DaysRemaining = 0
    values = o.GetAllSettings(APP_NAME, ApplicationKey)
    If Not IsEmpty(values) Then
        cNoDays = values(3, 1)
        cNoDays = Encrypt(cNoDays, False)
        DaysRemaining = nMaxDays - CInt(cNoDays)
    End If
End Function

Public Function GetTerminalName() As String
    Dim buffer As String * 512, length As Long
    length = Len(buffer)
    If GetComputerName(buffer, length) Then
        GetTerminalName = UCase(Left$(buffer, length))
    End If
End Function

Private Function GetTranDays(nDays As Integer) As Integer
    Dim oData As New RemantecDataAccess
    Dim rs As ADODB.Recordset
    
    If pVar.CurrentDate = 0 Then
        GetTranDays = nDays
        Exit Function
    End If
    
    Set rs = oData.GetResult(pVar.ConnectPos, "SELECT TRANSDATE FROM POS_TRN_H ORDER BY TRANSDATE DESC", 0)
    If Not rs.EOF Then
        dDate = rs("TransDate")
        GetTranDays = pVar.CurrentDate - rs("TransDate")
    Else
        GetTranDays = nDays
    End If
    rs.Close
    Set rs = Nothing
End Function
