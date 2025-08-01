VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public sConnection As String
'Public mCompanyName As String
'Public mSysName As String
Public mReportTitle As String
Public mReportRange As String
Public mSQL As String
Public mRs As ADODB.Recordset
Public mEnabledGroupTree As Boolean

Private mViewer As frmCRViewer
 
Private Sub Class_Initialize()
    Set mViewer = New frmCRViewer
End Sub

Private Sub Class_Terminate()
    Set mViewer = Nothing
End Sub

Public Sub Show(ByRef Rpt As CRAXDRT.Report, Optional lPrintOnly As Boolean)
    'Dim clsLicense As New RemantecLicenseKeyReg
    
    'If Not clsLicense.IsLicenseCopy("RPOS-TOUCH", 30) Then
    '    Exit Sub
    'End If
    
    If SetData(Rpt) Then
        mViewer.CRViewer1.DisplayGroupTree = mEnabledGroupTree
        mViewer.CRViewer1.EnableGroupTree = mEnabledGroupTree
        Rpt.ReportAuthor = UCase(pCompanyName)
        Rpt.ReportProgName.SetText pSysName
        Rpt.ReportTitle = mReportTitle
        Rpt.ReportComments = mReportRange
        If lPrintOnly Then
            Rpt.PrintOut False
        End If
        mViewer.Caption = mReportTitle
        mViewer.CRViewer1.ReportSource = Rpt
        If lPrintOnly Then
            mViewer.PrintReport
        Else
            mViewer.ViewReport
            mViewer.Show vbModal
        End If
    End If
End Sub

Private Function SetData(ByRef Rpt As CRAXDRT.Report) As Boolean
    Dim rs As ADODB.Recordset
    Dim oData As New RemantecDataAccess
    
    If mSQL <> "" Then
        Set rs = oData.GetResult(sConnection, mSQL, 0)
        Rpt.Database.SetDataSource rs, 3, 1
        Rpt.ReadRecords
        rs.Close
        Set rs = Nothing
        SetData = True
    Else
        Rpt.Database.SetDataSource mRs, 3, 1
        Rpt.ReadRecords
        SetData = True
    End If
End Function
