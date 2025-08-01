VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecRptDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public sConnection As String

Public mReportTitle As String
Public mSQL As String
Public mRs As ADODB.Recordset

Private oData As RemantecDataAccess
Public mViewer As frmCRViewer
 
Private Sub Class_Initialize()
    Set mViewer = New frmCRViewer
    Set oData = New RemantecDataAccess
End Sub

Private Sub Class_Terminate()
    Set mViewer = Nothing
    Set oData = Nothing
End Sub

Public Sub Show(ByRef Rpt As CRAXDRT.Report, Optional lPrintOnly As Boolean)
    If SetData(Rpt) Then
    
        Rpt.PrinterSetup 0
        If lPrintOnly Then
            Rpt.PrintOut False
        End If
        
       ' Rpt.PaperOrientation = crDefaultPaperOrientation
       ' Rpt.PaperSize = 5.5
  
        mViewer.CRViewer1.DisplayGroupTree = False
        mViewer.CRViewer1.DisplayTabs = False
        mViewer.CRViewer1.EnableSearchControl = False
        mViewer.CRViewer1.EnableSearchExpertButton = False
        mViewer.CRViewer1.EnableSelectExpertButton = False
        'mViewer.CRViewer1.EnableExportButton = False
        mViewer.CRViewer1.EnableGroupTree = False
    
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
    Dim cn As ADODB.Connection
    
    If mSQL <> "" Then
        ' Add the datasource to the report
        Set rs = oData.GetResult(sConnection, mSQL, 0)
        Rpt.Database.SetDataSource rs, 3, 1
        Rpt.ReadRecords
        rs.Close
        Set rs = Nothing
        SetData = True
    ElseIf Not (mRs Is Nothing) Then
        Rpt.Database.SetDataSource mRs, 3, 1
        Rpt.ReadRecords
        SetData = True
    Else
        SetData = True
    End If
End Function

