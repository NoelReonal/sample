VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecRptSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public sConnection As String
'Public mCompanyName As String
'Public mSysName As String
Public mReportTitle As String
Public mReportRange As String
Public mSQL_Main As String
Public mRs_Main As ADODB.Recordset
Public mSubRpt_rs1 As ADODB.Recordset, mSubRpt_rs2 As ADODB.Recordset, mSubRpt_rs3 As ADODB.Recordset
Public mSubRpt_SQL1 As String, mSubRpt_SQL2 As String, mSubRpt_SQL3 As String
Public mSubRpt_Name1 As String, mSubRpt_Name2 As String, mSubRpt_Name3 As String

Private oData As RemantecDataAccess
Private oConn As ADODB.Connection

Private mViewer As frmCRViewer
 
Private Sub Class_Initialize()
    Set mViewer = New frmCRViewer
    Set oData = New RemantecDataAccess
End Sub

Private Sub Class_Terminate()
    Set mViewer = Nothing
    Set oData = Nothing
End Sub

Public Sub Show(ByRef Rpt As CRAXDRT.Report, Optional lPrintOnly As Boolean)
    If oData.ConnectOpen(sConnection, oConn) Then
        ADO_Connect Rpt
        Rpt.ReportAuthor = UCase(pCompanyName)
        Rpt.ReportTitle = mReportTitle
        Rpt.ReportProgName.SetText pSysName
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
        oData.ConnectClose oConn, True
    End If
End Sub


Private Sub ADO_Connect(ByRef mReport As CRAXDRT.Report)

    Dim mADODBCon As ADODB.Connection
    Dim datMain As ADODB.Command
    Dim datSubRpt1 As ADODB.Command
    Dim datSubRpt2 As ADODB.Command
    Dim datSubRpt3 As ADODB.Command
    
    Dim strCnn As String
    
    Dim CRXTables As CRAXDRT.DatabaseTables
    Dim CRXTable As CRAXDRT.DatabaseTable
    Dim CRXSections As CRAXDRT.Sections
    Dim CRXSection As CRAXDRT.Section
    Dim CRXSubreportObj As CRAXDRT.SubreportObject
    Dim CRXReportObjects As CRAXDRT.ReportObjects
    Dim CRXSubreport As CRAXDRT.Report
    Dim CRXReportObject As Object
    
    ' Open the data connection
    Set mADODBCon = oConn
    
    If mSQL_Main = "" Then
        mReport.Database.SetDataSource mRs_Main
    Else
        ' Create a new instance of an ADO command object
        Set datMain = New ADODB.Command
        Set datMain.ActiveConnection = mADODBCon
        datMain.CommandText = mSQL_Main
        mReport.Database.SetDataSource datMain
    End If
        
'    Set datSubRpt1 = New ADODB.Command
'    Set datSubRpt1.ActiveConnection = mADODBCon
'    datSubRpt1.CommandText = mSubRpt_SQL1
    
'    Set datSubRpt2 = New ADODB.Command
'    Set datSubRpt2.ActiveConnection = mADODBCon
'    datSubRpt2.CommandText = mSubRpt_SQL2
        
'    Set datSubRpt3 = New ADODB.Command
'    Set datSubRpt3.ActiveConnection = mADODBCon
'    datSubRpt3.CommandText = mSubRpt_SQL3
        
    Set CRXSections = mReport.Sections 'You start by getting the sections from the Main report
    For Each CRXSection In CRXSections 'You begin by cycling through each section in the main report
        Set CRXReportObjects = CRXSection.ReportObjects 'In each section, you get all the objects in the section
        For Each CRXReportObject In CRXReportObjects 'You cycle through the objects
            If CRXReportObject.Kind = crSubreportObject Then 'You test the objects to see if they’re subreports
                Set CRXSubreportObj = CRXReportObject 'When you find a subreport, you get a hold of it
                Set CRXSubreport = CRXSubreportObj.OpenSubreport 'Finally, you open the subreport and treat it as you would any other report
                
                Set CRXTables = CRXSubreport.Database.Tables
                Set CRXTable = CRXTables.Item(1)
                If CRXSubreportObj.Name = mSubRpt_Name1 Then
                    If mSubRpt_SQL1 <> "" Then
                        Set datSubRpt1 = New ADODB.Command
                        Set datSubRpt1.ActiveConnection = mADODBCon
                        datSubRpt1.CommandText = mSubRpt_SQL1
                        CRXTable.SetDataSource datSubRpt1
                    Else
                        CRXTable.SetDataSource mSubRpt_rs1
                    End If
                ElseIf CRXSubreportObj.Name = mSubRpt_Name2 Then
                    If mSubRpt_SQL2 <> "" Then
                        Set datSubRpt2 = New ADODB.Command
                        Set datSubRpt2.ActiveConnection = mADODBCon
                        datSubRpt2.CommandText = mSubRpt_SQL2
                        CRXTable.SetDataSource datSubRpt2
                    Else
                        CRXTable.SetDataSource mSubRpt_rs2
                    End If
                ElseIf CRXSubreportObj.Name = mSubRpt_Name3 Then
                    If mSubRpt_SQL3 <> "" Then
                        Set datSubRpt3 = New ADODB.Command
                        Set datSubRpt3.ActiveConnection = mADODBCon
                        datSubRpt3.CommandText = mSubRpt_SQL3
                        CRXTable.SetDataSource datSubRpt3
                    Else
                        CRXTable.SetDataSource mSubRpt_rs3
                    End If
                End If
            End If
        Next CRXReportObject
    Next CRXSection
End Sub
