VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RemantecDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function IsLeapYear(intYear As Integer) As Boolean
    '
    ' Check if a year is a leap year.
    '
    IsLeapYear = False
    
    If intYear Mod 4 = 0 Then
        If intYear Mod 100 = 0 Then
            If intYear Mod 400 = 0 Then
                IsLeapYear = True
            End If
        Else
            IsLeapYear = True
        End If
    End If

End Function

Public Function MonthEnd(dDate As Date, _
                blnSkipCurMthEnd As Boolean) As Date

'
' Convert a date in "yyyymmdd" format
' to the closest month end date.
'
Dim strYear    As String
Dim strMth     As String
Dim strDay     As String
Dim strNextDay As String
Dim strDate    As String
Dim strMonthEnd  As String
Dim dNextDay As Date

    strDate = Format$(dDate, "yyyymmdd")
    strYear = mID$(strDate, 1, 4)
    strMth = mID$(strDate, 5, 2)
    
    ' Get the current month end date.
    '
    Select Case strMth
        Case "04", "06", "09", "11"
            strDay = "30"
            
        Case "02"
            If IsLeapYear(Val(strYear)) Then
                strDay = "29"
            Else
                strDay = "28"
            End If
        
        Case Else
            strDay = "31"
    End Select
    strMonthEnd = strYear & strMth & strDay
    MonthEnd = Format(strMonthEnd, "@@@@-@@-@@")
    
    '
    ' If the date is the month end, and the "skip current
    ' month end" flag is set, get the next month end date.
    '
    If (strMonthEnd = strDate) And blnSkipCurMthEnd Then
        dNextDay = NextDay(dDate)
        MonthEnd = MonthEnd(dNextDay, False)
    End If
    
End Function



Public Function NextDay(dDate As Date) As Date
'Dim dteDate As Date
    '
    ' Convert a date in "yyyymmdd" format to
    ' the following calendar day.
    '
    'dteDate = Format(strDate, "@@@@-@@-@@")
    'NextDay = Format$(DateAdd("d", 1, dteDate), "yyyymmdd")
    
    NextDay = DateAdd("d", 1, dDate)
End Function



Public Function QuarterEnd(dDate As Date, _
            blnSkipCurQtrEnd As Boolean) As Date
'
' Convert a date in "yyyymmdd" format to
' the closest quarter end date.
'
Dim strYear    As String
Dim strMth     As String
Dim strMthEnd  As String
Dim strNextDay As String
Dim strDate    As String
Dim strQuarterEnd As String
Dim dNextDay As Date
    
    strDate = Format$(dDate, "yyyymmdd")
    strYear = mID$(strDate, 1, 4)
    strMth = mID$(strDate, 5, 2)
    '
    ' Get the current quarter end date.
    '
    Select Case strMth
        Case "01", "02", "03"
            strMthEnd = "0331"
            
        Case "04", "05", "06"
            strMthEnd = "0630"
            
        Case "07", "08", "09"
            strMthEnd = "0930"
            
        Case "10", "11", "12"
            strMthEnd = "1231"
    End Select
    strQuarterEnd = strYear & strMthEnd
    QuarterEnd = Format(strQuarterEnd, "@@@@-@@-@@")
    '
    ' If the date is the quarter end, and the "skip current
    ' quarter end" flag is set, get the next quarter end date.
    '
    If (strQuarterEnd = strDate) And blnSkipCurQtrEnd Then
        dNextDay = NextDay(dDate)
        QuarterEnd = QuarterEnd(dNextDay, False)
    End If
    
End Function


Public Function YearAdd(dDate As Date, nYear As Integer) As Date
    Dim nDay As Integer
    Dim nYearNew As Integer
    
    If IsLeapYear(Year(dDate)) And Month(dDate) = 2 Then
        nDay = 28
    Else
        nDay = Day(dDate)
    End If
    nYearNew = Year(dDate) + nYear
    YearAdd = DateSerial(nYearNew, Month(dDate), nDay)
End Function


Public Function NextMonth(dDate As Date, Optional nOrigDay As Integer) As Date
    Dim nDay As Integer
    Dim nYear As Integer
    Dim nMonth As Integer
    Dim nNextMonth As Integer
    
    If nOrigDay > 0 Then
        nDay = nOrigDay
    Else
        nDay = Day(dDate)
    End If
    nYear = Year(dDate)
    nMonth = Month(dDate)
    If nMonth = 12 Then
        nYear = nYear + 1
        nNextMonth = 1
    Else
        nNextMonth = nMonth + 1
    End If
    
    Select Case nNextMonth
        Case 4, 6, 9, 11
            If nDay >= 30 Then
                nDay = 30
            End If
        Case 2
            If IsLeapYear(nYear) Then
                If nDay >= 29 Then
                    nDay = 29
                End If
            Else
                If nDay >= 28 Then
                    nDay = 28
                End If
            End If
     End Select
    
    NextMonth = DateSerial(nYear, nNextMonth, nDay)
End Function


Public Function MonthAdd(dDate As Date, nMonth As Integer) As Date
    Dim i As Integer
    Dim dNewDate As Date
    Dim nDay As Integer
    
    nDay = Day(dDate)
    dNewDate = dDate
    For i = 1 To nMonth
        dNewDate = NextMonth(dNewDate, nDay)
    Next
    MonthAdd = dNewDate
End Function


