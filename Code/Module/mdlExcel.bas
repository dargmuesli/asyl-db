Attribute VB_Name = "mdlExcel"
Option Compare Database
Option Explicit

Private excelApplication As Excel.Application
Private excelWorkbook As Excel.Workbook
Private excelWorksheet As Excel.worksheet

Public Property Get objExcel() As Excel.Application

    ' Get current or new Excel application
    If excelApplication Is Nothing Then
        Set excelApplication = New Excel.Application
    End If
    
    Set objExcel = excelApplication
End Property

Public Property Get objWorkbook(excelfile As String) As Excel.Workbook

    ' Get current or selected Excel file
    If Not IsWorkBookOpen(excelfile) Then
        Set excelWorkbook = objExcel.Workbooks.Open(excelfile)
    Else
        If Not excelWorkbook.Path & "\" & excelWorkbook.Name = excelfile Then
            Set excelWorkbook = objExcel.Workbooks.Open(excelfile)
        End If
    End If
    
    Set objWorkbook = excelWorkbook
End Property

Private Function IsWorkBookOpen(fileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    
    ff = FreeFile()
    
    ' Try to open workbook locking the file
    Open fileName For Input Lock Read As #ff
    Close ff
    
    ErrNo = Err
    
    On Error GoTo 0
    
    ' Return state depending on the success
    Select Case ErrNo
        Case 0:
            IsWorkBookOpen = False
        Case 70:
            IsWorkBookOpen = True
        Case Else:
            Error ErrNo
    End Select
End Function

Public Function getTableData(excelfile As String) As Dictionary
    Dim tableData As New Dictionary
    Dim worksheet As Excel.worksheet
    Dim worksheets As sheets
    Dim worksheetData As New collection
    Dim currentHeader As String
    Dim i As Integer
    
    Set worksheets = objWorkbook(excelfile).worksheets
    
    ' Search in every worksheet
    For Each worksheet In worksheets
        Set worksheetData = New collection
        
        i = 0
        
        ' Save each column header to its worksheet's name
        Do While i >= 0
            i = i + 1
            
            currentHeader = worksheet.Cells(1, i)
            
            If currentHeader = "" Then
                Exit Do
            End If
            
            worksheetData.Add Replace(Replace(currentHeader, vbLf, "_"), ".", "#")
        Loop
        
        tableData.Add worksheet.Name, worksheetData
    Next worksheet
    
    ' Exit Excel to prevent open and unused processes
    excelApplication.Quit
    
    Set getTableData = tableData
End Function

Public Function finishExcel()
    Set excelWorkbook = Nothing
    
    ' Exit Excel to prevent open and unused processes
    If Not excelApplication Is Nothing Then
        excelApplication.Quit
        Set excelApplication = Nothing
    End If
End Function

Public Function ExcelVersion() As String
    Dim objExcel As Excel.Application
    
    On Error Resume Next
    
    Set objExcel = CreateObject("Excel.Application")
    
    ' Return version number, if able to
    If Err.Number = 0 Then
        With objExcel
            ExcelVersion = .Version
        End With
    End If
    
    ' Exit Excel to prevent open and unused processes
    objExcel.Quit
    
    Set objExcel = Nothing
End Function

Public Function GetSpreadsheetTypeExcelConstant(strVersion As String)
    Dim intSpreadsheetType As Integer
    
    Select Case Val(strVersion)
        'Excel 2007
        Case 12
            intSpreadsheetType = 9
            
        'Excel 2003, 2002, 2000,97
        Case 11, 10, 9, 8
            intSpreadsheetType = 8
    End Select
    
    GetSpreadsheetTypeExcelConstant = intSpreadsheetType
End Function

Public Function GetExcelRangeFromLong(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As String
    Dim excelRange As String
    
    ' Build range string
    excelRange = GetExcelChar(x1) & y1 & ":" & GetExcelChar(x2) & y2
    GetExcelRangeFromLong = excelRange
End Function

Public Function GetExcelChar(ByVal lng As Long) As String
    Dim str As String
    Dim lngDiff As Long
    
    ' Get the A-ZZ representation of a number for Excel column addressing
    lngDiff = lng Mod 26
    
    If lngDiff = 0 Then
        lngDiff = 26
        str = "Z"
    Else
        str = Chr((lngDiff Mod 26) + 64)
    End If
    
    lng = (lng - lngDiff) / 26
    
    Do While lng > 0
        lngDiff = lng Mod 26
        
        If lngDiff = 0 Then
            lngDiff = 26
            str = "Z" & str
        Else
            str = Chr((lng Mod 26) + 64) & str
        End If
        
        lng = (lng - lngDiff) / 26
    Loop
    
    GetExcelChar = str
End Function
