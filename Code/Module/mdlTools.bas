Attribute VB_Name = "mdlTools"
Option Compare Database
Option Explicit

Function OpenFilename(Optional startDir As String, Optional sTitle As String = "Datei auswählen:", Optional sFilter As String = "Access-DB (*.mdb)|Alle Dateien (*.*)") As String
    Static sDir As String
    
    ' Magic variable
    WizHook.key = 51488399
    
    ' Try to set start directory from given information
    If Len(startDir) = 0 Then
        If Len(sDir) = 0 Then
            startDir = CurrentProject.Path
        Else
            startDir = sDir
        End If
    End If
    
    ' Open file browser dialog
    Call WizHook.GetFileName(Application.hWndAccessApp, "Microsoft Access", sTitle, "Öffnen", OpenFilename, startDir, sFilter, 0&, 0&, &H40, False)
    
    ' Try to return selected path
    If Len(OpenFilename) > 0 Then
        sDir = Left(OpenFilename, InStrRev(OpenFilename, "\", , vbTextCompare))
    End If
End Function

Public Function FileFolderExists(strFullPath As String) As Boolean
    On Error GoTo EarlyExit
    
    ' Check existence indicators
    If Not Dir(strFullPath, vbDirectory) = vbNullString And Not strFullPath = "" Then
        FileFolderExists = True
    End If
    
EarlyExit:
    On Error GoTo 0
End Function

Public Function IsFormOpen(strForm As String) As Boolean

    ' Return the form opening state
    IsFormOpen = SysCmd(acSysCmdGetObjectState, acForm, strForm) > 0
End Function

Function IsInArray(searchString As String, arr As Variant) As Boolean

    ' Return the existence state
    IsInArray = UBound(Filter(arr, searchString)) > -1
End Function

Function GetArrayItemIndex(arr, v) As Integer
    Dim lb As Long
    Dim ub As Long
    Dim i As Long
    
    lb = LBound(arr)
    ub = UBound(arr)
    
    ' Check for every item in array if it's the searched value
    For i = lb To ub
        If arr(i) Like v Then
            Exit For
        End If
    Next i
    
    ' Return found index or -1
    If i > ub Then
        GetArrayItemIndex = -1
    Else
        GetArrayItemIndex = i
    End If
End Function

Function GetCollectionItemIndex(col As collection, v As Variant) As Integer
    Dim c As Variant
    Dim i As Integer
    
    i = 0
    
    ' Check for every item in array if it's the searched value
    For Each c In col
        If c Like v Then
            Exit For
        End If
        
        i = i + 1
    Next c
    
    ' Return found index or -1
    If i = col.Count Then
        GetCollectionItemIndex = -1
    Else
        GetCollectionItemIndex = i
    End If
End Function

Public Function dhAge(dtmBD As Date, Optional dtmDate As Date = 0) As Integer
    Dim intAge As Integer
    
    If dtmDate = 0 Then
        dtmDate = Date
    End If
    
    ' Calculate the year difference
    intAge = DateDiff("yyyy", dtmBD, dtmDate)
    
    ' Compensate for differences in days
    If dtmDate < DateSerial(Year(dtmDate), Month(dtmBD), Day(dtmBD)) Then
        intAge = intAge - 1
    End If
    
    dhAge = intAge
End Function

Public Function DLookUpMulti(expression As String, domain As String, Optional criteria As String) As String
    Dim result As String
    Dim criterias() As String
    Dim i As Long
    
    ' Parse input
    criterias() = Split(criteria, "; ")
    
    ' Add multiple DLookups together
    For i = 0 To UBound(criterias)
        If i > 0 Then
            result = result & ", "
        End If
        
        result = result & DLookup(expression, domain, "ID = " & criterias(i))
    Next
    
    DLookUpMulti = result
End Function

Public Function CollectionToString(collection As collection) As String
    Dim item As Variant
    Dim output As String
    
    ' Generate output
    For Each item In collection
        output = output + item + ";"
    Next item
    
    CollectionToString = output
End Function

Public Function collectionsEqual(col1 As collection, col2 As collection)
    Dim equalCollections As Boolean
    Dim col1Item As Variant
    Dim i As Integer
    
    equalCollections = True
    i = 0
    
    ' Pre-check the collections' sizes
    If col1.Count = col2.Count Then
        For Each col1Item In col1
            i = i + 1
            
            ' Compare the items
            If Not col1Item = col2.item(i) Then
                equalCollections = False
                
                Exit For
            End If
        Next
    Else
        equalCollections = False
    End If
    
    collectionsEqual = equalCollections
End Function

Public Function HasParent(p) As Boolean
    On Error GoTo handler
    
    HasParent = TypeName(p.Parent.Name) = "String"
    
    Exit Function
handler:
End Function
