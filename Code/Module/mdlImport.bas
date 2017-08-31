Attribute VB_Name = "mdlImport"
Option Compare Database
Option Explicit

Public Sub preImportTable(targetTableName As String, fileName As String, sheet As String, headers As collection, headerIndexes As collection)
    Dim db As Database
    Dim cnn As ADODB.Connection
    Dim sql, sqlInsert, sqlSelect, sqlFrom, sqlIn, sqlWhere As String
    Dim headerIndex As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    Set cnn = CurrentProject.Connection
    i = 0
    
    ' Reset autonumber in table
    sql = "INSERT INTO [" & targetTableName & "] ([ID]) VALUES (0)"
    cnn.Execute sql
    
    ' Delete table's rows
    sql = "DELETE FROM [" & targetTableName & "]"
    cnn.Execute sql
    
    ' Static (generated) SQL strings
    sqlSelect = " SELECT "
    
    For Each headerIndex In headerIndexes
        i = i + 1
        
        sqlSelect = sqlSelect & "[" & headers(headerIndex + 1) & "]"
        
        If i < headerIndexes.Count Then
            sqlSelect = sqlSelect & ", "
        End If
    Next
    
    sqlInsert = "INSERT INTO [" & targetTableName & "] (" _
        & "[Person ID], " _
        & "[File Number], " _
        & "[Case Name], " _
        & "[Last Name], " _
        & "[First Name], " _
        & "[Accommodation], " _
        & "[Place], " _
        & "[Birth Date], " _
        & "[Age], " _
        & "[Gender], " _
        & "[Martial Status], " _
        & "[Nationality], " _
        & "[Designation], " _
        & "[People Number], " _
        & "[Address Type], " _
        & "[Point Designation]" _
    & ")"
    sqlIn = " IN '" & fileName & "' [Excel 8.0;HDR=Yes;IMEX=1]"
    sqlFrom = " FROM [" & sheet & "$]"
    
    ' Insert new data into table
    sql = sqlInsert & sqlSelect & sqlFrom & sqlIn
    Debug.Print sql
    cnn.Execute sql
    
    cnn.Close
    Set cnn = Nothing
End Sub

Public Sub importTable(sourceTableName As String)
    Dim db As Database
    Dim cnn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fld As Object
    Dim sql, sqlInsert, sqlSelect As String
    
    Set db = CurrentDb
    Set cnn = CurrentProject.Connection
    
    ' Delete removed entries
    sql = "DELETE FROM [Table Base-Data] " _
        & "WHERE [Table Base-Data].[ID] " _
        & "IN (SELECT ID FROM [Query Import Changes Removed])"
    Debug.Print sql
    cnn.Execute sql

    ' Delete changed entries (above method takes long)
    sql = "SELECT ID FROM [Query Import Changes Changed-ID]"
    Debug.Print sql
    Set rs = cnn.Execute(sql)

    Do Until rs.EOF
        For Each fld In rs.Fields
            sql = "DELETE FROM [Table Base-Data] " _
                & "WHERE [Table Base-Data].[ID] = " & Chr(34) & fld.value & Chr(34)
            Debug.Print sql
            cnn.Execute sql
        Next fld

        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    ' Insert new and changed entries
    sqlInsert = "INSERT INTO [Table Base-Data] (" _
        & "[ID], " _
        & "[File Number], " _
        & "[Case Name], " _
        & "[Last Name], " _
        & "[First Name], " _
        & "[Accommodation], " _
        & "[Place], " _
        & "[Birth Date], " _
        & "[Age], " _
        & "[Gender], " _
        & "[Martial Status], " _
        & "[Nationality], " _
        & "[Designation], " _
        & "[People Number], " _
        & "[Address Type], " _
        & "[Point Designation]" _
    & ")"
    sqlSelect = "SELECT " _
            & "[ID], " _
            & "[Aktenzeichen], " _
            & "[Fallname], " _
            & "[Nachname], " _
            & "[Vorname], " _
            & "[Unterkunft], " _
            & "[Ort], " _
            & "[Geburtsdatum], " _
            & "[Alter], " _
            & "[Geschlecht], " _
            & "[Familienstand], " _
            & "[Staatsangehörigkeit], " _
            & "[Bezeichnung], " _
            & "[Personenkreisnummer], " _
            & "[Adresstyp], " _
            & "[Stellenbezeichnung]" _
        & "FROM [Query Import Changes Added]"
    sql = sqlInsert & sqlSelect
    
    Debug.Print sql
    cnn.Execute sql
    
    cnn.Close
    Set cnn = Nothing
End Sub
