VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form Sec Import Source-Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private settings As New clsSettings
Private tableData As New Dictionary

Private Function gTableData(Optional filePath As String) As Dictionary

    ' Try to initialize table data
    If tableData.Count = 0 Then
        Set tableData = getTableData(filePath)
    End If
    
    ' Return table data
    Set gTableData = tableData
End Function

Private Sub cmbAccommodation_Change()
    allSetUpdate
End Sub

Private Sub cmbAdresstype_Change()
    allSetUpdate
End Sub

Private Sub cmbAge_Change()
    allSetUpdate
End Sub

Private Sub cmbBirthdate_Change()
    allSetUpdate
End Sub

Private Sub cmbCasename_Change()
    allSetUpdate
End Sub

Private Sub cmbDesignation_Change()
    allSetUpdate
End Sub

Private Sub cmbFilenumber_Change()
    allSetUpdate
End Sub

Private Sub cmbFirstname_Change()
    allSetUpdate
End Sub

Private Sub cmbGender_Change()
    allSetUpdate
End Sub

Private Sub cmbJobtitle_Change()
    allSetUpdate
End Sub

Private Sub cmbLastname_Change()
    allSetUpdate
End Sub

Private Sub cmbMaritalstatus_Change()
    allSetUpdate
End Sub

Private Sub cmbNationality_Change()
    allSetUpdate
End Sub

Private Sub cmbPeoplenumber_Change()
    allSetUpdate
End Sub

Private Sub cmbPlace_Change()
    allSetUpdate
End Sub

Private Sub cmdSource_Click()
    Dim newFilePath As String
    
    Me.cmdSource.enabled = False
    
    ' Start from last chosen path or from project directory
    If FileFolderExists(Me.Parent.saveFilePath) Then
        newFilePath = Left(Me.txtSource, InStrRev(Me.Parent.saveFilePath, "\") - 1)
    Else
        newFilePath = CurrentProject.Path
    End If
    
    ' Open folder browser dialog
    newFilePath = OpenFilename(newFilePath, "Exceldatei auswählen", "Excel 2002-2003 (*.xls)|Excel 2007 (*.xlsx)|Alle Dateien (*.*)")
    
    ' Check if new path was selected
    If Len(newFilePath) > 0 And Len(Dir(newFilePath)) > 0 Then
        
        ' Load new path
        Set tableData = New Dictionary
        loadFilePath newFilePath
        allSetUpdate
    Else
    
        ' Check if valid path already exists
        If FileFolderExists(Me.Parent.saveFilePath) Then
        
            ' Load old path
            loadFilePath Me.Parent.saveFilePath
        Else
        
            ' Reset inputs
            Me.txtSource.value = "N/A"
            Me.txtSource.enabled = False
            Me.txtSource.Locked = False
            Me.cmbSheet.RowSource = ""
            Me.cmbSheet.value = "N/A"
            Me.cmbSheet.enabled = False
            
            updateAllComboboxes True, "N/A", False
            
            Me.Parent.btnNext.enabled = False
            Me.Parent.btnComplete.enabled = False
            
            If Me.Parent.saveFilePath <> "" Then
                MsgBox "Die angegebene Datei kann nicht gefunden werden."
                
                Me.Parent.saveFilePath ""
            End If
        End If
    End If
    
    Me.cmdSource.enabled = True
End Sub

Public Sub loadFilePath(filePath As String)
    Dim key As Variant
    
    Me.Parent.saveFilePath = filePath
    Me.txtSource = filePath
    Me.txtSource.enabled = True
    Me.txtSource.Locked = True
    
    ' Load tablenames into combobox and select the first
    With Me.cmbSheet
        .value = "Loading ..."
        .enabled = True
        .RowSource = ""
        
        ' Read all worksheets to the combobox
        For Each key In gTableData(filePath).Keys
            .AddItem key
        Next key
        
        .SetFocus
        .ListIndex = 0
    End With
End Sub

Private Sub cmbSheet_Change()
    Dim i As Integer
    Dim itemIndex As Integer
    Dim tableHeaders As collection
    Dim tableHeader As Variant
    Dim cmbElements As New collection
    Dim cmbElement As Object
    Dim cmbValue As collection
    Dim allSet As Boolean
    
    Set cmbElements = getCmbElements
    
    ' Clear all dropdown items from all comboboxes
    For Each cmbElement In cmbElements
        cmbElement.RowSource = ""
    Next cmbElement
    
    ' Get all table headers
    Set tableHeaders = gTableData()(gTableData().Keys(Me.cmbSheet.ListIndex))
    
    ' Fill all comboboxes with table headers
    For Each tableHeader In tableHeaders
        For Each cmbElement In cmbElements
            cmbElement.AddItem tableHeader
        Next cmbElement
    Next tableHeader
    
    i = 0
    allSet = True
    
    ' Iterate over all comboboxes
    For Each cmbValue In settings.pCmbValues
        Dim cmbValueValue As Variant
        
        itemIndex = -1
        
        ' Search table headers for known value
        For Each cmbValueValue In cmbValue
            itemIndex = GetCollectionItemIndex(tableHeaders, cmbValueValue)
            
            ' Exit if item was found
            If itemIndex <> -1 Then
                Exit For
            End If
        Next cmbValueValue
        
        ' Mark findings as incomplete
        If itemIndex = -1 Then
            allSet = False
        End If
        
        ' Select found item in corresponding combobox
        cmbElements.item(i + 1) = cmbElements.item(i + 1).ItemData(itemIndex)
        
        i = i + 1
    Next cmbValue
    
    ' Enable all lower inputs
    updateAllComboboxes False, "", True
    
    ' Allow next step if findings are complete
    If allSet Then
        Me.Parent.btnNext.enabled = True
        Me.Parent.btnNext.SetFocus
        Me.Parent.btnComplete.enabled = True
    Else
        Me.Parent.btnNext.enabled = False
        Me.Parent.btnComplete.enabled = False
    End If
End Sub

Public Sub allSetUpdate()
    Dim allSet As Boolean
    Dim cmbElements As New collection
    Set cmbElements = getCmbElements
    Dim cmbElement As Object
    Dim test As Variant
    
    allSet = True
    
    ' Check every combobox for a valid selection
    For Each cmbElement In cmbElements
        cmbElement.SetFocus
        If cmbElement.Text = "N/A" Or cmbElement.Text = "" Then
            allSet = False
            
            Exit For
        End If
    Next
    
    ' Allow next step if findings are complete
    If allSet Then
        Me.Parent.btnNext.enabled = True
        Me.Parent.btnNext.SetFocus
        Me.Parent.btnComplete.enabled = True
    Else
        Me.Parent.btnNext.enabled = False
        Me.Parent.btnComplete.enabled = False
    End If
End Sub

Private Sub updateAllComboboxes(clear As Boolean, value As String, enabled As Boolean)
    Dim cmbElements As New collection
    Dim cmbElement As Object
    
    Set cmbElements = getCmbElements
    
    ' All table header comboboxes
    For Each cmbElement In cmbElements
        
        ' Remove dropdown items
        If clear Then
            cmbElement.RowSource = ""
        End If
        
        ' Set text
        If value <> "" Then
            cmbElement.value = value
        End If
        
        ' En-/Disable
        cmbElement.enabled = enabled
    Next
End Sub

Public Function getCmbElements() As collection
    Dim cmbElements As New collection
    
    ' Add every table header combobox to collection
    cmbElements.Add Me.cmbID
    cmbElements.Add Me.cmbFilenumber
    cmbElements.Add Me.cmbCasename
    cmbElements.Add Me.cmbLastname
    cmbElements.Add Me.cmbFirstname
    cmbElements.Add Me.cmbAccommodation
    cmbElements.Add Me.cmbPlace
    cmbElements.Add Me.cmbBirthdate
    cmbElements.Add Me.cmbAge
    cmbElements.Add Me.cmbGender
    cmbElements.Add Me.cmbMaritalstatus
    cmbElements.Add Me.cmbNationality
    cmbElements.Add Me.cmbDesignation
    cmbElements.Add Me.cmbPeoplenumber
    cmbElements.Add Me.cmbAdresstype
    cmbElements.Add Me.cmbJobtitle
    
    Set getCmbElements = cmbElements
End Function

Private Sub Form_Load()

    ' Start with disabled content
    changeActivity False
End Sub

Public Sub changeActivity(state As Boolean)
    Dim cmbElements As New collection
    Set cmbElements = getCmbElements
    Dim cmbElement As Object
    
    ' Change all elements to enabled or disabled / locked or unlocked mode
    Me.txtSource.enabled = state
    Me.txtSource.Locked = state
    
    Me.cmbSheet.enabled = state
    
    For Each cmbElement In cmbElements
        cmbElement.enabled = state
    Next
End Sub

Public Function getSheets() As collection
    Dim sheets As New collection
    Dim i As Long
    
    ' Add all items of the sheet combobox to a collection
    With Me.cmbSheet
        For i = 0 To .ListCount
            If Not IsNull(.ItemData(i)) Then
                sheets.Add .ItemData(i)
            End If
        Next i
    End With
    
    Set getSheets = sheets
End Function

Public Function getSheetIndex() As Long
    getSheetIndex = Me.cmbSheet.ListIndex
End Function

Public Function getComboboxItems() As collection
    Dim comboboxItems As New collection
    Dim i As Long
    
    ' Add all items of the comboboxex to a collection
    With Me.cmbFilenumber
        For i = 0 To .ListCount
            If Not IsNull(.ItemData(i)) Then
                comboboxItems.Add .ItemData(i)
            End If
        Next i
    End With
    
    Set getComboboxItems = comboboxItems
End Function

Public Function getComboboxIndexes() As collection
    Dim comboboxElements As collection
    Dim comboboxElement As Variant
    Dim comboboxIndexes As New collection
    
    Set comboboxElements = getCmbElements()
    
    ' Add all indexes of the comboboxex to a collection
    For Each comboboxElement In comboboxElements
        comboboxIndexes.Add comboboxElement.ListIndex
    Next comboboxElement
    
    Set getComboboxIndexes = comboboxIndexes
End Function
