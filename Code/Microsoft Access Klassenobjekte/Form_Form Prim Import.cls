VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form Prim Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private intStep As Integer
Private intPreviousStep As Integer

Public saveFilePath As String
Public saveSheets As collection
Public saveSheetIndex As Long
Public saveComboboxItems As collection
Public saveComboboxIndexes As collection
Public savePreImportTableName As String

Private WithEvents frmSourceSettings As Form
Attribute frmSourceSettings.VB_VarHelpID = -1
Private WithEvents frmInterpretation As Form
Attribute frmInterpretation.VB_VarHelpID = -1
Private WithEvents frmChanges As Form
Attribute frmChanges.VB_VarHelpID = -1

Private Sub Form_Load()
    MsgBox "Vor dem Import neuer Daten und in gewissen zeitlichen Abständen ist es immer ratsam eine Sicherung der aktuellen Datenbank anzulegen, damit keine Daten verloren gehen."

    ' Load first subform
    Call openSourceSettingsSubform
End Sub

Private Sub btnCancel_Click()

    ' Exit form
    Call finishExcel

    DoCmd.Close acForm, Me.Name
End Sub

Private Sub btnComplete_Click()
    disableControlButtons
    
    ' Step fast forward
    DoStep "ff"
End Sub

Private Sub btnNext_Click()
    disableControlButtons
    
    ' Step forward
    DoStep "+1"
End Sub

Private Sub btnBack_Click()
    disableControlButtons

    ' Step backward
    DoStep "-1"
End Sub

Private Sub DoStep(direction As String)
    
    ' Decide which task to do next depending on the button pressed
    Select Case Me.sfrmImport.SourceObject
        Case "Form Sec Import Source-Settings"
            Select Case direction
                Case "+1"
                    openInterpretationSubform
                Case "ff"
                    Call frmSourceSettings_Unload(False)
                    finishImport
            End Select
        Case "Form Sec Import Interpretation"
            Select Case direction
                Case "-1"
                    openSourceSettingsSubform
                Case "+1"
                    openChangesSubform
                Case "ff"
                    finishImport
            End Select
        Case "Form Sec Import Changes"
            Select Case direction
                Case "-1"
                    openInterpretationSubform
                Case "ff"
                    finishImport
            End Select
    End Select
End Sub

Private Sub openSourceSettingsSubform()

    ' Load subform "Quelleinstellungen"
    Me.sfrmImport.SourceObject = "Form Sec Import Source-Settings"
    Set frmSourceSettings = Me.sfrmImport.Form
    
    With frmSourceSettings
        ' Add unload event procedure
        .OnUnload = "[Event Procedure]"
        
        ' Try to load saved settings
        If saveFilePath <> "" Then
            .loadFilePath saveFilePath
            
            If Not trySetSavedSourceSettings Then
                MsgBox "Die Quelltabelle hat sich geändert!"
            End If
        End If
    End With
End Sub

Private Sub openInterpretationSubform()

    ' Load subform "Interpretation"
    Me.sfrmImport.SourceObject = "Form Sec Import Interpretation"
    Set frmInterpretation = Me.sfrmImport.Form
    
    ' Enable all buttons for procedure steps
    Me.btnBack.enabled = True
    Me.btnNext.enabled = True
    Me.btnComplete.enabled = True
End Sub

Private Sub openChangesSubform()

    ' Load subform "Veränderungen"
    Me.sfrmImport.SourceObject = "Form Sec Import Changes"
    Set frmChanges = Me.sfrmImport.Form
    
    ' Enable all buttons for procedure steps, except "next"
    Me.btnBack.enabled = True
    Me.btnNext.enabled = False
    Me.btnComplete.enabled = True
End Sub

Private Sub frmSourceSettings_Unload(Cancel As Integer)

    ' Save all subform's variables for future comparison
    saveFilePath = frmSourceSettings.txtSource.value
    
    Set saveSheets = frmSourceSettings.getSheets
    saveSheetIndex = frmSourceSettings.getSheetIndex
    Set saveComboboxItems = frmSourceSettings.getComboboxItems
    Set saveComboboxIndexes = frmSourceSettings.getComboboxIndexes
    
    ' Try to start the import
    If saveFilePath <> "N/A" Then
        startImport
    End If
End Sub

Private Function trySetSavedSourceSettings()
    Dim setSuccess As Boolean
    
    setSuccess = True
    
    ' If any save isn't equal to the new value return false
    If collectionsEqual(Me.saveSheets, frmSourceSettings.getSheets) Then
        frmSourceSettings.cmbSheet.ListIndex = Me.saveSheetIndex
    Else
        setSuccess = False
    End If
    
    ' Check if Excel file has changed
    If collectionsEqual(Me.saveComboboxItems, frmSourceSettings.getComboboxItems) Then
        Dim cmbElements As New collection
        Dim cmbElement As Object
        Dim i As Integer
        
        Set cmbElements = frmSourceSettings.getCmbElements
        i = 0
        
        ' Set each saved selected index
        For Each cmbElement In cmbElements
            i = i + 1
            
            cmbElement.value = cmbElement.ItemData(Me.saveComboboxIndexes.item(i))
            'cmbElement.SetFocus
            'cmbElement.ListIndex = Me.saveComboboxIndexes.item(i)
        Next
    Else
        setSuccess = False
    End If
    
    frmSourceSettings.allSetUpdate
    
    trySetSavedSourceSettings = setSuccess
End Function

Private Sub startImport()

    ' Import Excel file to temporary table
    Call preImportTable("Table Base-Data Import", saveFilePath, saveSheets.item(saveSheetIndex + 1), saveComboboxItems, saveComboboxIndexes)
End Sub

Private Sub finishImport()

    ' Import and exit
    Call importTable(savePreImportTableName)
    Call btnCancel_Click
End Sub

Private Sub disableControlButtons()

    ' Disable all buttons for procedure steps
    Me.btnBack.enabled = False
    Me.btnNext.enabled = False
    Me.btnComplete.enabled = False
End Sub

