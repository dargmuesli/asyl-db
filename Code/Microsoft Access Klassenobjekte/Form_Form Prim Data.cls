VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form Prim Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnHome_Click()
    
    ' Load "Hauptmenü" when on "Hauptformular"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.sfrmMain.SourceObject = "Form Prim Menu"
    End If
End Sub

Private Sub Form_Current()

    ' Enable and disable the record navigation buttons depending on the current record position
    If Me.CurrentRecord = 1 Then
        btnHeaderFirst.enabled = False
        btnHeaderPrevious.enabled = False
        btnHeaderNext.enabled = True
        btnHeaderLast.enabled = True
    ElseIf Me.CurrentRecord > 1 And Me.CurrentRecord < Me.Recordset.RecordCount Then
        btnHeaderFirst.enabled = True
        btnHeaderPrevious.enabled = True
        btnHeaderNext.enabled = True
        btnHeaderLast.enabled = True
    ElseIf Me.CurrentRecord = Me.Recordset.RecordCount Then
        btnHeaderFirst.enabled = True
        btnHeaderPrevious.enabled = True
        btnHeaderNext.enabled = False
        btnHeaderLast.enabled = False
    End If
End Sub

Private Sub Form_Load()
    
    ' Set heading to "Hauptmenü" when on "Hauptformular"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.descHeader.Caption = "Daten"
        Me.btnHome.enabled = True
    End If
End Sub

Private Sub regData_Change()

    ' Disable all subforms
    Me.sfrmBasedata.enabled = False
    Me.sfrmSocialbackground.enabled = False
    Me.sfrmQualifications.enabled = False
    Me.sfrmLinguisticproficiency.enabled = False
    Me.sfrmHealth.enabled = False
    Me.sfrmResidentialstatus.enabled = False
    Me.sfrmResidentialcourse.enabled = False
    
    ' Enable the selected subform to prevent a page jump to its last element
    Select Case Me.regData.value
        Case 0
            Me.sfrmBasedata.enabled = True
        Case 1
            Me.sfrmSocialbackground.enabled = True
        Case 2
            Me.sfrmQualifications.enabled = True
        Case 3
            Me.sfrmLinguisticproficiency.enabled = True
        Case 4
            Me.sfrmHealth.enabled = True
        Case 5
            Me.sfrmResidentialstatus.enabled = True
        Case 6
            Me.sfrmResidentialcourse.enabled = True
    End Select
End Sub
