VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form Sec Data Linguistic-Proficiency-Unsalaried"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
    Dim RecordCount As Integer
    
    RecordCount = DCount("[ID]", "[Table Linguistic-Proficiency-Unsalaried]")

    ' Enable and disable the record navigation buttons depending on the current record position
    If Me.CurrentRecord = 1 Then
        btnFooterFirst.enabled = False
        btnFooterPrevious.enabled = False

        If RecordCount > 1 Then
            btnFooterNext.enabled = True
            btnFooterLast.enabled = True
        Else
            btnFooterNext.enabled = False
            btnFooterLast.enabled = False
        End If
    ElseIf Me.CurrentRecord > 1 And Me.CurrentRecord < RecordCount Then
        btnFooterFirst.enabled = True
        btnFooterPrevious.enabled = True
        btnFooterNext.enabled = True
        btnFooterLast.enabled = True
    ElseIf Me.CurrentRecord = Me.Recordset.RecordCount Then
        btnFooterFirst.enabled = True
        btnFooterPrevious.enabled = True
        btnFooterNext.enabled = False
        btnFooterLast.enabled = False
    End If
End Sub

