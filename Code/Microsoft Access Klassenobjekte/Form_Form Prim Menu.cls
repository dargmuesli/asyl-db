VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form Prim Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnData_Click()
    
    ' Load "Data" when on "Formular Primary"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.sfrmMain.SourceObject = "Form Prim Data"
    End If
End Sub

Private Sub btnBlankforms_Click()
    
    ' Load "Blank Forms" when on "Formular Primary"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.sfrmMain.SourceObject = "Form Prim Blank-Forms"
    End If
End Sub

Private Sub btnStatistic_Click()

    ' Load "Statistics" when on "Formular Primary"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.sfrmMain.SourceObject = "Form Prim Statistic"
    End If
End Sub

Private Sub Form_Load()
    
    ' Set heading to "Main Menu" when on "Formular Primary"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.descHeader.Caption = "Hauptmenü"
        
        Me.btnData.enabled = True
        Me.btnBlankforms.enabled = True
        Me.btnStatistic.enabled = True
    End If
End Sub
