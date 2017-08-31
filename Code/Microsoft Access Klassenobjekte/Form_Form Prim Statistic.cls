VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form Prim Statistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnResidentialstatus_Click()

    ' Open "Aufenthaltsstatus" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Residential-Status"
    btnResidentialstatus.enabled = False
End Sub

Private Sub btnResidentialcourse_Click()

    ' Open "Aufenthaltsverlauf" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Residential-Course"
    btnResidentialcourse.enabled = False
End Sub

Private Sub btnHealth_Click()

    ' Open "Gesundheit" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Health"
    btnHealth.enabled = False
End Sub

Private Sub btnQualifications_Click()

    ' Open "Qualifikationsprofil" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Qualifications"
    btnQualifications.enabled = False
End Sub

Private Sub btnSocialbackground_Click()

    ' Open "Sozialer Hintergrund" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Social-Background"
    btnSocialbackground.enabled = False
End Sub

Private Sub btnLinguisticproficiency_Click()

    ' Open "Sprachkenntnisse" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Linguistic-Proficiency"
    btnLinguisticproficiency.enabled = False
End Sub

Private Sub btnLinguisticproficiencyUnsalaried_Click()

    ' Open "Sprachkenntnisse" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Linguistic-Proficiency-Unsalaried"
    btnLinguisticproficiencyUnsalaried.enabled = False
End Sub

Private Sub btnLinguisticproficiencyOther_Click()

    ' Open "Sprachkenntnisse" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Linguistic-Proficiency-Other"
    btnLinguisticproficiencyOther.enabled = False
End Sub

Private Sub btnBasedata_Click()

    ' Open "Stammdaten" and grey out the calling button
    enableAllButtons
    sfrmReports.SourceObject = "Report.Rep Sec Stat Base-Data"
    btnBasedata.enabled = False
End Sub

Private Sub btnHome_Click()
    
    ' Load "Hauptmenü" when on "Hauptformular"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.sfrmMain.SourceObject = "Form Prim Menu"
    End If
End Sub

Private Sub Form_Load()
    
    ' Set heading to "Blankoformulare" when on "Hauptformular"
    If mdlTools.HasParent(Me) And CurrentProject.AllForms("Form Prim").IsLoaded Then
        Me.Parent.descHeader.Caption = "Statistikformulare"
        Me.btnHome.enabled = True
    End If
End Sub

Private Sub enableAllButtons()
    btnResidentialstatus.enabled = True
    btnResidentialcourse.enabled = True
    btnHealth.enabled = True
    btnQualifications.enabled = True
    btnSocialbackground.enabled = True
    btnLinguisticproficiency.enabled = True
    btnLinguisticproficiencyUnsalaried.enabled = True
    btnLinguisticproficiencyOther.enabled = True
    btnBasedata.enabled = True
End Sub
