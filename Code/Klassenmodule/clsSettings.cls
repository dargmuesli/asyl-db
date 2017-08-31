VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public pCmbValues As New collection

Private Sub Class_Initialize()
    Dim cmbIDValues As New collection
    Dim cmbFilenumberValues As New collection
    Dim cmbCasenameValues As New collection
    Dim cmbLastnameValues As New collection
    Dim cmbFirstnameValues As New collection
    Dim cmbAccommodationValues As New collection
    Dim cmbPlaceValues As New collection
    Dim cmbBirthdateValues As New collection
    Dim cmbAgeValues As New collection
    Dim cmbGenderValues As New collection
    Dim cmbMaritalstatusValues As New collection
    Dim cmbNationalityValues As New collection
    Dim cmbDesignationValues As New collection
    Dim cmbPeoplenumberValues As New collection
    Dim cmbAdresstypeValues As New collection
    Dim cmbJobtitleValues As New collection
    
    ' Add strings for automatic header recognition
    cmbIDValues.Add "PERSONID"
    cmbFilenumberValues.Add "Aktenzeichen"
    cmbCasenameValues.Add "Fallname (aktuell gültig)"
    cmbCasenameValues.Add "Fallname_(aktuell_gültig)"
    cmbCasenameValues.Add "Fallname"
    cmbLastnameValues.Add "Nachname"
    cmbFirstnameValues.Add "Vorname"
    cmbAccommodationValues.Add "Unterkunft/Einrichtung"
    cmbAccommodationValues.Add "Unterkunft"
    cmbAccommodationValues.Add "Einrichtung"
    cmbPlaceValues.Add "Ort"
    cmbBirthdateValues.Add "Geburtsdatum"
    cmbAgeValues.Add "Alter der Person am [0-9][0-9].[0-9][0-9].[0-9][0-9][0-9][0-9]"
    cmbAgeValues.Add "Alter_der Person_am [0-9][0-9][#][0-9][0-9][#][0-9][0-9][0-9][0-9]"
    cmbAgeValues.Add "Alter"
    cmbGenderValues.Add "Geschlecht"
    cmbMaritalstatusValues.Add "Familienstand"
    cmbNationalityValues.Add "Schlüssel und Staatsangehörigkeit"
    cmbNationalityValues.Add "Schlüssel_und Staatsangehörigkeit"
    cmbNationalityValues.Add "Schlüssel"
    cmbNationalityValues.Add "Staatsangehörigkeit"
    cmbDesignationValues.Add "Bezeichnung (Personenkreisdaten)"
    cmbDesignationValues.Add "Bezeichnung_(Personenkreisdaten)"
    cmbDesignationValues.Add "Bezeichnung"
    cmbPeoplenumberValues.Add "Personenkreisnummer (Personenkreisdaten)"
    cmbPeoplenumberValues.Add "Personenkreisnummer_(Personenkreisdaten)"
    cmbPeoplenumberValues.Add "Personenkreisnummer"
    cmbAdresstypeValues.Add "AdressTyp"
    cmbJobtitleValues.Add "Stellenbezeichnung (Stelle)"
    cmbJobtitleValues.Add "Stellenbezeichnung_(Stelle)"
    cmbJobtitleValues.Add "Stellenbezeichnung"
    
    Dim cmbValueCollection As New collection
    
    cmbValueCollection.Add cmbIDValues
    cmbValueCollection.Add cmbFilenumberValues
    cmbValueCollection.Add cmbCasenameValues
    cmbValueCollection.Add cmbLastnameValues
    cmbValueCollection.Add cmbFirstnameValues
    cmbValueCollection.Add cmbAccommodationValues
    cmbValueCollection.Add cmbPlaceValues
    cmbValueCollection.Add cmbBirthdateValues
    cmbValueCollection.Add cmbAgeValues
    cmbValueCollection.Add cmbGenderValues
    cmbValueCollection.Add cmbMaritalstatusValues
    cmbValueCollection.Add cmbNationalityValues
    cmbValueCollection.Add cmbDesignationValues
    cmbValueCollection.Add cmbPeoplenumberValues
    cmbValueCollection.Add cmbAdresstypeValues
    cmbValueCollection.Add cmbJobtitleValues
    
    Set pCmbValues = cmbValueCollection
End Sub
