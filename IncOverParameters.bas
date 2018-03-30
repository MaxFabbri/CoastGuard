Attribute VB_Name = "FillParameters"
Option Explicit

Public Function Run() As Boolean

Dim oS                          As CSection

    ' prepara la videata dei parametri
    With oSects
        .IniServerName = App.EXEName & ".ini"
        .DescrizioneGenerale = App.FileDescription
        Set oS = .Add("Parametri")
    End With
    
    oS.AddKey "Causali Digitate", ListPresentCode, "Elenco causali digitate che danno luogo al calcolo del recupero ore"
    oS.AddKey "Causali Da Maggiorare", ListPresentCode, "Elenco causali da maggiorare ogni ora 5'"
    oS.AddKey "Log", ComboBox, "Selezionare se attivare il LOG", "0,1", "No,Sì"

    FrmParameters.Show vbModal
    
End Function


