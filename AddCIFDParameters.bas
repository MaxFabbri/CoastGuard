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
    
    oS.AddKey "Causali Digitate", ListPresentCode, "Selezionare le causali digitate associata alla marcatura"
    oS.AddKey "Codice GL", ListPresentCode, "Selezionare il codice GL giornata di lavoro da inserire"
    oS.AddKey "Log", ComboBox, "Selezionare se attivare il LOG", "0,1", "No,Sì"

    FrmParameters.Show vbModal
    
End Function


