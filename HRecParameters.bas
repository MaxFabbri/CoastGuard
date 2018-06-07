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
    
    'oS.AddKey "Profili", ListTurnCode, "Elenco profili da includere nel calcolo"
    oS.AddKey "Causali Digitate", ListPresentCode, "Elenco causali digitate che danno luogo al calcolo del recupero ore"
    oS.AddKey "Codice Recupero", ListPresentCode, "Selezionare il codice di recupero da inserire sul cartellino"
    oS.AddKey "Codice CFG Feriale", ListPresentCode, "Selezionare il codice di CFG feriale Compenso Forfettario di Guardia da inserire sul cartellino"
    oS.AddKey "Codice CFG Festivo", ListPresentCode, "Selezionare il codice di CFG festivo Compenso Forfettario di Guardia da inserire sul cartellino"
    oS.AddKey "Codice GL", ListPresentCode, "Selezionare il codice GL giornata di lavoro da inserire sul cartellino"
    
    ' versione 1.0.4
    'Elenco Profili Calcolo Del Sabato Come Festivo
    oS.AddKey "Elenco Profili Calcolo Del Sabato Come Festivo", ListTurnCode, "Indicare i turni per il calcolo CFG e GL sui sabati"
    
    oS.AddKey "Log", ComboBox, "Selezionare se attivare il LOG", "0,1", "No,Sì"

    FrmParameters.Show vbModal
    
End Function


