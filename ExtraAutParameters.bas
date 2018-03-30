Attribute VB_Name = "FillParameters"
Option Explicit

Public Function Run() As Boolean

Dim oS                          As CSection

    ' prepara la videata dei parametri
    With oSects
        .IniServerName = App.EXEName & ".ini"
        .DescrizioneGenerale = App.FileDescription
        Set oS = .Add("Causali")
    End With
    
    oS.AddKey "Autorizzazione", ListPresentCode, "Impostare la causale di autorizzazione straordinari"
    
    Set oS = oSects.Add("Calcolo")
    
    oS.AddKey "A Fasce Orarie", ComboBox, "Imposta il controllo anche sulla fascia oraria richiesta non solo sulla quantità oraria", "0,1", "No,Sì"

    FrmParameters.Show vbModal
    
End Function


