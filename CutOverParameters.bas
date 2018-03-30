Attribute VB_Name = "FillParameters"
Option Explicit

Public Function Run() As Boolean

Dim oS                          As CSection

    ' prepara la videata dei parametri
    With oSects
        .IniServerName = App.EXEName & ".ini"
        .DescrizioneGenerale = App.FileDescription
    End With
    
    Set oS = oSects.Add("Causali Digitate")
    oS.AddKey "Elenco", ListPresentCode, "Impostare le causali digitate che lasciano invariato lo straordinario"
    
    Set oS = oSects.Add("Causali Straordinarie")
    oS.AddKey "Elenco", ListPresentCode, "Causali di straordinario da cancellare nel caso non vi siano le causali digitate specificate sopra"

    FrmParameters.Show vbModal
    
End Function


