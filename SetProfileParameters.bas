Attribute VB_Name = "FillParameters"
Option Explicit

Public Function Run() As Boolean

Dim i                           As Integer

Dim oS                          As CSection

    With oSects
        .IniServerName = App.EXEName & ".ini"
        .DescrizioneGenerale = App.FileDescription
    End With
    
'    Set oS = oSects.Add("Causali Digitate")
'    oS.AddKey "Numero", StringText, "Elencare il numero dei codici marcatura digitati a cui associare i profili orari"
'    For i = 1 To Val(oS.GetKey("Numero").Value)
'        oS.AddKey "Codice Causale Associata Alla Marcatura Digitata " & i, ListPresentCode, "Indicare il codice digitato associato alla marcatura " & i
'        oS.AddKey "Profilo Orario Da Caricare " & i, ListTurnCode, "Indicare il profilo orario da caricare associato al codice sopra " & i
'    Next
    
    Set oS = oSects.Add("Profilo Riposo")
    oS.AddKey "Codice Profilo", ListTurnCode, "Specificare il profilo di riposo da inserire sul cartellino"
    'oS.AddKey "Regola Assegnamento", ComboBox, "Inserire il profilo se mancano le marcature oppure se l'unica marcatura della giornata è una uscita", "0,1", "In mancanza di timbrature,In mancanza di timbrature / l'unica marcatura è una uscita"
    oS.AddKey "Eccezione Causali", ListPresentCode, "Se nel cartellino sono inserite queste causali anche in mancanza di marcature non inserisce il profilo di riposo"
    
    Set oS = oSects.Add("Parametri")
    'oS.AddKey "Caricare il turno anche se il profilo è bloccato", ComboBox, "Permette di caricare il profilo orario associato allo stabilimento anche se in precedenza il profilo sul cartellino era bloccato", "0,1", "No,Sì"
    oS.AddKey "Abilita log", ComboBox, "Consente di abilitare il file di log relativo alle operazioni eseguite da questa elaborazione", "0,1", "No,Sì"
    
    FrmParameters.Show vbModal
    
End Function


