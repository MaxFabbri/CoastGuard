VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit

Implements IMain

Private DebugOperation      As String
Private DebugDipValue       As String

Private mCurrDate           As Date

Private LogTrack            As Boolean

Private digitCodes()          As String

Private recCode()           As String
Private cfgCode()           As String ' compenso forfettario guardia feriale
Private cfgfCode()          As String ' compenso forfettario guardia festivo
Private glCode()            As String

Private preHolidayTurnCode()    As String ' profili che nei prefestivi calcolano come i festivi

Private Function IMain_Calculate() As ElabSrvr.CalcResultEnum

' post-elaborazione
' inserisce le ore di recupero
' il codice CFG Compenso Forfettario di Guardia
' il giorno di riposo GL

' il calcolo parte per determinati porfili orari e cambia in base al tipo di giornata feriale/festiva

' giorno feriale
' 24 h                    -
' 7 h previste            -
' 8 h (1 CFG)             =
'--------------------------
' 9 h di recupero

' non lavorativo (festivo domenica?)
' 32 h                    -
' 24 h (3 CFGF)           + serve per distinguerlo in paga
' 1 GL                    =
'--------------------------
' 8 h di recupero e 1 GL

Dim daily                           As CDailyData

Dim osM                             As IStackMessage

Dim somethingHappened               As Boolean

Dim minRec                          As Integer

Dim h                               As cHour
Dim hs                              As cHours

Dim mark                            As cMark

Dim cfgMin                          As Integer
Dim glMin                           As Integer

Dim digit                           As Boolean

    On Error GoTo Main_ERROR
    
    IMain_Calculate = wpECalcOK
    
    DebugOperation = "Inizializzazione oggetto organize"
    With oDip
    
        .Change DipID
        
        ' dati relativi ai profili
        .ReadsDailyProfiles = True
        .ReadsDetailsProfile = True
        
        ' dati relativi alle marcature
        .ReadsMarks = True
        .ReadsDigitCodes = True
        
        ' dati relativi alle ore
        .ReadsHours = True
        
        .OrganizeData DipID
        
    End With
    
    Set osM = New cStackMessages
    
    For mCurrDate = FromDate To ToDate
    
        DebugOperation = "ci sono dati il " & mCurrDate & " ?"
        Set daily = oDip.DailyDatas(mCurrDate)
        If (daily Is Nothing) Then
            GoTo SkipDay
        End If
         
        DebugDipValue = "data " & mCurrDate & " dipendente " & oDip.Nominativo
        
        With daily
        
            If (.marks.Count = 0) Then
                GoTo SkipDay
            End If
            
            osM.Push String(80, "_")
            osM.Push DebugDipValue
            
            If (.profiles(1) Is Nothing) Then
               GoTo SkipDay
            End If
            
            DebugView "test profiloc superato"
            
            ' verifica le causali digitate sulle marcature
            digit = False
            For Each mark In .marks
                If (mark.Code <> "") Then
                    If IsInArray(digitCodes, mark.Code) Then
                        digit = True
                    End If
                End If
            Next
            If Not digit Then
                GoTo SkipDay
            End If
            
            DebugView "test codice causale superata"

            Set h = New cHour
            
            ' cancella gli eventuali codici precedenti se persistenti
            ' ma solitamente non c'è nella
            Set hs = .GroupPreAss.Contains(cfgCode) ' feriale
            If Not (hs Is Nothing) Then
                For Each h In hs
                    oDip.IWriteHour.DeleteObjHour mCurrDate, h, 1440
                Next
            End If
            
            ' cancella gli eventuali codici precedenti se persistenti
            Set hs = .GroupPreAss.Contains(cfgfCode) ' festivo
            If Not (hs Is Nothing) Then
                For Each h In hs
                    oDip.IWriteHour.DeleteObjHour mCurrDate, h, 1440
                Next
            End If
            
            Set hs = .GroupPreAss.Contains(glCode)
            If Not (hs Is Nothing) Then
                For Each h In hs
                    oDip.IWriteHour.DeleteObjHour mCurrDate, h, 1440
                Next
            End If
            
            Set hs = .GroupPreAss.Contains(recCode)
            If Not (hs Is Nothing) Then
                For Each h In hs
                    oDip.IWriteHour.DeleteObjHour mCurrDate, h, 1440
                Next
            End If
            
            Set h = Nothing
            
            If oDip.IWriteHour.ToWrite Then
                oDip.IWriteHour.WriteObjHours
            End If
            
            oDip.ReloadObjHours
            
            ' non serve rileggere le ore
            'oDip.ReloadObjHours

            DebugOperation = StringFormat("Profilo {0} valido", .profiles(1).Code)
            osM.Push DebugOperation
            
            ' nel caso in cui trovi questi profili ed il giorno successivo sia festivo
            ' (giornata sia prefestiva)
            ' deve essere calcolata come festivo
            If IsInArray(preHolidayTurnCode, .profiles(1).Code) Then
            
                ' doveva essere un prefestivo è diventato il sabato
'                If (oDip.DailyDatas(mCurrDate + 1).Typology = wpTGFestDom) Or _
'                   (oDip.DailyDatas(mCurrDate + 1).Typology = wpTGFestInfra) Then
'                    .Typology = wpTGFestInfra
'                End If
            
                ' legge la giornata successiva e se festiva
                ' cambia la tipologia di questa giornata e la inserisce come non feriale
                If (WeekDay(mCurrDate, vbMonday) = 6) Then
                    .Typology = wpTGFestInfra
                End If

            End If
            
            ' verifica la tipologia della giornata
            If (.Typology = wpTGFeriale) Then
            
                ' giorno feriale
                ' 24 h                    -
                ' 7 h previste            -
                ' 8 h (1 CFG)             =
                '--------------------------
                ' 9 h di recupero
                minRec = 24 * 60
                
                ' calcolo ore previste
                ' le legge dalle ore minime se diverse da 0 altrimenti le legge dalle ore massime
                If (TimeToMinute(.profiles(1).MinimalHours) > 0) Then
                    minRec = minRec - TimeToMinute(.profiles(1).MinimalHours)
                Else
                    minRec = minRec - TimeToMinute(.profiles(1).MaximalHours)
                End If
                
                minRec = minRec - (8 * 60) ' 8 GFC
                ' inserisce 1 GFC (1 giorno) e recupero
                
                cfgMin = 60
                glMin = 0
                
                DebugView "tipo giorno feriale " & minRec
                
            Else
                ' non lavorativo (festivo domenica?)
                ' 32 h                    -
                ' 24 h (3 CFGF)           +
                ' 1 GL                    =
                '--------------------------
                ' 8 h di recupero e 1 GL
                minRec = 32 * 60
                minRec = minRec - (8 * 60) * 3 ' 8 GFC
                ' inserisce 3 GFC (sarebbero 3 giorni in paga inserisce 3 ore) 1 GL (inserisce 1 ora sarebbe 1 giorno in paga) e recupero
                
                cfgMin = 60 * 3
                glMin = 60
                
                DebugView "tipo giorno festivo " & minRec

            End If
            
            Set h = New cHour
            
            ' recupero
            h.Code = recCode(0)
            h.Minutes = TimeSerial(0, minRec, 0)
            h.Origin = 3
            h.Note = "Inserita dalla post"

            oDip.IWriteHour.InsertObjHour mCurrDate, h
            
            DebugView "inserito codice recupero " & h.Code & " in data " & mCurrDate
            
            Set h = New cHour
            
            ' 1 o 3 (gg) GFC
            
            If (.Typology = wpTGFeriale) Then
                h.Code = cfgCode(0)
            Else
                h.Code = cfgfCode(0)
            End If
            
            h.Minutes = TimeSerial(0, cfgMin, 0)
            h.Origin = 3
            h.Note = "Inserita dalla post"

            oDip.IWriteHour.InsertObjHour mCurrDate, h
            
            DebugView "inserito codice cfg " & h.Code & " in data " & mCurrDate
            
            oDip.ReloadObjHours
            
            If (glMin > 0) Then
            
                'MsgBox "GL"
                
                Set h = New cHour
                ' 1 GL
                h.Code = glCode(0)
                h.Minutes = TimeSerial(0, glMin, 0)
                h.Origin = 3
                h.Note = "Inserita dalla post"
                oDip.IWriteHour.InsertObjHour mCurrDate, h
                
                DebugView "inserito codice gl " & h.Code & " in data " & mCurrDate
                
            End If
            
            Set h = Nothing
            
            somethingHappened = True
            
        End With
        
SkipDay:

    Next mCurrDate
    
    If oDip.IWriteHour.ToWrite Then
        DebugView "scrive in data " & mCurrDate
        oDip.IWriteHour.WriteObjHours
    End If
    
    If LogTrack And somethingHappened Then
        ' scrive il log
        Do While (osM.Count > 0)
            WriteLog osM.Pop
        Loop
    End If
    
    Set osM = Nothing
   
    IMain_Calculate = wpECalcOK
    
Main_EXIT:

    Exit Function
    
Main_ERROR:

    IMain_Calculate = wpECalcError
    
    DebugView "errore"
    
    If InDesign() Then
        Debug.Print Err.Description
        Debug.Assert False
        Resume
    End If
    
    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".Main(" & DebugOperation & ")", VbMsgBoxStyle.vbAbortRetryIgnore Or VbMsgBoxStyle.vbQuestion, True)
    Case VbMsgBoxResult.vbRetry
        Resume
    Case VbMsgBoxResult.vbIgnore
        Resume Next
    Case VbMsgBoxResult.vbCancel, VbMsgBoxResult.vbAbort
        Resume Main_EXIT
    End Select

End Function

Private Function IMain_Initialize() As ElabSrvr.CalcResultEnum

    DebugOperation = "Lettura file ini"
    
    digitCodes = Split(RINI("Parametri", "Causali Digitate"), ",")
    
    recCode = Split(RINI("Parametri", "Codice Recupero"), ",")
    cfgCode = Split(RINI("Parametri", "Codice CFG Feriale"), ",")
    cfgfCode = Split(RINI("Parametri", "Codice CFG Festivo"), ",")
    glCode = Split(RINI("Parametri", "Codice GL"), ",")
    
    preHolidayTurnCode = Split(RINI("Parametri", "Elenco Profili Calcolo Del Sabato Come Festivo"), ",")
    
    LogTrack = (Val(RINI("Parametri", "Log")) <> 0)
    
    
    IMain_Initialize = wpECalcOK
    
Initialize_EXIT:

    Exit Function
    
Initialize_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".Initialize(" & DebugOperation & ")", VbMsgBoxStyle.vbAbortRetryIgnore Or VbMsgBoxStyle.vbQuestion, True)
    Case vbRetry
        Resume
    Case vbCancel, vbAbort
        IMain_Initialize = False
        Resume Initialize_EXIT
    End Select

End Function

Private Property Let IMain_Phase(ByVal RHS As ElabSrvr.CustElabPhases)

End Property

Private Property Get IMain_Phase() As ElabSrvr.CustElabPhases

End Property

Private Function IMain_Terminate() As ElabSrvr.CalcResultEnum
    IMain_Terminate = wpECalcOK
End Function

