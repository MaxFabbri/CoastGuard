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
Private debugDipValue       As String

Private currDate            As Date

Private logTrack            As Boolean

'Dim digitMarks              As ADODB.Recordset

Dim sm                      As IStackMessage

Dim BreakProfile            As String
Dim WhenAssignProfile       As Integer
Dim ExceptionCodes()        As String

Private Function IMain_Calculate() As ElabSrvr.CalcResultEnum
    
' pre-elaborazione carica il profilo orario nel caso di marcatura digitata
' modificata !!!
' carica il profilo di riposo in mancanza di timbrature
' tutta la parte commentata riguarda la vecchia personalizzazione che utilizzava la causale digitata
    
Dim daily                           As CDailyData

Dim profiles                        As cProfiles
Dim profile                         As cProfile

Dim ProfileCode                     As String

Dim somethingHappened               As Boolean

Dim mark                            As cMark

Dim countMark                       As Integer
Dim isMarkCorrect                   As Boolean


    On Error GoTo Main_ERROR
    
    IMain_Calculate = wpECalcOK
    
    DebugOperation = "Inizializzazione oggetto organize"
    With oDip
    
        .Change DipID
        
        .ReadsDailyProfiles = True
        .ReadsDetailsProfile = True
        .ReadsMarks = True
        .ReadsHours = True
        
        ' ERA necessario che la persona avesse il profilo orario in anagrafica
'        .ReadsProfileSetInAnag = True
        
        .OrganizeData DipID
        
    End With
    
    somethingHappened = False
    
    For currDate = FromDate To ToDate
    
        DebugOperation = "ci sono dati il " & currDate & " ?"
        Set daily = oDip.DailyDatas(currDate)
        If (daily Is Nothing) Then
            GoTo SkipDay
        End If
         
        debugDipValue = oDip.Nominativo & " data " & currDate
        
        With daily
        
'
'            For Each mark In .marks
'                If (mark.Code <> "") Then
'                    ' marcatura con causale digitata
'                    digitMarks.Filter = "causale = '" & mark.Code & "'"
'                    If (digitMarks.RecordCount > 0) Then
'                        ProfileCode = digitMarks.Fields("Profilo").Value
'                        Exit For
'                    End If
'                End If
'            Next
            
            ' testa eccezione sulle causali
            ' se ne trova anche solo 1 il profilo di riposo non viene inserito
            If (.GroupPreAss.Contains(ExceptionCodes).Count = 0) Then
                
                'If (.marks.Count = 0) Then
                
                    ' nessuna marcatura assegna sicurameente il profilo di riposo
                    'oDip.AddProfilePreElab currDate, wpTBProfilo, , BreakProfile
                    'somethingHappened = True
                   
                'ElseIf (WhenAssignProfile = 1) Then
                
                    ' ulteriore possibilità può essere assegnato anche se nella giornata c'è solamente una uscita
                    
                    For Each mark In .marks
                        If Not mark.isDummy Then
                            countMark = countMark + 1
                            isMarkCorrect = (mark.Way = outway)
                        End If
                    Next
                    
                    If isMarkCorrect And (countMark = 1) Then
                    
                        ' ok una sola marcatura nella giornata in uscita
                        ' va a vedere se nella giornata precedente l'ultima marcatura è una entrata
                        ' solo in questo caso può assegnara il R.I.P.
                        
                        For Each mark In oDip.DailyDatas(currDate - 1).marks
                            If Not mark.isDummy Then
                                'countMark = countMark + 1
                                isMarkCorrect = (mark.Way = inway)
                            End If
                        Next
                        
                        If (isMarkCorrect) Then
                            oDip.AddProfilePreElab currDate, wpTBProfilo, , BreakProfile
                            somethingHappened = True
                        End If
                        
                    End If

                        
                        
                    'End If
                    
                'End If
                
            End If

             
'            If (ProfileCode = "") Then
'                GoTo SkipDay
'            End If

'            oDip.AddProfilePreElab currDate, wpTBProfilo, , ProfileCode
'            somethingHappened = True
            
            ' ERA necessario che la persona avesse il profilo orario in anagrafica
'            Set profiles = oDip.GetAssignedProfiles(WeekDay(currDate, vbMonday))

'            For Each profile In profiles
'                If (Trim$(profile.Code) = Trim$(ProfileCode)) Then
'                    ' carica il profilo
'                    DebugOperation = StringFormat("{0} profilo {1} caricato", debugDipValue, profile.Code)
'                    sm.Push DebugOperation
'                    oDip.AddProfilePreElab currDate, wpTBProfilo, profile
'                    somethingHappened = True
'                End If
'            Next
            
        End With
        
SkipDay:

    Next
    
    If logTrack And somethingHappened Then
        Do While (sm.Count > 0)
            WriteLog sm.Pop
        Loop
    End If
   
    IMain_Calculate = wpECalcOK
    
Main_EXIT:

    Exit Function
    
Main_ERROR:

    IMain_Calculate = wpECalcError
    
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

    On Error GoTo Initialize_ERROR

    DebugOperation = "Lettura file ini"
    
'    Set digitMarks = New ADODB.Recordset
'    With digitMarks.Fields
'        .Append "Causale", adChar, 255
'        .Append "Profilo", adChar, 255
'    End With
'
'    digitMarks.Open
'
'    For i = 1 To Val(RINI("Causali Digitate", "Numero"))
'
'        digitMarks.AddNew
'
'        digitMarks.Fields("Causale") = RINI("Causali Digitate", "Codice Causale Associata Alla Marcatura Digitata " & i)
'        digitMarks.Fields("Profilo") = RINI("Causali Digitate", "Profilo Orario Da Caricare " & i)
'
'    Next
    
    'loadTurnEver = (Val(RINI("Parametri", "Caricare il turno anche se il profilo è bloccato")) > 0)
    
    BreakProfile = RINI("Profilo Riposo", "Codice Profilo")
    'WhenAssignProfile = Val(RINI("Profilo Riposo", "Regola Assegnamento"))
    ExceptionCodes = Split(RINI("Profilo Riposo", "Eccezione Causali"), ",")
    
    logTrack = (Val(RINI("Parametri", "Abilita log")) > 0)
    
    RenFileLOG 99
    
    Set sm = New cStackMessages
    sm.Push ">>> Inizio elaborazione"
    If logTrack Then
        WriteLog sm.Pop
    End If
    
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

    ' deve essere vuoto
    Do While (sm.Count > 0)
        sm.Pop
    Loop

    sm.Push ">>> Fine elaborazione"
    If logTrack Then
        WriteLog sm.Pop
    End If
    Set sm = Nothing
    
    IMain_Terminate = wpECalcOK
    
End Function



