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

Private digitMarkCode()        As String
Private glCode()            As String

Private Function IMain_Calculate() As ElabSrvr.CalcResultEnum

' post-elaborazione
' inserisce 1 giorno di riposo GL
' se determinati codici digitati

Dim daily                           As CDailyData
Dim Hours                           As cHours
Dim hour                            As cHour
Dim mark                            As cMark

Dim osM                             As IStackMessage

Dim somethingHappened               As Boolean


Dim foundDigit                      As Boolean

    On Error GoTo Main_ERROR
    
    IMain_Calculate = wpECalcOK
    
    DebugOperation = "Inizializzazione oggetto organize"
    With oDip
    
        .Change DipID
        
        ' dati relativi ai profili
        .ReadsDailyProfiles = True
        '.ReadsDetailsProfile = True
        
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
            
            If .profiles(1) Is Nothing Then
               GoTo SkipDay
            End If
            
            ' la tipologia giornata deve essere festiva
            If (.Typology = wpTGFeriale) Then
                GoTo SkipDay
            End If
            
            foundDigit = False
            ' lettura causali digitate
            For Each mark In .marks
                If (mark.Code <> "") Then
                    If IsInArray(digitMarkCode, mark.Code) Then
                        ' trovato codice marcature digitata
                        foundDigit = True
                        Exit For
                    End If
                End If
            Next
            
            If foundDigit Then
            
                Set hour = New cHour
            
                ' cancella gli eventuali codici precedenti se persistenti
                Set Hours = .GroupPreAss.Contains(glCode) ' festivo
                If Not (Hours Is Nothing) Then
                    For Each hour In Hours
                        oDip.IWriteHour.DeleteObjHour mCurrDate, hour, 1440
                    Next
                End If
                
                Set Hours = Nothing
                
                Set hour = New cHour
                ' inserisce la causale GL
                With hour
                    .Code = glCode(0)
                    .Minutes = TimeSerial(1, 0, 0)
                    .Origin = 3
                    .Note = "Inserita dalla post"
                End With
    
                oDip.IWriteHour.InsertObjHour mCurrDate, hour
                Set hour = Nothing

                
            End If
            
            
            
            somethingHappened = True
            
        End With
        
SkipDay:

    Next mCurrDate
    
    If oDip.IWriteHour.ToWrite Then
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
    WriteLog DebugOperation
    
    digitMarkCode = Split(RINI("Parametri", "Causali Digitate"), ",")
    glCode = Split(RINI("Parametri", "Codice GL"), ",")
    
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

