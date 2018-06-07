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

Private digitCodes()        As String
Private overCodes()         As String

Private Function IMain_Calculate() As ElabSrvr.CalcResultEnum

' post-elaborazione
' maggiorazione oraria 5' su straordinario digitato

Dim daily                           As CDailyData

Dim osM                             As IStackMessage

Dim somethingHappened               As Boolean

Dim hNew                            As cHour
Dim h                               As cHour
Dim hs                              As cHours

Dim mark                            As cMark

Dim digit                           As Boolean

Dim overMinutes                     As Integer

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
            
            If .profiles(1) Is Nothing Then
               GoTo SkipDay
            End If
            
            DebugView "test profilo superato"
            
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
            
            'DebugView "test codice causale superata"
            
            Set hs = .GroupPreAss.Contains(overCodes) ' codici da maggiorare
            
            
            If Not (hs Is Nothing) Then
                For Each h In hs
                
                    overMinutes = (h.GetMinutes \ 60) * 5
                    
                    h.Note = "Inserita dalla post " & App.EXEName
                    h.Origin = 3
                    
                    ' rimane solamente la maggiorazione
                    oDip.IWriteHour.DeleteObjHour mCurrDate, h, h.GetMinutes() - overMinutes
                    
                    DebugView "codice maggiorato rimanente " & h.Code & " in data " & mCurrDate & " " & overMinutes
                    
                    ' originariamente aggiungeva la maggiorazione alle ore della causale
                    
'                    overMinutes = (h.GetMinutes \ 60) * 5
'
'                    If (overMinutes > 0) Then
'
'                        Set hNew = New cHour
'                        hNew.Code = h.Code
'                        hNew.Minutes = TimeSerial(0, overMinutes, 0)
'                        hNew.Origin = 3
'                        hNew.Note = "Inserita dalla post " & App.EXEName
'
'                        oDip.IWriteHour.InsertObjHour mCurrDate, hNew
'
'                        DebugView "inserito codice maggiorato " & hNew.Code & " in data " & mCurrDate
'
'                        Set hNew = Nothing
'
'                    End If
                        
                Next
            End If
            
            
            
        End With
        
SkipDay:

    Next mCurrDate
    
    If oDip.IWriteHour.ToWrite Then
        somethingHappened = True
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
    overCodes = Split(RINI("Parametri", "Causali Da Maggiorare"), ",")
    
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

