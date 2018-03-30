VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMain

Private DebugOperation      As String
Private digitCodes()        As String
Private overCodes()         As String


Private Function IMain_Calculate() As ElabSrvr.CalcResultEnum

Dim hour                            As cHour

Dim mark                            As cMark

Dim Day                            As CDailyData

Dim currDate                        As Date

Dim debugDip                        As String

    On Error GoTo Main_ERROR

    IMain_Calculate = wpECalcOK

    DebugOperation = "Inizializzazione oggetto organize"
    With oDip
        .Change DipID
        .ReadsHours = True
        .ReadsMarks = True
        .ReadsDailyProfiles = True
        .OrganizeData DipID
    End With

    For currDate = FromDate To ToDate

        debugDip = StringFormat("lettura dati dipendente {0} in data {1} ", oDip.Nominativo, currDate)

        Set Day = oDip.DailyDatas(currDate)

        debugDip = StringFormat("- data {0} - {1}", currDate, oDip.Nominativo)

        With Day

            DebugOperation = StringFormat("Ricerca se profilo valido {0}", debugDip)

            If (.profiles(1) Is Nothing) Then
                GoTo SkipDay
            End If
            
            ' verifica la presenza delle causali digitate
            For Each mark In .marks
                If (mark.Code <> "") Then
                    If IsInArray(digitCodes, mark.Code) Then
                        GoTo SkipDay
                    End If
                End If
            Next
            
            ' cancella le causali di straordinario indicate
            For Each hour In .GroupPreAss.Contains(overCodes)
                DebugOperation = StringFormat("cancellazione codice non digitato {1} minuti {2} {0} ", debugDip, hour.Code, hour.GetMinutes)
                oDip.IWriteHour.DeleteObjHour currDate, hour, hour.GetMinutes
                WriteLog DebugOperation
            Next
            

        End With

        If oDip.IWriteHour.ToWrite Then
            DebugOperation = StringFormat("scrive le ore {0}", debugDip)
            oDip.IWriteHour.WriteObjHours
        End If

SkipDay:

    Next

    IMain_Calculate = wpECalcOK

Main_QUIT:

    Exit Function

Main_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, StringFormat("{0}.main({1})", App.EXEName, DebugOperation), _
                         VbMsgBoxStyle.vbAbortRetryIgnore Or VbMsgBoxStyle.vbQuestion, True)
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    Case vbCancel, vbAbort
        IMain_Calculate = wpECalcError
        Resume Main_QUIT
    End Select


End Function


Private Function IMain_Initialize() As ElabSrvr.CalcResultEnum

    IMain_Initialize = True

    DebugOperation = "Lettura file ini"
    
    digitCodes = Split(RINI("Causali Digitate", "Elenco"), ",")
    overCodes = Split(RINI("Causali Straordinarie", "Elenco"), ",")
    
    RenFileLOG 99

Initialize_EXIT:

    Exit Function

Initialize_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, _
                StringFormat("{0}.MainElab.Initialize({1})", App.EXEName, DebugOperation, vbRetryCancel Or vbQuestion))
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

End Function
