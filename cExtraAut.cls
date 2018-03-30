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
Private CodesToControl()    As String
Private mAuthCode           As String
Private mCalcFasce          As Boolean

Private Function IMain_Calculate() As ElabSrvr.CalcResultEnum

Dim ohsAut                          As cHours

Dim oH                              As cHour
Dim ohAut                           As cHour

Dim oDay                            As CDailyData

Dim IHD                             As IObjHourData

Dim currDate                        As Date

Dim sFound                          As Date
Dim eFound                          As Date

Dim iAuth                           As Integer
Dim iOver                           As Integer

Dim DebugDip                        As String

    On Error GoTo Main_ERROR

    IMain_Calculate = wpECalcOK

    DebugOperation = "Inizializzazione oggetto organize"
    With oDip
        .Change DipID
        .ReadsHours = True
        .ReadsDailyProfiles = True
        .OrganizeData DipID
    End With

    Set IHD = oDip

    For currDate = FromDate To ToDate

        DebugDip = StringFormat("lettura dati dipendente {0} in data {1} ", oDip.Nominativo, currDate)

        Set oDay = oDip.DailyDatas(currDate)

        DebugDip = StringFormat("in data {0} dipendente {1}", currDate, oDip.Nominativo)

        With oDay

            DebugOperation = StringFormat("Ricerca se profilo valido {0}", DebugDip)

            If (.profiles(1) Is Nothing) Then
                GoTo SkipDay
            End If

            ' aggiunge o cancella anomalie
            iAuth = 0
            iOver = 0
            For Each oH In .PresenceHours
                If (oH.Code = mAuthCode) Then
                    iAuth = iAuth + oH.GetMinutes
                ElseIf oH.ExtraordinaryCode Then
                    iOver = iOver + oH.GetMinutes
                End If
            Next

            If (iOver <= iAuth) Then
                oDip.DeleteWarnings currDate, , wpAnomOreNonAutor
                oDip.DeleteWarnings currDate, , wpAnomOreNonProg
                oDip.WriteData
            ElseIf (iOver > iAuth) Then
            End If

            If Not mCalcFasce Then

                For Each oH In .PresenceHours.Sort(True)
                    If oH.ExtraordinaryCode Then
                        If (oH.GetMinutes > iAuth) Then
                            If ((oH.GetMinutes - iAuth) > 0) And (iAuth > 0) Then

                                'cancella in parte
                                DebugOperation = StringFormat("cancellazione ore di straordinario non autorizzate {0} fascia {1} {2}", DebugDip, Format(TimeValue(oH.StartDate)), Format(TimeValue(oH.EndDate)))
                                IHD.DeleteObjHour currDate, oH, oH.GetMinutes
                                DebugOperation = StringFormat("Modifica ore di straordinario autorizzate {0} da fascia {1} {2} a fascia {3} {4}", DebugDip, Format(TimeValue(oH.StartDate)), Format(TimeValue(oH.EndDate)), Format(sFound), Format(eFound))
                                oH.EndDate = currDate + TimeSerial(0, TimeToMinute(TimeValue(oH.EndDate)) - (oH.GetMinutes - iAuth), 0)
                                oH.Minutes = oH.EndDate - oH.StartDate
                                IHD.InsertObjHour currDate, oH, True

                            Else

                                ' cancella tutto
                                DebugOperation = StringFormat("cancellazione ore di straordinario non autorizzate {0} fascia {1} {2}", DebugDip, Format(TimeValue(oH.StartDate)), Format(TimeValue(oH.EndDate)))
                                IHD.DeleteObjHour currDate, oH, oH.GetMinutes

                            End If
                        End If
                        iAuth = iAuth - oH.GetMinutes
                    End If
                Next


            Else

                If .PresenceHours.isCodeLoaded(mAuthCode) Is Nothing Then
                    ' cancella le causali di straordinario
                    For Each oH In .PresenceHours
                        If oH.ExtraordinaryCode Then
                            DebugOperation = StringFormat("cancellazione ore di straordinario non autorizzate {0} fascia {1} {2}", DebugDip, Format(TimeValue(oH.StartDate)), Format(TimeValue(oH.EndDate)))
                            WriteLog DebugOperation
                            IHD.DeleteObjHour currDate, oH, oH.GetMinutes
                        End If
                    Next
                End If


                ' convalida in base alla fascia oraria
                For Each ohAut In .PresenceHours
                    'legge la fascia oraria delle varie causali di autorizzazione straordinari
                    If UCase(ohAut.Code) = UCase(mAuthCode) Then
                        ' verifica che ci sia uno straordinario dentro la fascia oraria autorizzata
                        For Each oH In .PresenceHours
                            If oH.ExtraordinaryCode Then
                                DebugOperation = StringFormat("Verifica fasce di autorizzazione coperte {0}", DebugDip)
                                oCS.CoverSlot ohAut.StartDate, ohAut.EndDate, oH.StartDate, oH.EndDate, sFound, eFound
                                If (eFound <> oCS.NullSlot) Then

                                    If (sFound <> TimeValue(oH.StartDate)) Or (eFound <> TimeValue(oH.EndDate)) Then
                                        DebugOperation = StringFormat("Modifica ore di straordinario autorizzate {0} da fascia {1} {2} a fascia {3} {4}", DebugDip, Format(TimeValue(oH.StartDate)), Format(TimeValue(oH.EndDate)), Format(sFound), Format(eFound))
                                        IHD.DeleteObjHour currDate, oH, oH.GetMinutes
                                        With oH
                                            .StartDate = currDate + sFound
                                            .EndDate = currDate + eFound
                                            .Minutes = MinuteToTime(TimeToMinute(eFound) - TimeToMinute(sFound))
                                        End With
                                        IHD.InsertObjHour currDate, oH, oH.GetMinutes
                                        WriteLog DebugOperation
                                    End If

                                Else

                                    DebugOperation = StringFormat("cancellazione ore di straordinario non autorizzate {0} fascia {1} {2}", DebugDip, Format(TimeValue(oH.StartDate)), Format(TimeValue(oH.EndDate)))
                                    WriteLog DebugOperation
                                    IHD.DeleteObjHour currDate, oH, oH.GetMinutes

                                End If
                            End If
                        Next
                    End If
                Next

            End If

        End With

        If IHD.ToWrite Then
            DebugOperation = StringFormat("scrive le ore {0}", DebugDip)
            IHD.WriteObjHours
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

Dim INIFile         As String

    IMain_Initialize = True


    INIFile = StringFormat("{0}.ini", App.EXEName)

    DebugOperation = "Lettura file ini"
    With RP

        mAuthCode = .GetSetting(INIFile, "Causali", "Autorizzazione")
        mCalcFasce = (Val(.GetSetting(INIFile, "Calcolo", "A Fasce Orarie")) <> 0)

    End With

    DebugOperation = StringFormat("Rinomina file {0}", GetLocalPath() & App.EXEName & ".LOG")
    RenFileProgr GetLocalPath() & App.EXEName & ".LOG", 99, frmHide.ListApp, ""

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
