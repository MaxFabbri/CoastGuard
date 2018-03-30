VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPutPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Implements IPutPay

' dati standard

Private mFromDate               As Date
Private mToDate                 As Date

Private DebugOperation          As String

Private mNominativo             As String

' contatori ore eccedenti
Private counterOver()           As String
Private minCounterOver()        As Integer

' causali straordinari compensati negativi
Private codeMinus()             As String

' contatore ore a recupero
Private counterRec              As String
Private minCounterRec           As Integer

' causale ore recuperate
Private codeRecMinus            As String

' versione 1.0.1
' causale ore mancanti
Private missingCode             As String

' contatore e codice ore mancanti
Private missingCounterCode      As String
Private minMissingCounter       As Integer

' ore mancanti recuperate
Private missingCodeMinus        As String

' causali di straordinario nel caso siano da compensare con straordinari STR-
Private codeOver()              As String

' causale di recupero nel caso sia da compensare con REC-
Private codeRec                 As String

Private maxCount                As Integer

Private causIDs                 As String

Private Function IPutPay_AfterInitializeDip(CustomMessage As String) As Boolean
    IPutPay_AfterInitializeDip = True
End Function

Private Sub IPutPay_BuildOutPutFile(WebCaller As Boolean)

End Sub

Private Function IPutPay_ChangeDay(oday As CDatoGiornaliero) As ePrctInfoStatus

Dim minus       As Integer
Dim Value       As Integer

Dim hour        As cOra

Dim totOver     As Integer

Dim i           As Integer
            
    On Error GoTo Pay_ERROR

    DebugOut "Function IPutPay_ChangeDay"
    
    
'    ' sommatoria straordinari recuperabili
'    For i = 1 To maxCount
'        totOver = totOver + minCounterOver(i)
'    Next
        
    With oday
    
        If Not (.Profili(1) Is Nothing) Then
        
            ' recupera la differenza oraria tra le ore massime del profilo orario
            ' e il totale delle ore di presenza/assenza da conteggiare (senza ignora)
            minus = TimeToMinute(.Profili(1).OreMassime) - .ORE.GetNotIgnore.Sum()
            
            If (minus > 0) Then
            
                ' controlla prima il cumulo ore sulla causale di recupero
                Value = Min(minus, minCounterRec)
                
                If (Value > 0) Then
                    ' inserisce causale di REC-
                    Set hour = New cOra
                    hour.Causale = codeRecMinus
                    hour.DateRif = oday.Giorno
                    hour.Ora = TimeSerial(0, Value, 0)
                    oDip.AddObjHour hour
                    Set hour = Nothing
                    minCounterRec = minCounterRec - Value
                    IPutPay_ChangeDay = Successful
                End If
                
                ' poi controlla il cumulo ore sulle ore di straordinario compensabili
                For i = 1 To maxCount
                    minus = minus - Value
                    Value = Min(minus, minCounterOver(i))
                    If (Value > 0) Then
                        ' inserisce causale di STR-
                        Set hour = New cOra
                        hour.Causale = codeMinus(i)
                        hour.DateRif = oday.Giorno
                        hour.Ora = TimeSerial(0, Value, 0)
                        oDip.AddObjHour hour
                        Set hour = Nothing
                        minCounterOver(i) = minCounterOver(i) - Value
                        IPutPay_ChangeDay = Successful
                    End If
                Next
                
                ' il resto se ne rimane viene inserito come ore mancanti
                minus = minus - Value
                If (minus > 0) Then
                    ' mancano ancora dei minuti ma non ci sono più ore da recuperare
                    Set hour = New cOra
                    hour.Causale = missingCode
                    hour.DateRif = oday.Giorno
                    hour.Ora = TimeSerial(0, minus, 0)
                    oDip.AddObjHour hour
                    Set hour = Nothing
                    IPutPay_ChangeDay = Successful
                End If
                    
            Else
            
                ' verifica se ci sono ore di recupero o di straordinario da compensare con ore mancanti
                If (minMissingCounter > 0) Then
                
                    ' verifica prima il codice di recupero
                    totOver = .ORE.ContainsCode(codeRec).Sum()
                    Value = Min(totOver, minMissingCounter)
                    
                    If (Value > 0) Then
                    
                        ' inserisce causale di REC-
                        Set hour = New cOra
                        hour.Causale = codeRecMinus
                        hour.DateRif = oday.Giorno
                        hour.Ora = TimeSerial(0, Value, 0)
                        oDip.AddObjHour hour
                        Set hour = Nothing
                        
                        ' associata alla causale di recupero ore mancanti OreM-
                        Set hour = New cOra
                        hour.Causale = missingCodeMinus
                        hour.DateRif = oday.Giorno
                        hour.Ora = TimeSerial(0, Value, 0)
                        oDip.AddObjHour hour
                        Set hour = Nothing
                        
                        minMissingCounter = minMissingCounter - Value
                        
                        IPutPay_ChangeDay = Successful
                        
                    End If

                    For i = 1 To maxCount
                    
                        ' poi i codici di straordinario ordinati...
                        totOver = .ORE.ContainsCode(codeOver(i)).Sum()
                        ' inserisce lo straodinario recuperato
                        Value = Min(totOver, minMissingCounter)
                        
                        If (Value > 0) Then
                        
                            ' inserisce la causale STR-
                            Set hour = New cOra
                            hour.Causale = codeMinus(i)
                            hour.DateRif = oday.Giorno
                            hour.Ora = TimeSerial(0, Value, 0)
                            oDip.AddObjHour hour
                            Set hour = Nothing
                            
                            ' associata alla causale di recupero ore mancanti OreM-
                            Set hour = New cOra
                            hour.Causale = missingCodeMinus
                            hour.DateRif = oday.Giorno
                            hour.Ora = TimeSerial(0, Value, 0)
                            oDip.AddObjHour hour
                            Set hour = Nothing
                        
                            minMissingCounter = minMissingCounter - Value
                            IPutPay_ChangeDay = Successful
                                
                        End If
                    Next
                End If
                
            End If
        End If
    End With
    
    
Pay_EXIT:

    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_ChangeDay" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Function IPutPay_CheckPeriod() As Boolean

    IPutPay_CheckPeriod = oDip.IsInForce(mFromDate, mToDate)
    
End Function

Private Function IPutPay_DipSelection(ByVal dFromDate As Date, ByVal dToDate As Date, ByVal Name As String) As ADODB.Recordset

End Function

Private Property Let IPutPay_FromDate(ByVal RHS As Date)
    mFromDate = RHS
End Property

Private Property Get IPutPay_FromDate() As Date
    IPutPay_FromDate = mFromDate
End Property

Private Property Get IPutPay_FullCalendar() As Boolean
    IPutPay_FullCalendar = True
End Property

Private Function IPutPay_GetContext() As String

End Function

Private Property Get IPutPay_GetFileList() As String()

End Property

Private Function IPutPay_GetPassword() As String

End Function

Private Function IPutPay_GetUser() As String

End Function

Private Function IPutPay_Initialize(WebCaller As Boolean) As Boolean

Dim i           As Integer

    On Error GoTo Pay_ERROR
    
    DebugOut "Sub IPutPay_Initialize"
    
    IPutPay_Initialize = True
    
    RenFileLOG 99
    
    ' legge le ore e i contatori
    With oDip
        .FromDate = mFromDate
        .ToDate = mToDate
        
        .ReadsDailyCalendar = True
        .ReadsDailyProfiles = True
        
        .ReadsCounter = True
        
        .ReadsHours = True
        
    End With
    
    ' ricerca i CausID x segare direttamente le ore
    ' questi sono gli straordinari negativi
    causIDs = Concat(causIDs, codeRecMinus)
    For i = 1 To maxCount
        causIDs = Concat(causIDs, codeMinus(i))
    Next
    
    ' versione 1.0.1 agginge le ore mancanti
    ' le ore mancanti OreM
    causIDs = Concat(causIDs, missingCode)
    
    ' versione 1.0.1 e le ore mancanti negative
    ' le ore mancanti OreM-
    causIDs = Concat(causIDs, missingCodeMinus)
    
    causIDs = GetCausIDs(causIDs)
    
    IPutPay_Initialize = True
    
Pay_EXIT:
    Exit Function
    

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_Initialize" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Function IPutPay_InitializeDip(CustomMessage As String) As Boolean

Dim rsCounter       As ADODB.Recordset

Dim i               As Integer

    On Error GoTo Pay_ERROR
    
    DebugOut "Sub IPutPay_InitializeDip"
    
    mNominativo = oDip.Nominativo
    
    ' lettura contatori dipendente
    ' --- straodinari recuperabili
    For i = 1 To maxCount
        Set rsCounter = oDip.GetCounterFiltered(counterOver(i), mFromDate, True)
        If Not (rsCounter Is Nothing) Then
            If Not rsCounter.EOF Then
                minCounterOver(i) = Val(rsCounter.Fields("Totale") & "")
            End If
        End If
    Next

    ' --- causale ore di recupero
    Set rsCounter = oDip.GetCounterFiltered(counterRec, mFromDate, True)
    If Not (rsCounter Is Nothing) Then
        If Not rsCounter.EOF Then
            minCounterRec = Val(rsCounter.Fields("Totale") & "")
        End If
    End If
    
    ' versione 1.0.1
    ' --- ore mancanti
    Set rsCounter = oDip.GetCounterFiltered(missingCounterCode, mFromDate, True)
    If Not (rsCounter Is Nothing) Then
        If Not rsCounter.EOF Then
            minMissingCounter = Val(rsCounter.Fields("Totale") & "")
        End If
    End If
    
    Set rsCounter = Nothing
    
    IPutPay_InitializeDip = True
    
Pay_EXIT:
    Exit Function
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_InitializeDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Function


Private Sub IPutPay_InitializePerc(oPerc As IPercentageInfo)

    DebugOut "Sub IPutPay_InitializePerc"
    
    oPerc.DefaultCalculatedValue = Information

End Sub


Private Sub IPutPay_MakeSectionsKeys()

Dim oS              As CSection

Dim i               As Integer

Dim maxCont         As Integer

    On Error GoTo Pay_ERROR
    
    oSects.IniServerName = App.EXEName & ".INI"
    oSects.DescrizioneGenerale = App.FileDescription & " " & GetVersion()

    ' configurazione contatori riguardanti l'eccedenza oraria
    Set oS = oSects.Add("Contatori Di Eccedenza")
    oS.AddKey "Recupero", ListCounterCode, "Inserire il contatore contenente le ore da recuperare"
    oS.AddKey "Numero Contatori Straordinari", StringText, "Inserire il numero dei contatori di straordinario"
    maxCont = Val(oS.GetKey("Numero Contatori Straordinari").Value)
    For i = 1 To maxCont
        oS.AddKey "Straordinario" & i, ListCounterCode, "Inserire il contatore contenente le ore di straordinario pagabili con priorità " & i
    Next
    
    ' versione 1.0.1
    ' causali eccedenti
    Set oS = oSects.Add("Causali Di Eccedenza")
    oS.AddKey "Recupero", ListPresentCode, "Inserire il codice di ore da recuperare"
    ' è uguale al numero dei contatori di straordinario
    ' non lo può più leggere perché ha cambiato sezione mantiene il valore precedente
    ' maxCont = Val(oS.GetKey("Numero Contatori Straordinari").Value)
    For i = 1 To maxCont
        oS.AddKey "Straordinario" & i, ListPresentCode, "Inserire la causale di straordinario con priorità " & i
    Next
    
    ' configurazione contatori riguardanti le causali di recupero orario
    Set oS = oSects.Add("Causali Di Recupero")
    oS.AddKey "Recupero", ListPresentCode, "Inserire il codice di ore recuperate"
    ' ogni contatore ha la sua causale che lo decrementa
    For i = 1 To maxCont
        oS.AddKey "Straordinario" & i, ListPresentCode, "Inserire il codice di straordinario recuperato " & i
    Next
    
    ' versione 1.0.1
    ' ore mancanti contatore e codice causale
    Set oS = oSects.Add("Causali Ore Mancanti")
    oS.AddKey "Codice", ListPresentCode, "Inserire il codice delle ore mancanti"
    oS.AddKey "Contatore", ListCounterCode, "Inserire il contatore contenente le ore mancanti da recuperare"
    oS.AddKey "Codice Negativo", ListPresentCode, "Inserire il codice che decrementa il contatore delle ore mancanti"
    
Pay_EXIT:
    Exit Sub
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_MakeSectionsKeys()")
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Sub

Private Property Get IPutPay_MemoLogin() As Boolean
    IPutPay_MemoLogin = True
End Property

Private Function IPutPay_PreInitialize() As Boolean
    IPutPay_PreInitialize = True
End Function

Private Function IPutPay_PreInitializeDip(CustomMessage As String) As Boolean

Dim MC              As IMonthlyCalc

Dim SQL             As String

    On Error GoTo Pay_ERROR

    DebugOut "Function IPutPay_PreInitializeDip"
    
    ' cancella le causali di recupero assegnate precedentemente dal programma x dipid nel periodo
    If (causIDs <> "") Then
        SQL = "DELETE FROM ORE WHERE DataOremin BETWEEN " & GetSQLDate(mFromDate) & " AND " & GetSQLDate(mToDate) & " AND Causid IN(" & causIDs & ") "
        DebugOperation = SQL
        Debug.Print SQL
        RP.GetPersConnection(oDip.DipID).ExecuteCommand SQL
    End If
    
    ' esegue il ricalcolo mensile dei contatori dopo che le causali assegnate precedentemente dal programma sono state cancellate
    Set MC = RP.GetMonthlyCalc
    With MC
        .StartCalc oDip.DipID, mFromDate, mToDate, prpCMCont
        If Not .GetLastActionResult() Then
            MsgBox .GetLastActionErrorMessages, vbExclamation, "Attenzione"
        End If
    End With

    IPutPay_PreInitializeDip = True
    
Pay_EXIT:
    Exit Function

Pay_ERROR:

    IPutPay_PreInitializeDip = False
    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_PreInitializeDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

    
End Function

Private Sub IPutPay_ReadINIParameters()

Dim v() As String

Dim i           As Integer

    ' lettura codici contatori
    maxCount = Val(RINI("Contatori Di Eccedenza", "Numero Contatori Straordinari"))
    
    If maxCount > 0 Then
    
        ReDim counterOver(1 To maxCount) As String
        ReDim minCounterOver(1 To maxCount) As Integer
        ReDim codeMinus(1 To maxCount) As String
        
        ReDim codeOver(1 To maxCount) As String
    
        For i = 1 To maxCount
        
            v = Split(RINI("Contatori Di Eccedenza", "Straordinario" & i), ",")
            If isStringArrayDim(v) Then
                counterOver(i) = v(0) ' prende il primo
            End If
            
            v = Split(RINI("Causali Di Recupero", "Straordinario" & i), ",")
            If isStringArrayDim(v) Then
                codeMinus(i) = v(0) ' prende il primo
            End If
            
            ' versione 1.02
            v = Split(RINI("Causali Di Eccedenza", "Straordinario" & i), ",")
            If isStringArrayDim(v) Then
                codeOver(i) = v(0) ' prende il primo
            End If
    
        Next
        
    End If
   
    v = Split(RINI("Contatori Di Eccedenza", "Recupero"), ",")
    If isStringArrayDim(v) Then
        counterRec = v(0) ' prende il primo
    End If
    
    v = Split(RINI("Causali Di Recupero", "Recupero"), ",")
    If isStringArrayDim(v) Then
        codeRecMinus = v(0) ' prende il primo
    End If
    
    ' versione 1.0.1
    v = Split(RINI("Causali Di Eccedenza", "Recupero"), ",")
    If isStringArrayDim(v) Then
        codeRec = v(0) ' prende il primo
    End If

    
    ' versione 1.0.1
    ' ore mancanti contatore e codice causale
    missingCode = RINI("Causali Ore Mancanti", "Codice")
    missingCounterCode = RINI("Causali Ore Mancanti", "Contatore")
    missingCodeMinus = RINI("Causali Ore Mancanti", "Codice Negativo")
    
End Sub

Private Property Get IPutPay_SilentMode() As Boolean
    IPutPay_SilentMode = gAutoStart
End Property

Private Sub IPutPay_Terminate(WebCaller As Boolean)

    On Error GoTo Pay_ERROR

    DebugOut "Sub IPutPay_Terminate"
    
Pay_EXIT:
    Exit Sub

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_Terminate" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Sub


Private Function IPutPay_TerminateDip(eState As ePrctInfoStatus, CustomMessage As String) As ePrctInfoStatus

Dim MC              As IMonthlyCalc

    On Error GoTo Pay_ERROR
    
    DebugOut "Function IPutPay_TerminateDip()"
    
    If (eState = Successful) Then
    
        ' sono stati inseriti dei codici di recupero orario li assegna alla persona
        oDip.WriteData
        
        ' dopo l'inserimento dei codici aggiorna mensilmente i contatori
        Set MC = RP.GetMonthlyCalc
        With MC
            .StartCalc oDip.DipID, mFromDate, mToDate, prpCMCont
            If Not .GetLastActionResult() Then
                MsgBox .GetLastActionErrorMessages, vbExclamation, "Attenzione"
            End If
        End With
    End If

    IPutPay_TerminateDip = eState
    
Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_TerminateDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Function

Private Property Let IPutPay_ToDate(ByVal RHS As Date)
    mToDate = RHS
End Property

Private Property Get IPutPay_ToDate() As Date
    IPutPay_ToDate = mToDate
End Property

Private Function IPutPay_ViewCalendar(dFromDate As Date, dToDate As Date, ByVal Name As String) As Boolean

End Function

Private Function IPutPay_ViewFullCalendar(StartDate As Date, EndDate As Date, ByVal Name As String) As ADODB.Recordset

    If Not gAutoStart Then
    
        Dim oC As New cCalendar
        With oC
        
            .Caption = "Periodo " & Name
            .CalendarMode = ThisMonth
            Set .Icon = LoadResPicture(101, 1)
            
            .DynamicPeriod = True
            .Fullmonth = True
            
            Set IPutPay_ViewFullCalendar = .Run()
            
            StartDate = DateValue(.StartDate)
            EndDate = DateValue(.EndDate) + TimeSerial(23, 59, 59)
            
        End With
    Else
    
        ' escuzione automatica mese corrente completo
        StartDate = Now()
        StartDate = DateSerial(Year(StartDate), Month(StartDate), 1)
        EndDate = DateSerial(Year(StartDate), Month(StartDate), DaysOfMonth(StartDate))
        
        Set IPutPay_ViewFullCalendar = GetAnagFilteredRecordset("*", "Dipendenti", vbNullString, , "DipID", StartDate, EndDate)

    End If
    
    mFromDate = StartDate
    mToDate = EndDate

End Function

Private Property Get IPutPay_ViewLOG(LogList As String, Optional WaitSec As Integer = 15) As Boolean

End Property

Private Property Get IPutPay_ViewPercentage() As Boolean
    IPutPay_ViewPercentage = Not gAutoStart
End Property

Private Function GetVersion() As String
    GetVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

