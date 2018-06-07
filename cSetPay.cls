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

' assegnamento del pagato dal pagabile in base ad una percentuale %

' contatori ore eccedenti pagabili
Private counterOver()           As String
Private minCounterOver()        As Integer

' versione 1.0.1
' contatore a ore o a giorni
'Private counterDDOrMinute()     As Boolean

' contatori ore pagate
Private counterPay()            As String
'Private minCounterPay()         As Integer

' percentuale pagabile pagato
Private percentagePay           As Integer
Private percentagePayResidual   As Integer

' numero contatori coinvolti
Private maxCount                As Integer

Private resetPrevPay            As Boolean

' versione 1.0.1
' residuo contatore CFG/F da inserire su contatore REC2 mese successivo a ore 1 gg = 8 ore
Private counterCodeToConvert    As String
Private counterCodeToConverted  As String
Private counterHourConversion   As Integer
' fine versione 1.0.1

Private Function IPutPay_AfterInitializeDip(CustomMessage As String) As Boolean
    IPutPay_AfterInitializeDip = True
End Function

Private Sub IPutPay_BuildOutPutFile(WebCaller As Boolean)

End Sub

Private Function IPutPay_ChangeDay(oday As CDatoGiornaliero) As ePrctInfoStatus

    
'    On Error GoTo Pay_ERROR
'
'Pay_EXIT:
'
'    Exit Function
'
'Pay_ERROR:
'
'    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_ChangeDay" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
'    Case vbAbort
'        Resume Pay_EXIT
'    Case vbRetry
'        Resume
'    Case vbIgnore
'        Resume Next
'    End Select

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

    Dim rsCounters  As ADODB.Recordset
    Dim i           As Integer

    On Error GoTo Pay_ERROR
    
    DebugOut "Sub IPutPay_Initialize"
    
    IPutPay_Initialize = True
    
    RenFileLOG 99
    
    ' legge i contatori
    With oDip
        .FromDate = mFromDate
        .ToDate = mToDate
        .ReadsCounter = True
    End With
    
'    Set rsCounters = oDip.GetDefCounter
'
'    For i = 1 To maxCount
'        ' legge se il contatore è a giorni o a ore
'        rsCounters.Filter = "Codice = " & counterOver(i)
'        If (Not rsCounters.EOF) Then
'            ' true giorni false ore
'            counterDDOrMinute(i) = (rsCounters.Fields("TipoConteggio").Value = 1)
'        End If
'    Next

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

Dim MC              As IMonthlyCalc

    On Error GoTo Pay_ERROR
    
    DebugOut "Sub IPutPay_InitializeDip"
    
    mNominativo = oDip.Nominativo
    
    ' azzera le eventuali variazioni manuali precedenti
    
    If resetPrevPay Then
        For i = 1 To maxCount
            PutCounterValue counterOver(i), mFromDate, 0
        Next
        
        ' salva
        oDip.WriteData
    End If
    
    ' esegue cmq il ricalcolo mensile dei contatori
    Set MC = RP.GetMonthlyCalc
    With MC
        .StartCalc oDip.DipID, mFromDate, mToDate, prpCMCont
        If Not .GetLastActionResult() Then
            MsgBox .GetLastActionErrorMessages, vbExclamation, "Attenzione"
        End If
    End With
    
    If resetPrevPay Then
        oDip.PreLoadAnagData False
        oDip.LoadDetailData
    End If
    
    ' legge i contatori delle ore pagabili
    ' va aggiunto il contatore CFG/F
    For i = 1 To maxCount
        Set rsCounter = oDip.GetCounterFiltered(counterOver(i), mFromDate, True)
        If Not (rsCounter Is Nothing) Then
            If Not rsCounter.EOF Then
                minCounterOver(i) = Val(rsCounter.Fields("Totale") & "")
            End If
        End If
    Next
    
    

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
    Set oS = oSects.Add("Contatori Pagabile")
    oS.AddKey "Numero Contatori", StringText, "Inserire il numero dei contatori di straordinario pagabile"
    maxCont = Val(oS.GetKey("Numero Contatori").Value)
    For i = 1 To maxCont
        oS.AddKey "Codice Contatore" & i, ListCounterCode, "Inserire il codice del contatore contenente le ore di straordinario pagabili " & i
    Next
    oS.AddKey "Rigenerazione Valore Precedente Inserito In Paga", ComboBox, "Valore consigilato Sì", "0,1", "Sì,No"
    
    Set oS = oSects.Add("Contatori Pagato")
    For i = 1 To maxCont
        oS.AddKey "Codice Contatore" & i, ListCounterCode, "Inserire il codice del contatore contenente le ore di straordinario pagate " & i
    Next
    
    ' percentuale da inserire in paga
    oS.AddKey "Percentuale Da Passare In Paga", StringText, "Inserire la percentuale del pagabile da passare in paga"
    
    ' versione 1.0.1
    ' contatori speciali residuo
    Set oS = oSects.Add("Contatore Speciale Residuo")
    oS.AddKey "Codice Contatore", ListCounterCode, "Inserire il codice del contatore su cui calcolare il residuo"
    oS.AddKey "Codice Contatore Residuo A Ore Su Mese Successivo", ListCounterCode, "Inserire il codice del contatore ricalcolate come recupero"
    oS.AddKey "Numero Minuti Conversione Per Singolo Giorno", StringText, "Inserire la quantità di una singola giornata convertita in minuti"
    oS.AddKey "Percentuale Da Passare In Paga", StringText, "Inserire la percentuale del pagabile da passare in paga"
    ' versione 1.0.1
    
    
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

'Dim MC              As IMonthlyCalc
'
'    On Error GoTo Pay_ERROR
'
'    ' esegue cmq il ricalcolo mensile dei contatori
'    Set MC = RP.GetMonthlyCalc
'    With MC
'        .StartCalc oDip.DipID, mFromDate, mToDate, prpCMCont
'        If Not .GetLastActionResult() Then
'            MsgBox .GetLastActionErrorMessages, vbExclamation, "Attenzione"
'        End If
'    End With

    IPutPay_PreInitializeDip = True
    
'Pay_EXIT:
'    Exit Function
'
'Pay_ERROR:
'
'    IPutPay_PreInitializeDip = False
'    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".IPutPay_PreInitializeDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
'    Case vbAbort
'        Resume Pay_EXIT
'    Case vbRetry
'        Resume
'    Case vbIgnore
'        Resume Next
'    End Select

    
End Function

Private Sub IPutPay_ReadINIParameters()

Dim v() As String

Dim i           As Integer

    resetPrevPay = (Val(RINI("Contatori Pagabile", "Rigenerazione Valore Precedente Inserito In Paga")) <> 1)
    
    ' lettura codici contatori
    ' versione 1.0.1
    ' la percentuale viene fatta anche in base alle giornate non solo sulle ore
    ' leggere la tipologia del contatore
    ' se a giorni fare la proporzione in difetto es: 1.5 giorni = 1 giorno
    maxCount = Val(RINI("Contatori Pagabile", "Numero Contatori"))
    
    If (maxCount > 0) Then
    
        ReDim counterOver(1 To maxCount) As String
        ReDim minCounterOver(1 To maxCount) As Integer
        ReDim counterPay(1 To maxCount) As String
'        ReDim counterDDOrMinute(1 To maxCount) As Boolean
        
        'ReDim minCounterPay(1 To maxCount) As Integer
        
        For i = 1 To maxCount
        
            v = Split(RINI("Contatori Pagabile", "Codice Contatore" & i), ",")
            If isStringArrayDim(v) Then
                counterOver(i) = v(0) ' prende il primo
            End If
            
            v = Split(RINI("Contatori Pagato", "Codice Contatore" & i), ",")
            If isStringArrayDim(v) Then
                counterPay(i) = v(0) ' prende il primo
            End If
            
        Next
        
    End If
    
    percentagePay = Val(RINI("Contatori Pagato", "Percentuale Da Passare In Paga"))
    
    ' versione 1.0.1
    v = Split(RINI("Contatore Speciale Residuo", "Codice Contatore"), ",")
    If isStringArrayDim(v) Then
        counterCodeToConvert = v(0)  ' prende il primo
    End If
    v = Split(RINI("Contatore Speciale Residuo", "Codice Contatore Residuo A Ore Su Mese Successivo"), ",")
    If isStringArrayDim(v) Then
        counterCodeToConverted = v(0)  ' prende il primo
    End If
    counterHourConversion = Val(RINI("Contatore Speciale Residuo", "Numero Minuti Conversione Per Singolo Giorno"))
    
    percentagePayResidual = Val(RINI("Contatore Speciale Residuo", "Percentuale Da Passare In Paga"))
    ' fine versione 1.0.1
    
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

Dim i               As Integer

Dim MC              As IMonthlyCalc

Dim valueDD         As Integer
Dim valueResidual   As Integer

    On Error GoTo Pay_ERROR
    
    DebugOut "Function IPutPay_TerminateDip()"
    
    
    For i = 1 To maxCount
        If (minCounterOver(i) > 0) Then
        
            valueDD = (minCounterOver(i) / 100) * percentagePay
            
            If (counterOver(i) = counterCodeToConvert) Then
                
                valueDD = (minCounterOver(i) / 100) * percentagePayResidual
                ' questi sono giorni per cui vanno divisi per 60
                valueDD = valueDD \ 60
                ' calcola il residuo da inserire nel mese successivo
                valueResidual = (minCounterOver(i) \ 60) - valueDD
                PutCounterValue counterCodeToConverted, DateValue(mToDate + 1), valueResidual * counterHourConversion
                
            End If
            
            PutCounterValue counterPay(i), mFromDate, valueDD
            
        Else
            PutCounterValue counterPay(i), mFromDate, 0
        End If
    Next
    
'    If (eState = Successful) Then
    
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
'    End If

    IPutPay_TerminateDip = Successful
    
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

Private Sub PutCounterValue(counterCode As String, Data As Date, Value As Integer)

    Dim rsCounter   As ADODB.Recordset

    Set rsCounter = oDip.GetCounterFiltered(counterCode, Data, True)
    If Not (rsCounter Is Nothing) Then
        With rsCounter
            If Not .EOF Then
                ' aggiorna la variazione
                .Fields("Variazione") = Value
                '.Fields("Parziale") = Value ' questa non funziona in quanto viene alimentata direttamente dalle causali
            End If
        End With
    End If

End Sub

