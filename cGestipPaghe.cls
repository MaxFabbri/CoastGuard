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

Private mFromDate       As Date
Private mToDate         As Date

Private DebugOperation  As String

Private mNominativo     As String

Dim hFile               As Byte

Dim FileName            As String

Dim fileSchema          As ISchema
Dim fileNameSchema      As String

Dim monthSum            As cPays

Private Type mapC
    Code            As String
    HourField       As String
    MinField        As String
End Type

Private Type mapCounter
    Code            As String
    HourField       As String
    MinField        As String
    typeField       As payType
End Type

Private Type mapM
    Field           As String
    HourField       As String
    MinField        As String
End Type

Private Type mapAnag
    PresentField    As String
    PayField        As String
    DafaultValue    As String
'    TypeField       As String
End Type

Private Type mapPay
    toPayCode   As String
    paidCode    As String
End Type

Private Enum payType
    undefined = -1
    partial = 0
    residual = 1
    total = 2
    totalPlusVariation = 3
End Enum

Dim mapsCart()      As mapC
Dim maxMapCart      As Integer

Dim mapsCount()     As mapCounter
Dim maxMapCount     As Integer

Dim mapsMonth()     As mapM
Dim maxMapMonth     As Integer

Dim mapsAnag()      As mapAnag
Dim maxMapAnag      As Integer

Dim mapsPay()       As mapPay
Dim mapsPayNumber   As Integer

Dim fileProcessed   As Boolean

Private Function IPutPay_AfterInitializeDip(CustomMessage As String) As Boolean
    IPutPay_AfterInitializeDip = True
End Function

Private Sub IPutPay_BuildOutPutFile(WebCaller As Boolean)

End Sub

Private Function IPutPay_ChangeDay(oday As CDatoGiornaliero) As ePrctInfoStatus
            
Dim hour            As cOra

Dim i               As Integer

Dim itemSum         As cPay

    On Error GoTo Pay_ERROR
    
    For Each hour In oday.ORE.Group
        
        For i = 1 To maxMapCart
            ' ricerca le causali mappate da passare
            If (hour.Causale = mapsCart(i).Code) Then
                ' le totalizza mensilmente
                Set itemSum = monthSum(hour.Causale)
                With itemSum
                    ' al momento testa solo se la causale è in centesimi non testa se a giorni/ore
                    .Cent = hour.Centesimi
                    .IsDaily = (hour.CountType = atday)
                    ' ci sono delle causali a giorni come il compenso forfettario di guardia CFG
                    ' da cui bisogna prelevare le ore e farle diventare giorni
                    ' quindi questa funzione non va bene....
                    '.Value = .Value + Hour.ResolvePayQuantity(oday.Due)
                    If Not .IsDaily Then
                        .Value = .Value + hour.ResolvePayQuantity(oday.Due)
                    Else
                        ' da ore a giornate
                        .Value = .Value + (hour.GetMinutes() \ 60)
                    End If
                    .PayCode = hour.CodConvPaghe
                    IPutPay_ChangeDay = Successful
                End With
            End If
        Next
        
    Next
    

Pay_EXIT:

    Exit Function
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.ExeName & ".IPutPay_ChangeDay" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
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
    If InDesign() Then
        IPutPay_GetPassword = "ARTECHADMIN"
    End If
End Function

Private Function IPutPay_GetUser() As String
    If InDesign() Then
        IPutPay_GetUser = "ARTECH"
    End If
End Function

Private Function IPutPay_Initialize(WebCaller As Boolean) As Boolean

    On Error GoTo Pay_ERROR
    
    fileProcessed = True
    
    With oDip
    
        .ReadsHours = True
        .ReadsDailyCalendar = True
        .ReadsDailyProfiles = True
        
        .ReadsDataPeriod = True
        .DataPeriodFiltered = True
        
        .ReadsCounter = True
        
        .FromDate = mFromDate
        .ToDate = mToDate
        
    End With


    ' legge le sezioni di default dall'ini devono corrispondere a quelle dello schema.ini
    fileNameSchema = RINI("OutPutFile", "Schema")
    
    Set fileSchema = New cSchema
    With fileSchema
    
        ' percorso di scrittura dei file
        .FilePath = PresentUserLocalFolder
        
        ' legge lo schema dal server così non è una risorsa e lo posso modificare
        .ExtractSchema True, fileNameSchema
        
        ' inizializzazione campi statici x ciascuna sezione
        .Initialize fileNameSchema
        
        ' mappatura fissa da programma
        .ReplaceValue fileNameSchema, Format$(mFromDate, "yyyy"), ColName:="ANNO_RIF"
        .ReplaceValue fileNameSchema, Format$(mFromDate, "mm"), ColName:="MESE_RIF"
        .ReplaceValue fileNameSchema, RP.AccountManager.Accounts.GetAccount(RP.LoggedAccountID).UserName, ColName:="USERNAME"
        
        .KeepAsMainDefault fileNameSchema
        
    End With

    ' è completo di percorso e nome
    FileName = RINI("Files", "Path", eINIWhere.LocalINI)

    DebugOperation = "Cancellazione file " & FileName
    DebugOut DebugOperation
    
    If FileExist(FileName) Then
        KillFile FileName, True
    End If
    
    
    IPutPay_Initialize = True
    
Pay_EXIT:

    Exit Function
    

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.ExeName & ".IPutPay_Initialize" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

Private Function IPutPay_InitializeDip(CustomMessage As String) As Boolean

Dim i           As Integer

Dim v           As String

Dim MC          As IMonthlyCalc

    On Error GoTo Pay_ERROR
    
    ' crea la collezione standard per le paghe
    DebugOperation = "Inizializzazione collezione oPayLU"

    mNominativo = oDip.Nominativo
    
    'imposta i campi fissi badge matricola e nominativo
    With fileSchema
    
        .Reset fileNameSchema
        
        ' impostare i campi mappati anagraficamente present <=> paghe da fare
        For i = 1 To maxMapAnag
        
            'Debug.Assert i <> 2
            
            v = oDip.GetAnagData(mapsAnag(i).PresentField, mFromDate)
            If v = "" And mapsAnag(i).DafaultValue <> "" Then
                v = mapsAnag(i).DafaultValue
            End If
            
            .ReplaceValue fileNameSchema, v, ColName:=mapsAnag(i).PayField
            
        Next
        
        .KeepAsMainDataDip fileNameSchema ' li memorizza
        
    End With
    
    Set monthSum = New cPays
    
    IPutPay_InitializeDip = True
    
Pay_EXIT:
    Exit Function
    
Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.ExeName & ".IPutPay_InitializeDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select
    
End Function


Private Sub IPutPay_InitializePerc(oPerc As IPercentageInfo)

    oPerc.DefaultCalculatedValue = Information

End Sub


Private Sub IPutPay_MakeSectionsKeys()
    
Dim Section     As CSection
Dim i           As Integer

    oSects.IniServerName = App.ExeName & ".ini"
    oSects.DescrizioneGenerale = App.FileDescription & " versione " & App.Major & "." & App.Minor & "." & App.Revision
    
    ' se il campo è Predefinito o con valore prefissato si imposta direttamente nel file di schema
    Set Section = oSects.Add("Mappatura Campi Paghe")
    Section.AddKey "Numero", StringText, ""
    For i = 1 To Val(Section.GetKey("Numero").Value)
        Section.AddKey "Campo Present " & i, StringText, "Consente di specificare il campo paghe da mappare"
        Section.AddKey "Campo Paghe " & i, StringText, "Consente di specificare il campo paghe minuti da mappare"
        Section.AddKey "Valore Predefinito " & i, StringText, ""
    Next

    ' causali giornaliere sul cartellino sommate mensilmente
    Set Section = oSects.Add("Mappatura Causali Paghe Da Cartellino")
    Section.AddKey "Numero", StringText, "Consente la mappatura tra il codice causale di present ed il campo paghe giornaliero da cartellino"
    For i = 1 To Val(Section.GetKey("Numero").Value)
        ' per ogni causale abbiamo il campo a ore e quello a minuti del file paghe
        Section.AddKey "Codice " & i, ListPresentCode, "Consente di specificare i codici delle causali da mappare giornalmente sul cartellino"
        Section.AddKey "Campo Paghe Ore " & i, StringText, "Consente di specificare il campo paghe ore/giorni da mappare"
        Section.AddKey "Campo Paghe Minuti " & i, StringText, "Consente di specificare il campo paghe minuti da mappare. NON usare se la causale è a giorni"
    Next
    
'    ' valori da prospetto mensile
'    Set section = oSects.Add("Mappatura Causali Paghe Da Prospetto Mensile")
'    section.AddKey "Numero", StringText, "Consente la mappatura tra il codice causale di present ed il campo paghe da prospetto mensile"
'    For i = 1 To Val(section.GetKey("Numero").Value)
'        ' per ogni campo del prospetto mensile presenze abbiamo il campo a ore e minuti del file paghe
'        section.AddKey "Campo " & i, StringText, "Consente di specificare il campo del prospetto mensile da mappare mensilmente"
'        section.AddKey "Campo Paghe Ore " & i, StringText, "Consente di specificare il campo paghe ore da mappare"
'        section.AddKey "Campo Paghe Minuti " & i, StringText, "Consente di specificare il campo paghe minuti da mappare"
'    Next
    
    ' Contatori mensili
    Set Section = oSects.Add("Mappatura Causali Paghe Da Contatori")
    Section.AddKey "Numero", StringText, "Consente la mappatura tra il codice contatore di present ed il campo paghe giornaliero da cartellino"
    For i = 1 To Val(Section.GetKey("Numero").Value)
        ' per ogni causale abbiamo il campo a ore e quello a minuti del file paghe
        Section.AddKey "Codice " & i, ListCounterCode, "Consente di specificare i contatori da mappare mensilmente"
        Section.AddKey "Campo Paghe Ore " & i, StringText, "Consente di specificare il campo paghe ore da mappare"
        Section.AddKey "Campo Paghe Minuti " & i, StringText, "Consente di specificare il campo paghe minuti da mappare"
        'Section.AddKey "Valore " & i, ComboBox, "Consente di specificare il valore da leggere", "0,1,2", "Parziale mensile,Residuo,Totale"
    Next
    
    ' nella versione 1.0.0 c'era questo che non serve più
    ' abilita passaggio pagato
    'Set Section = oSects.Add("Straordinario Pagato")
    'Section.AddKey "Abilita Passaggio", ComboBox, "Se non abililtato i dati contenenti lo straordinario pagato non verranno passati", "0,1", "No,Sì"
    
    ' da versione 1.0.1 lettura pagato da passare in paga da apposito contatore
    Set Section = oSects.Add("Straordinari Pagati")
    Section.AddKey "Numero Contatori", StringText, ""
    For i = 1 To Val(Section.GetKey("Numero Contatori").Value)
        ' per ogni contatore contenente il pagato abbiamo un corrispettivo pagabile
        ' il valore pagato è inserito nel campo manuale legge il parziale mensile
        Section.AddKey "Codice Contatore Contenente Quanitità Pagabile " & i, ListCounterCode, "Contatore " & i & " contenente la quantità pagabile"
        Section.AddKey "Codice Contatore Contenente Quanitità Pagata " & i, ListCounterCode, "Contatore " & i & " contenente la quantità pagata"
    Next
    
End Sub

Private Property Get IPutPay_MemoLogin() As Boolean
End Property

Private Function IPutPay_PreInitialize() As Boolean
    IPutPay_PreInitialize = True
End Function

Private Function IPutPay_PreInitializeDip(CustomMessage As String) As Boolean
    IPutPay_PreInitializeDip = True
End Function

Private Sub IPutPay_ReadINIParameters()

Dim i           As Integer

    FileName = RINI("Files", "Path", eINIWhere.LocalINI)
    
    ' il numero è 12
    maxMapAnag = Val(RINI("Mappatura Campi Paghe", "Numero"))
    If (maxMapAnag > 0) Then
        ReDim mapsAnag(1 To maxMapAnag) As mapAnag
        For i = 1 To maxMapAnag
            ' se il campo è Predefinito si imposta direttamente nello schema con valore prefissato si imposta direttamente nel file di schema
            mapsAnag(i).PresentField = RINI("Mappatura Campi Paghe", "Campo Present " & i)
            mapsAnag(i).PayField = RINI("Mappatura Campi Paghe", "Campo Paghe " & i)
            mapsAnag(i).DafaultValue = RINI("Mappatura Campi Paghe", "Valore Predefinito " & i)
            'mapsAnag(i).TypeField = RINI("Mappatura Campi Paghe", "Tipo Campo " & i)
        Next
    End If

    ' lettura causali da cartellino
    maxMapCart = Val(RINI("Mappatura Causali Paghe Da Cartellino", "Numero"))
    If (maxMapCart > 0) Then
        ReDim mapsCart(1 To maxMapCart) As mapC
        For i = 1 To maxMapCart
            mapsCart(i).Code = RINI("Mappatura Causali Paghe Da Cartellino", "Codice " & i)
            mapsCart(i).HourField = RINI("Mappatura Causali Paghe Da Cartellino", "Campo Paghe Ore " & i)
            mapsCart(i).MinField = RINI("Mappatura Causali Paghe Da Cartellino", "Campo Paghe Minuti " & i)
        Next
    End If
    
    ' lettura causali da prospetto mensile
    maxMapMonth = Val(RINI("Mappatura Causali Paghe Da Prospetto Mensile", "Numero"))
    If (maxMapMonth > 0) Then
        ReDim mapsMonth(1 To maxMapMonth) As mapM
        For i = 1 To maxMapMonth
            mapsMonth(i).Field = RINI("Mappatura Causali Paghe Da Prospetto Mensile", "Campo " & i)
            mapsMonth(i).HourField = RINI("Mappatura Causali Paghe Da Prospetto Mensile", "Campo Paghe Ore " & i)
            mapsMonth(i).MinField = RINI("Mappatura Causali Paghe Da Prospetto Mensile", "Campo Paghe Minuti " & i)
        Next
    End If
    
    ' configura eventuali contatori
    maxMapCount = Val(RINI("Mappatura Causali Paghe Da Contatori", "Numero"))
    If (maxMapCount > 0) Then
        ReDim mapsCount(1 To maxMapCount) As mapCounter
        For i = 1 To maxMapCount
            mapsCount(i).Code = RINI("Mappatura Causali Paghe Da Contatori", "Codice " & i)
            mapsCount(i).HourField = RINI("Mappatura Causali Paghe Da Contatori", "Campo Paghe Ore " & i)
            mapsCount(i).MinField = RINI("Mappatura Causali Paghe Da Contatori", "Campo Paghe Minuti " & i)
            'mapsCount(i).typeField = Val(RINI("Mappatura Causali Paghe Da Contatori", "Valore " & i))
        Next
    End If

    ' versione 1.0.1 non c'è +
    ' abilita passaggio straordinario pagato
    'overInPay = (Val(RINI("Straordinario Pagato", "Abilita Passaggio")) <> 0)
    
    mapsPayNumber = RINI("Straordinari Pagati", "Numero Contatori")
    If (mapsPayNumber > 0) Then
        ReDim mapsPay(1 To mapsPayNumber) As mapPay
        For i = 1 To maxMapCount
            mapsPay(i).paidCode = RINI("Straordinari Pagati", "Codice Contatore Contenente Quanitità Pagata " & i)
            mapsPay(i).toPayCode = RINI("Straordinari Pagati", "Codice Contatore Contenente Quanitità Pagabile " & i)
        Next
    End If

End Sub

Private Property Get IPutPay_SilentMode() As Boolean

End Property

Private Sub IPutPay_Terminate(WebCaller As Boolean)

Dim viewFile   As cFilePath

    On Error GoTo Pay_ERROR
    
    Set fileSchema = Nothing

    If Not IsFileNotExistOrEmpty(PresentUserLocalFolder() & fileNameSchema) Then
        FileCopy PresentUserLocalFolder() & fileNameSchema, FileName
    End If

    If fileProcessed Then
        Set viewFile = New cFilePath
        viewFile.OutputFileName FileName, False
        Set viewFile = Nothing
    End If
    
Pay_EXIT:
    Exit Sub

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.ExeName & ".IPutPay_Terminate" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
    Case vbAbort
        Resume Pay_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Sub


Private Function IPutPay_TerminateDip(eState As ePrctInfoStatus, CustomMessage As String) As ePrctInfoStatus

Dim itemSum         As cPay

Dim hour            As String
Dim minute          As String

Dim i               As Integer

Dim toPayValue      As Integer
Dim paidValue       As Integer

Dim nullFields()    As String

Dim MC              As IMonthlyCalc

Dim paidCounter     As String
Dim paidIndex       As Integer

Dim Value           As Integer

    On Error GoTo Pay_ERROR
    
    IPutPay_TerminateDip = eState
    
    ' mappatura fissa da programma dopo aver effettuato tutti i calcoli
    
    ' probabilmente va letto da una causale sul cartellino
    'CFG_GIORNI_MAT_FER                  NUMBER(3)           CFG Feriali - Maturati
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="CFG_GIORNI_MAT_FER"
    
    'CFG_GIORNI_MAT_FES                  NUMBER(3)           CFG Festivi - Maturati
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="CFG_GIORNI_MAT_FES"
    
    ' questi sono da verificare...
    'REC_GF_MAT                          NUMBER(3)           GF (Giornate Festive) -  Maturati
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="REC_GF_MAT"
    
    'REC_GNL_MAT                         NUMBER(3)           GNL (Giornate Non Lavorative) - Maturati
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="REC_GNL_MAT"
    
    'REC_GF_REC                          NUMBER(3)           GF (Giornate Festive) - Recuperati
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="REC_GF_REC"
    
    'REC_GNL_REC                         NUMBER(3)           GNL (Giornate Non Lavorative) - Recuperati
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="REC_GNL_REC"
    
    'PY1_REC_GF_REC               integer
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="PY1_REC_GF_REC"
    
    'PY1_REC_GNL_REC              integer
    'fileSchema.ReplaceValue fileNameSchema, 0, ColName:="PY1_REC_GNL_REC"
    
    'T_I                          Char width 1
    'fileSchema.ReplaceValue fileNameSchema, "0", ColName:="T_I"
    
    'T_I_DAL_TI                   date
    'fileSchema.ReplaceValue fileNameSchema, DateValue(Now), ColName:="T_I_DAL_TI"
    fileSchema.ReplaceValue fileNameSchema, "NULL", ColName:="T_I_DAL_TI"
    
    'T_I_AL_TI                    date
    'fileSchema.ReplaceValue fileNameSchema, DateValue(Now), ColName:="T_I_AL_TI"
    fileSchema.ReplaceValue fileNameSchema, "NULL", ColName:="T_I_AL_TI"
    
    'T_I_SDS_PROV                 Char width 4
    'fileSchema.ReplaceValue fileNameSchema, "", ColName:="T_I_SDS_PROV"
    
    'T_I_SDS_DEST                 Char width 4
    'fileSchema.ReplaceValue fileNameSchema, "", ColName:="T_I_SDS_DEST"
    
    'T_I_DESC                     Char width 32
    'fileSchema.ReplaceValue fileNameSchema, "", ColName:="T_I_DESC"
    
    'FLG_INSERT                   Char width 1
    'fileSchema.ReplaceValue fileNameSchema, "0", ColName:="FLG_INSERT"

    For i = 1 To maxMapCart
    
        hour = 0
        minute = 0
        
        ' bisogna lasciare tutti i campi del file paghe altrimenti non li inizializza
        ' rimangono i segnaposto e dà errore durante la scrittura
        If (mapsCart(i).Code <> "") Then
        
            If Not monthSum.isCodeLoaded(mapsCart(i).Code) Is Nothing Then
            
                Set itemSum = monthSum(mapsCart(i).Code)
                
                If Not (itemSum Is Nothing) Then
                    If itemSum.Cent Then
                        If InStr(itemSum.Value, ".") > 0 Then ' toglie il punto
                            hour = Left$(itemSum.Value, InStr(itemSum.Value, ".") - 1)
                            minute = Right$(itemSum.Value, InStr(itemSum.Value, ".") - 1)
                        ElseIf InStr(itemSum.Value, ",") > 0 Then ' toglie la virgola
                            hour = Left$(itemSum.Value, InStr(itemSum.Value, ",") - 1)
                            minute = Right$(itemSum.Value, InStr(itemSum.Value, ",") - 1)
                        End If
                    Else
                        If Not itemSum.IsDaily Then
                            ' in sessantesimi ore e minuti
                            hour = itemSum.Value \ 60
                            minute = itemSum.Value Mod 60
                        Else
                            ' a giorni
                            hour = itemSum.Value
                        End If
                    End If
                    
                    fileSchema.ReplaceValue fileNameSchema, hour, ColName:=mapsCart(i).HourField
                    
                    ' se il valore è giornaliero non occorre il secondo campo che normalmente viene utilizzato per i minuti
                    If Not itemSum.IsDaily Then
                        If (mapsCart(i).MinField <> "") Then
                            fileSchema.ReplaceValue fileNameSchema, minute, ColName:=mapsCart(i).MinField
                        End If
                    End If
                    
                End If
                
            End If
            
            IPutPay_TerminateDip = Successful
            
        End If
        
    Next
    
    ' lettura dati dal prospetto mensile
    With oDip.GetDataPeriod
        If (.RecordCount <> 0) And (.RecordCount <> -1) Then
            .MoveFirst
            For i = 0 To maxMapMonth
                On Error Resume Next
                
                Value = Val(.Fields(mapsMonth(i).Field).Value & "")
                
                If (Value <> 0) Then
                
                    hour = Value \ 60
                    minute = Value Mod 60
                    
                    fileSchema.ReplaceValue fileNameSchema, hour, ColName:=mapsMonth(i).HourField
                    fileSchema.ReplaceValue fileNameSchema, minute, ColName:=mapsMonth(i).MinField
                    
                    IPutPay_TerminateDip = Successful
                End If
                
                On Error GoTo Pay_ERROR
            Next
        End If
    End With
    


    ' lettura contatori mese corrente da passare in paga
    If (maxMapCount > 0) Then
        For i = 1 To maxMapCount
        
            ' legge il pagabile dal contatore
            'toPayValue = GetCounterValue(mapsCount(i).Code, mFromDate, mapsCount(i).typeField)
            toPayValue = GetCounterValue(mapsCount(i).Code, mFromDate, payType.totalPlusVariation)
            
            ' recupera il contatore del pagato
            paidIndex = GetCounterIndex(mapsCount(i).Code)
            If (paidIndex > 0) Then
                paidValue = GetCounterValue(mapsPay(paidIndex).paidCode, mFromDate, payType.total)
            End If
            
            ' passa il minore tra i due valori
            paidValue = Min(toPayValue, paidValue)

            If (paidValue > 0) Then
                    
                hour = paidValue \ 60
                minute = paidValue Mod 60
                        
                fileSchema.ReplaceValue fileNameSchema, hour, ColName:=mapsCount(i).HourField
                fileSchema.ReplaceValue fileNameSchema, minute, ColName:=mapsCount(i).MinField
                        
                If (paidValue > 0) Then
                    PutCounterValue mapsCount(i).Code, mFromDate, -paidValue
                End If
                
                IPutPay_TerminateDip = Successful
                
            End If
                        
        Next
    End If
        
    ' prima di scrivere la riga
    ' inizializza i campi rimasti vuoti
    nullFields = fileSchema.GetNullValue(fileNameSchema)
    If isStringArrayDim(nullFields) Then
        For i = 1 To UBound(nullFields)
            ' li inizializza
            Select Case fileSchema.GetFieldType(fileNameSchema, nullFields(i))
            Case Is = "number"
                fileSchema.ReplaceValue fileNameSchema, "0", ColName:=nullFields(i)
            Case Is = "date"
                fileSchema.ReplaceValue fileNameSchema, "NULL", ColName:=nullFields(i)
            Case Is = "string"
                fileSchema.ReplaceValue fileNameSchema, "", ColName:=nullFields(i)
            End Select
        Next
    End If

    fileSchema.WriteRow fileNameSchema
        
    Set monthSum = Nothing
    
    If (IPutPay_TerminateDip = Successful) Then
    
        oDip.WriteData
        
        ' calcolo mensile
        Set MC = RP.GetMonthlyCalc
        With MC
            .StartCalc oDip.DipID, mFromDate, mToDate, prpCMCont ' prpCMCont
            If Not .GetLastActionResult() Then
                MsgBox .GetLastActionErrorMessages, vbExclamation, "Attenzione"
            End If
        End With
    End If


Pay_EXIT:
    Exit Function

Pay_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.ExeName & ".IPutPay_TerminateDip" & vbCrLf & "(" & DebugOperation & ")" & vbCrLf & mNominativo)
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

    Dim oC As New cCalendar
    With oC
    
        .SelectorCount = 6
        .OnlyPath = False
        .CalendarMode = PreviousMonth
        Set .Icon = LoadResPicture(101, 1)
        .NumberFileOutPut = 1
        .FileName(1) = "Nome e percorso file paghe"
        .Fullmonth = True
        .DynamicPeriod = True
        Set IPutPay_ViewFullCalendar = .Run()
        StartDate = .StartDate
        EndDate = .EndDate
        '.AllowParametersEveryOne = True
    End With

    mFromDate = StartDate
    mToDate = EndDate

End Function

Private Property Get IPutPay_ViewLOG(LogList As String, Optional WaitSec As Integer = 15) As Boolean

End Property

Private Property Get IPutPay_ViewPercentage() As Boolean
    IPutPay_ViewPercentage = True
End Property

Private Function GetDailyFormatValue(D As Double) As String

    If InStr(D, ".") > 0 Then ' toglie il punto
        GetDailyFormatValue = Replace(Format(D, "000.0"), ".", "")
    ElseIf InStr(D, ",") > 0 Then ' toglie la virgola
        GetDailyFormatValue = Replace(Format(D, "000.0"), ",", "")
    Else
        'GetDailyFormatValue= Format(D, "0000000") & "00000"
        GetDailyFormatValue = Format(D, "0000")
    End If

End Function


Private Function GetCounterIndex(counterCode As String) As Integer
    
Dim i   As Integer

    For i = 1 To mapsPayNumber
        If mapsPay(i).toPayCode = counterCode Then
            GetCounterIndex = i
            Exit Function
        End If
    Next
    
End Function

Private Function GetCounterValue(counterCode As String, Data As Date, typeField As payType) As Integer

    Dim rsCounter   As ADODB.Recordset

    Set rsCounter = oDip.GetCounterFiltered(counterCode, Data, True)
    If Not (rsCounter Is Nothing) Then
        With rsCounter
            If Not .EOF Then
                ' parziale / residuo / totale
                If (typeField = partial) Then
                    ' parziale
                    GetCounterValue = Val(.Fields("Parziale") & "")
                ElseIf (typeField = residual) Then
                    ' residuo
                    GetCounterValue = Val(.Fields("Tetto") & "") + Val(.Fields("Riporto") & "") - Val(.Fields("Totale") & "")
                ElseIf (typeField = total) Then
                    ' totale
                    ' il totale non deve tener conto del valore inserito manualmente sul mese
                    ' perché se si lanciano + volte le paghe non deve tener conto del valore inserito precedentemente
                    GetCounterValue = Val(.Fields("Totale") & "")
                ElseIf (typeField = totalPlusVariation) Then
                    ' totale
                    ' il totale non deve tener conto del valore inserito manualmente sul mese
                    ' perché se si lanciano + volte le paghe non deve tener conto del valore inserito precedentemente
                    GetCounterValue = Val(.Fields("Totale") & "") + Abs(Val(.Fields("Variazione") & ""))
                End If
            End If
        End With
    End If

End Function

Private Sub PutCounterValue(counterCode As String, Data As Date, Value As Integer)

    Dim rsCounter   As ADODB.Recordset

    Set rsCounter = oDip.GetCounterFiltered(counterCode, Data, True)
    If Not (rsCounter Is Nothing) Then
        With rsCounter
            If Not .EOF Then
                ' aggiorna la variazione
                .Fields("Variazione") = Value
            End If
        End With
    End If

End Sub
