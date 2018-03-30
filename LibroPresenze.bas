Attribute VB_Name = "MainRun"
Option Explicit


Function Run() As Boolean


Dim rsDipSel                As ADODB.Recordset
Dim rsAnag                  As ADODB.Recordset
Dim rsOre                   As ADODB.Recordset
Dim rsGGLavRetr             As ADODB.Recordset

Dim ListCodes()             As String
Dim ListWorkedCodes()       As String
Dim ListPaidCodes()         As String

Dim WorkedDays              As Integer
Dim PaidDays                As Integer

Dim oS                      As CSection

Dim oD                      As CDatoGiornaliero

Dim oCR                     As cCrystalReport

Dim oH                      As cOra

Dim iP                      As IPercentageInfo

Dim Calculated              As ePrctInfoStatus

Dim dFromDate               As Date
Dim dToDate                 As Date
Dim CurrDate                As Date

Dim sMessage                As String
Dim DebugOperation          As String

Dim i                       As Integer

Dim bGo                     As Boolean

Dim PrinterDriverName       As String
Dim PrinterDeviceName       As String
Dim PrinterPort             As String
Dim PrinterOrientation      As Integer

    On Error GoTo Main_ERROR

    
    With oSects
    
        .IniServerName = App.EXEName & ".ini"
        .DescrizioneGenerale = App.FileDescription & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
        " ------------------------------------------------------------------------"
        
        Set oS = .Add("Codes", "Nella voce sottostante inseirer i simboli settimanali")
        oS.AddKey "Codici Da Stampare", ListPresentCode, "Codici da stampare"
        oS.AddKey "Codici Giorni Lavorati", ListPresentCode, "Codici validi ai fini delle giornate lavorate"
        'oS.AddKey "Codici Retribuiti In Negativo", ListPresentCode, "26 giorni meno le giornate complete con i codici specificati"
        
    End With
    
    oDip.SelectDipString = "Dipendenti.DipID,Dipendenti.Matricola,dipendenti.Badge,dipendenti.Cognome,dipendenti.Nome,dipendenti.CodAzienda"

    Dim oC As New cCalendar
    With oC
        .CalendarMode = PreviousMonth
        .Caption = "Stampa libro unico - Versione " & App.Major & "." & App.Minor & "." & App.Revision
        .EnablePrinter = True
        .NumberFileOutPut = 0
        .SelectorCount = 6
        .Fullmonth = True
        
        
        Set rsDipSel = .Run("(([Dipendenti] INNER JOIN [PeriodiServizio] ON [Dipendenti].[DipID]=[PeriodiServizio].[DipID]) left join Aziende on Dipendenti.CodAzienda = Aziende.CodAzienda)", "Aziende.Descrizione")
        
        PrinterDriverName = .PrinterDriverName
        PrinterDeviceName = .PrinterDeviceName
        PrinterPort = .PrinterPort
        'PrinterOrientation = .PrinterOrientation
        PrinterOrientation = PrinterOrientationConstants.cdlLandscape
        
        dFromDate = oC.StartDate
        dToDate = oC.EndDate
        
    End With
    Set oC = Nothing
    
    ListCodes = Split(RINI("Codes", "Codici Da Stampare"), ",")
    ListWorkedCodes = Split(RINI("Codes", "Codici Giorni Lavorati"), ",")
    
    ' i giorni retribuiti possono essere o 26 - determinati giorni completi di determinate causali
    ' oppure giornate con determinate causali
    ' al momento attivo la prima
    
    ListPaidCodes = Split(RINI("Codes", "Codici Retribuiti In Negativo"), ",")
        
    DebugOperation = "Generazione schema recordset disconnesso"
    Set rsAnag = New ADODB.Recordset
    With rsAnag
        .Fields.Append "Sede", adChar, 30
        .Fields.Append "Nominativo", adChar, 60
        .Fields.Append "Matricola", adChar, 20
        .Fields.Append "ID", adBigInt
        .Open
    End With

    Set rsOre = New ADODB.Recordset
    With rsOre
        .Fields.Append "ID", adBigInt
        .Fields.Append "Causale", adChar, 5
        .Fields.Append "GGOre", adSmallInt
        .Fields.Append "TotOre", adBigInt
        .Fields.Append "TypeCode", adSmallInt
        For i = 1 To 31
            .Fields.Append "Ore" & i, adBigInt
        Next
        .Fields.Append "Color", adBigInt
        .Open
    End With
    
    Set rsGGLavRetr = New ADODB.Recordset
    With rsGGLavRetr
        .Fields.Append "ID", adBigInt
        .Fields.Append "GGLav", adSmallInt
        .Fields.Append "GGRetr", adSmallInt
        .Open
    End With

    If Not rsDipSel Is Nothing Then
        If (rsDipSel.RecordCount = 0) Then
            GoTo Main_EXIT
        End If
    Else
        GoTo Main_EXIT
    End If
    
    CheckRefreshReport "Report", "Custom", pls.LocalFilesPath
    
    DebugOperation = "Inizializzazione oggetto oPercentageInfo"
    Set iP = New cPercentageInfo
    With iP
        .Caption = App.FileDescription & " - Versione " & App.Major & "." & App.Minor & "." & App.Revision
        .MaxValue = rsDipSel.RecordCount
        .EndMessage = "Elaborazione dati terminata"
        .DefaultCalculatedValue = Information
    End With
    
    With oDip
        .Initialize
        .FromDate = dFromDate
        .ToDate = dToDate
        .ReadsHours = True
        .ReadsDailyCalendar = True
        .ReadsDailyProfiles = True
    End With
   
    Do While Not rsDipSel.EOF
    
        DebugOperation = "Lettura dati DipID = " & rsDipSel!DipID
        With oDip
            
            .Change
            .DipID = rsDipSel!DipID
            .PreLoadAnagData False
            
            Select Case .DipState
                Case eDipState.DipNotFound
                    sMessage = "Dipendente con DipID = " & Format$(rsDipSel!DipID) & " non presente in anagrafica"
                Case eDipState.NotLoaded
                    sMessage = "Dipendente con DipID = " & Format$(rsDipSel!DipID) & " non impostato correttamente"
                Case eDipState.IsLoaded
                    sMessage = StringFormat("Nominativo {0} Badge {1}", .Nominativo, .Badge)
            End Select
            
        End With
        
        Calculated = Failed
        iP.ChangePercentage , sMessage, DefaultValue
        
        If (oDip.DipState <> eDipState.IsLoaded) Then
            GoTo SkipDip
        End If
        
        DebugOperation = "Dipendenti in forza" & sMessage
        If Not oDip.IsInForce(dFromDate, dToDate) Then
            GoTo SkipDip
        End If
        
        oDip.LoadDetailData
        
        WorkedDays = 0
        PaidDays = 26
        
        ' - ass cess
        If Format$(oDip.GetAnagData("DataAssunz", dFromDate), "mmyyyy") = Format$(dFromDate, "mmyyyy") Then
            ' se è assunto il primo del mese non conta
            PaidDays = PaidDays - (Val(Format$(oDip.GetAnagData("DataAssunz", dFromDate), "dd")) - 1)
        End If
        If Format$(oDip.GetAnagData("DataCessaz", dToDate), "mmyyyy") = Format$(dToDate, "mmyyyy") Then
            PaidDays = PaidDays - DaysOfMonth(dToDate) - Val(Format$(oDip.GetAnagData("DataCessaz", dToDate), "dd"))
        End If

        For CurrDate = dFromDate To dToDate
        
            Set oD = oDip.DailyData(CurrDate)

            ' calcolo giorni lavorati
            If (oD.ORE.Contains(ListWorkedCodes).Group.Sum > 0) Then
                WorkedDays = WorkedDays + 1
            End If
            
            ' calcolo giorni retribuiti
            If Not oD.Profili(1) Is Nothing Then
                If (TimeToMinute(oD.Profili(1).OreMassime) > 0) Then
                    If (oD.ORE.Contains(ListPaidCodes).Group.Sum >= (TimeToMinute(oD.Profili(1).OreMassime))) Then
                        PaidDays = PaidDays - 1
                    End If
                End If
            End If
                
            For Each oH In oD.ORE.Contains(ListCodes).Group.Sort
                With rsOre
                    .Filter = "ID = " & oDip.DipID & " AND Causale = '" & oH.Causale & "'"
                    If (.RecordCount = 0) Then
                        .AddNew
                        .Fields("ID") = oDip.DipID
                        .Fields("Causale") = oH.Causale
                        .Fields("Color") = oH.ColorCode
                        If oH.IsDailyPayType Then
                            .Fields("GGOre") = 1
                        Else
                            .Fields("GGOre") = 0
                        End If
                        If oH.Presenza And Not oH.Straordinaria Then
                             ' l'ordinario prima
                            .Fields("TypeCode") = 99
                        Else
                            ' poi il resto
                            .Fields("TypeCode") = oH.Grade
                        End If
                        For i = 1 To 31
                            .Fields("Ore" & i) = 0
                        Next
                    End If
                    If oH.IsDailyPayType Then
                        .Fields("Ore" & Day(CurrDate)) = oH.ResolvePayQuantity(oD.Due)
                    Else
                        .Fields("Ore" & Day(CurrDate)) = oH.GetMinutes
                    End If
                    Calculated = Successful
                End With
            Next
        Next
        
        If (Calculated = Successful) Then
            With rsAnag
                .AddNew
                .Fields("Nominativo") = oDip.Nominativo
                .Fields("Matricola") = oDip.Matricola
                '.Fields("Sede") = oDip.GetAnagData("CodAzienda", dFromDate)
                .Fields("Sede") = rsDipSel!Descrizione & ""
                .Fields("ID") = oDip.DipID
            End With
            With rsGGLavRetr
                .AddNew
                .Fields("ID") = oDip.DipID
                .Fields("GGLav") = WorkedDays
                .Fields("GGRetr") = PaidDays
            End With
            
            bGo = True
            
        End If

        
SkipDip:

        With iP
            .ChangeMessageOnFly sMessage, Calculated, , , True
            If .IsInterrupted Then
                Exit Do
            End If
        End With

        rsDipSel.MoveNext
        
    Loop
    
    DebugOperation = "Chiusura oggetti"
    If Not (iP Is Nothing) Then
        With iP
            .IsTerminated
            .Finalize
        End With
        Set iP = Nothing
    End If
    
    With rsOre
        .Filter = ""
        .Sort = "TypeCode DESC"
    End With
    
    If Not bGo Then
        MsgBox "Nessun dato da stampare", vbInformation
        GoTo Main_EXIT
    End If
    
    Set oCR = New cCrystalReport
    With oCR
    
        '.SetParameters pls.LocalFilesPath & App.EXEName & ".rpt", "Stampa libro presenze - Versione " & App.Major & "." & App.Minor & "." & App.Revision, PrinterOrientationConstants.cdlLandscape, , True
        .SetParameters pls.LocalFilesPath & App.EXEName & ".rpt", "Stampa libro presenze - Versione " & App.Major & "." & App.Minor & "." & App.Revision, PrinterOrientation
        If Not .IsReportFound Then
            GoTo Main_EXIT
        End If
        
        .AddDisconnectedRecordset rsAnag
        .AddDisconnectedRecordsetSubReport rsOre, "Subreport1"
        .AddDisconnectedRecordsetSubReport rsGGLavRetr, "Subreport2"
        
        .OpenReport PrinterDriverName, PrinterDeviceName, PrinterPort, PrinterOrientation
        
        For CurrDate = dFromDate To dToDate
            .SetFormula "G" & Format(CurrDate, "dd"), """" & Left(WeekdayName(Weekday(CurrDate, vbMonday)), 2) & """"
        Next
        
        .SetFormula "MeseAnno", """" & UCase(MonthName(Month(dFromDate))) & " " & Year(dFromDate) & """"
        
        For i = 31 To DaysOfMonth(dFromDate) + 1 Step -1
            .SetTextObject CStr(i), , True
        Next
        
        .Preview
        
    End With
    
    ' se arriva qui non ci sono stati errori
    Run = True
    
Main_EXIT:

    DebugOperation = "Chiusura oggetti"
    
    If Not (iP Is Nothing) Then
        iP.Finalize
        Set iP = Nothing
    End If
    
    If Not (rsDipSel Is Nothing) Then
        Set rsDipSel = Nothing
    End If
    
    If Not (rsAnag Is Nothing) Then
        Set rsAnag = Nothing
    End If
    
    If Not (rsOre Is Nothing) Then
        Set rsOre = Nothing
    End If
    
    If Not oCR Is Nothing Then
        Set oCR = Nothing
    End If
    
    Exit Function
    
Main_ERROR:

    Select Case CatchErr(Err.Description, Err.Number, App.EXEName & ".Main(" & DebugOperation & ")")
    Case vbAbort
        Resume Main_EXIT
    Case vbRetry
        Resume
    Case vbIgnore
        Resume Next
    End Select

End Function

