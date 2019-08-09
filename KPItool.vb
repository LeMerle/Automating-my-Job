Sub loopNAppendKPI()
    'Uses a case statement to add a marker to the end of a report so I can identify ones I work on.
    Dim i As Long
    Dim LR  As Long
    Dim Mark As String
    
    Mark = " - N"
    
    LR = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 3 To LR
        Select Case Range("A" & i).Value
        Case "AARONSFUR", "AARONSFURN", "AMAZONFBA", "BBB", "BBB01", "BEALLS", "BEALLSAPL", "BEALLSAPLDEPT", _
        "BEALLSOUTART", "DLFBA", "DLS", "HGART", "HGARTPOE", "HOMEGOODS", "HOMEGOODSWHS", "HSN", "HSNBASIC" _
        , "HSNIMP", "KMART", "MARMAXX", "MARSHALLS", "MarshallsCan", "MARSHAP", "PETCO", "PETFBA", "PETSMART", _
        "SEARS", "SNSPET", "SYNCHTECH", "TJMAXX", "TJMAXXAP", "TJMAXXAP01", "TJMAXXAPART", "TJXAUST", "TKMAXX", _
        "TKMAXXART", "VFAPL", "Zulily", "ZULINC"
        Range("V" & i) = Mark
        End Select
    Next i
    
End Sub

Sub TallyUp()
    Dim i As Long
    Dim LR  As Long
    Dim orders As Long
    Dim containers As Long
    
    orders = 0
    containers = 0

    LR = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 3 To LR
        If Range("V" & i).Value = " - N" Then
        containers = containers + Range("G" & i).Value
        orders = orders + Range("E" & i).Value
        End If
    Next i
    
    Range("W3").Value = containers
    Range("W4").Value = orders
End Sub

Sub TallyUpIndividuals()
    'Uses a case statement to add a marker to the end of a report so I can identify ones I work on.
    Dim i As Long
    Dim LR  As Long
    Dim Mark As String
    
    Mark = " - N"
    
    LR = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 3 To LR
        Select Case Range("A" & i).Value
        Case "AARONSFUR", "AARONSFURN", "AMAZONFBA", "BBB", "BBB01", "BEALLS", "BEALLSAPL", "BEALLSAPLDEPT", _
        "BEALLSOUTART", "DLFBA", "DLS", "HGART", "HGARTPOE", "HOMEGOODS", "HOMEGOODSWHS", "HSN", "HSNBASIC" _
        , "HSNIMP", "KMART", "MARMAXX", "MARSHALLS", "MarshallsCan", "MARSHAP", "PETCO", "PETFBA", "PETSMART", _
        "SEARS", "SNSPET", "SYNCHTECH", "TJMAXX", "TJMAXXAP", "TJMAXXAP01", "TJMAXXAPART", "TJXAUST", "TKMAXX", _
        "TKMAXXART", "VFAPL", "Zulily", "ZULINC"
        Range("V" & i) = Mark
        End Select
    Next i
    
End Sub
