 Sub deleteExportReportColumns()
    ' delete columns D, F, G, J-K, N-AA
    
    Range("D:D,F:F,G:G,J:K,N:AA").Delete
 End Sub
 
  Sub deleteContainerShortageColumns()
    ' delete columns B-I, K-L, N-Y, AA-AC
    
    Range("B:I,K:L,N:Y,AA:AC").Delete
 End Sub

Sub neuterZeros()
    Dim i As Long
    Dim LR As Long
    LR = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 3 To LR
        If Range("C" & i) = 0 Then
             Range("A" & i & ":E" & LR).Delete
             Exit For
        End If
    Next i
    
End Sub

Sub addColumnsNFormatting()
    
    Dim LR As Long
    Dim i As Long
    
    LR = Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("E2") = "Ctns Short"
    Range("F2") = "Container No"
    Range("A2:F2").Interior.ColorIndex = 5
    Range("A2:F2").Font.Bold = True
    Range("A2:F2").Font.ColorIndex = 2
    Range("A2:F" & LR).Borders.LineStyle = xlContinuous
    Range("A2:F" & LR).Columns.AutoFit
    
    For i = 3 To LR
        Range("E" & i).Formula = "=(C" & i & "/d" & i & ")"
    Next i
    
End Sub

Sub loopNAppend()
    ' loop until you find a blank cell
    ' while looping, check contents of G for DLFBA, AMAZONFBA & PETFBA (column B)
    ' check contents of E & F for BEALLS And DLS (column B)
    ' BEALLS is 250 lbs/25 ctns
    ' DLS is 150 lbs/15 ctns
    Dim i As Long
    Dim LR  As Long
    Dim LTL As String
    Dim UPSG As String
    Dim FX As String
    
    LTL = " - Ships LTL"
    UPSG = " - Ships UPSG"
    FXG = " - Ships FXG"
    
    
    LR = Cells(Rows.Count, "A").End(xlUp).Row
    ' Debug.Print LR
    ' for i < LR
    ' check B for the names I want
    ' check G or E & F
    If Range("B3").Value = "DLFBA" Or Range("B3").Value = "PETFBA" Or Range("B3").Value = "AMAZONFBA" Then
            For i = 3 To LR
                If Range("G" & i) >= 44 Then
                    Range("H" & i) = Range("H" & i) & LTL
                End If
                If Range("G" & i) < 44 Then
                    Range("H" & i) = Range("H" & i) & UPSG
                End If
            Next i
    End If
    If Range("B3").Value = "DLS" Then
        For i = 3 To LR
        If Range("E" & i) <= 15 And Range("F" & i) <= 150 Then
            Range("H" & i) = Range("H" & i) & FXG
        End If
        If Range("E" & i) > 15 Or Range("F" & i) > 150 Then
            Range("H" & i) = Range("H" & i) & LTL
        End If
        Next i
    End If
    If Range("B3").Value = "BEALLS" Then
        For i = 3 To LR
        If Range("E" & i) <= 25 And Range("F" & i) <= 250 Then
            Range("H" & i) = Range("H" & i) & FXG
        End If
        If Range("E" & i) > 25 Or Range("F" & i) > 250 Then
            Range("H" & i) = Range("H" & i) & LTL
        End If
        Next i
    End If
    
    Range("A2:H" & LR).Copy
    
End Sub


Sub containerShortage()

    deleteContainerShortageColumns
    neuterZeros
    addColumnsNFormatting
    
End Sub

Sub exportReport()

    deleteExportReportColumns
    loopNAppend

End Sub

