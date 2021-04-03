Attribute VB_Name = "Module3"
Sub StockAnalyzer()

Dim ws As Worksheet
Dim summarytable As Boolean
    summarytable = False
Dim summarytablerow As Long
Dim lastrow As Long
Dim ticker As String
Dim stockvolume As Double
Dim startprice As Double
Dim endprice As Double
Dim deltaprice As Double
Dim deltapercent As Double
Dim i As Long
summarytablerow = 2
ticker = " "
stockvolume = 0
startprice = 0
endprice = 0
deltaprice = 0
deltapercent = 0

    For Each ws In Worksheets
        If summarytable Then
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
        Else
            summarytable = True
        End If
        summarytablerow = 2
        startprice = ws.Cells(2, 3).Value
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
            stockvolume = 0
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                endprice = ws.Cells(i, 6).Value
                deltaprice = endprice - startprice
                If startprice <> 0 Then
                    deltapercent = (deltaprice / startprice) * 100
                Else
                    MsgBox ("Zero in an Opening Stock Volume space.")
                End If
            stockvolume = stockvolume + ws.Cells(i, 7).Value
            ws.Range("I" & summarytablerow).Value = ticker
            ws.Range("J" & summarytablerow).Value = deltaprice
                If (deltaprice > 0) Then
                    ws.Range("J" & summarytablerow).Interior.ColorIndex = 4
                ElseIf (deltaprice <= 0) Then
                    ws.Range("J" & summarytablerow).Interior.ColorIndex = 3
                End If
            ws.Range("K" & summarytablerow).Value = (CStr(deltapercent) & "%")
            ws.Range("L" & summarytablerow).Value = stockvolume
            summarytablerow = summarytablerow + 1
            deltaprice = 0
            endprice = 0
            startprice = ws.Cells(i + 1, 3).Value
            Else
                stockvolume = stockvolume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub
