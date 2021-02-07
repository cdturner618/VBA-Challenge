Attribute VB_Name = "Module1"
Option Explicit

Sub main()
'create variables
    Dim WorksheetName As String
    Dim Ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim totalstockvolume As Double
    Dim lastrow As Long
    Dim lastrow2 As Long
    Dim greatest_increase(1) As Variant
    Dim greatest_decrease(1) As Variant
    Dim greatest_volume(1) As Variant
    Dim startprice As Double
    Dim endprice As Double
    Dim ws As Worksheet
    Dim i As Long
    
'loop through all sheets
For Each ws In Worksheets
 
    ' Initial values
    greatest_increase(0) = -10000000000#
    greatest_decrease(0) = 10000000000#
    greatest_volume(0) = -1000000000000#
    Ticker = ""
    
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    ws.Range("I2:L" & lastrow).ClearContents
    

    'Worksheet Name
    'Worksheet = ws.Name
    'MsgBox ("WorksheetName")
    
    'Determine Last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (lastrow)
    
    'Add headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Cells(1, 12).Value = "Total stock Volume"
    
    'looping for ticker
    
    For i = 2 To lastrow

        If ws.Cells(i, 1).Value <> Ticker Then
           ' End of group things
            If i > 2 Then
                lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
                ws.Cells(lastrow2 + 1, 9).Value = Ticker
                ws.Cells(lastrow2 + 1, 12).Value = totalstockvolume
                endprice = ws.Cells(i - 1, 6).Value
                yearly_change = endprice - startprice
                ws.Cells(lastrow2 + 1, 10).Value = yearly_change
                If startprice = 0 Then
                    percent_change = 1E+16
                Else
                    percent_change = (yearly_change) / startprice
                    
                End If
                ws.Cells(lastrow2 + 1, 11).Value = percent_change
                
                If yearly_change > greatest_increase(0) Then
                    greatest_increase(0) = yearly_change
                    greatest_increase(1) = Ticker
                    
                End If
                
                'dido for greatest decrease and greatest volume
        
            End If
            'Beginning of group things
            Ticker = ws.Cells(i, 1).Value
            totalstockvolume = 0
            startprice = ws.Cells(i, 6).Value
     
        End If
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
            
        
    Next i
    'write out greatest change increase decrease
    
    'color formmating
    For i = 2 To lastrow
         If ws.Cells(i, 10) >= 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10) < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    Next i
    
Next ws
End Sub
