Attribute VB_Name = "Module1"
Sub stock()

Dim Lrow As Long
Dim Srow As Integer
Dim changeprice As Double
Dim openprice As Double
Dim closingprice As Double
Dim Totalvol As Double
Dim Pchange As Double
Dim I As Long
Dim pctr As Integer
Dim MaxPchange As Double      ' Max Percentage'
Dim MinPchange As Double      ' Min Percentage'
Dim MaxTVol As Double         ' Max Total Vol'
Dim TckrI As String           'ticker for greatest increase
Dim TckrD As String           'ticker for greatest decrease
Dim TckrV As String           'ticker for max total volume
Dim Ws As Worksheet



' Display the Headers
For Each Ws In Worksheets

Ws.Range("L1") = " Ticker"
Ws.Range("M1") = "Yearly Change"
Ws.Range("N1") = "Percentage Change"
Ws.Range("O1") = "Total Volume"
'Ws.Range("Q1") = "Close Price" 'display for testing
'Ws.Range("R1") = "Open Price"  'display for testing


'Counter for Last row and Sequence Row where we will transfer the result
Lrow = Ws.Cells(2, 1).End(xlDown).Row
Srow = 2


openprice = Ws.Cells(2, 3).Value   'Initial Value of open price




 For I = Srow To Lrow
   
  If Ws.Cells(I, 1) = Ws.Cells(I + 1, 1) Then
  
  Totalvol = Totalvol + Ws.Cells(I, 7).Value     'Add Volume per Ticker
  
  Else
  
  closingprice = Ws.Cells(I, 6).Value            'value of closing price of a ticker
  
  changeprice = closingprice - openprice      'Computation for change price
  
  If openprice > 0 Then
    Pchange = (changeprice / openprice)       'Computation for Percent Change
  Else
    Pchange = 0
  End If
   
  Totalvol = Totalvol + Ws.Cells(I, 7).Value     'Add Volume per Ticker

 Ws.Range("L" & Srow).Value = Ws.Cells(I, 1)        'Display value for Ticker
 Ws.Range("M" & Srow).Value = changeprice        'Display value for yearly price change
  
  If changeprice < 0 Then
   Ws.Range("M" & Srow).Interior.ColorIndex = 3  'fill row with colour red if price change is negative
   
  Else
  Ws.Range("M" & Srow).Interior.ColorIndex = 4   'fill row with color green if proce change is positive
  End If
 Ws.Range("N" & Srow).Value = Format(Pchange, "0.00%")  'Display and Format percent change data
 Ws.Range("O" & Srow).Value = Totalvol                  'Display Total Volume data
 
 
' Ws.Range("Q" & Srow).Value = closingprice 'use for testing
' Ws.Range("R" & Srow).Value = openprice    'use for testing
  
 
 openprice = Ws.Cells(I + 1, 3).Value          'get the next open price for a new ticker
 Srow = Srow + 1                            'increment the row where to display the unique ticker
Totalvol = 0                                'reset the volume to be able to compute for the new ticker
 End If

Next I                                      'looping for I


'Below code is to get and display the Greatest % increase&decrease and total Volume
 
 MaxPchange = 0
 MaxTVol = 0
 MinPchange = 0
For pctr = 2 To Srow - 1
    If Ws.Cells(pctr, 14) > MaxPchange Then
       MaxPchange = Ws.Cells(pctr, 14)
       TckrI = Ws.Cells(pctr, 12).Value
    End If
    If Ws.Cells(pctr, 14) < MinPchange Then     'Ws.Cells(pctr,14) is column N for yearly percent change and it will compares each row
       MinPchange = Ws.Cells(pctr, 14)
       TckrD = Ws.Cells(pctr, 12).Value
    End If
    If Ws.Cells(pctr, 15) > MaxTVol Then        'Ws.Cells(pctr,15) is column O for total volume and it will compares each row
       MaxTVol = Ws.Cells(pctr, 15)
       TckrV = Ws.Cells(pctr, 12).Value
    End If
Next pctr                                    'looping for each row
 Ws.Range("Q2").Value = "Greatest % increase"
 Ws.Range("Q3").Value = "Greatest % decrease"
 Ws.Range("Q4").Value = "Greatest total volume"
 Ws.Range("R1").Value = "Ticker"
 Ws.Range("R2").Value = TckrI
 Ws.Range("R3").Value = TckrD
 Ws.Range("R4").Value = TckrV
 Ws.Range("S1").Value = "Value"
 Ws.Range("S2").Value = Format(MaxPchange, "0.00%")
 Ws.Range("S3").Value = Format(MinPchange, "0.00%")
 Ws.Range("S4").Value = MaxTVol
 
Next Ws
 
End Sub

