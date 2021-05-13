Sub DQAnalysis()

rowStart = 2
rowEnd = 3013
totalVolume = 0

Worksheets("2018").Activate
 
For i = rowStart To rowEnd

    'increase totalVolume if ticker is "DQ"
    
    If Cells(i, 1).Value = "DQ" Then
        totalVolume = totalVolume + Cells(i, 8).Value
        
    End If

Next i

'MsgBox (totalVolume)

Worksheets("DQ Analysis").Activate
Cells(4, 1).Value = 2018
Cells(4, 2).Value = totalVolume

End Sub

