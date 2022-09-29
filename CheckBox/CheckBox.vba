Sub LinkChecks()
'Update by Extendoffice
Dim xCB
Dim xCChar
i = 2 'Linha que começa primeiro checkbox
xCChar = "B" 'Coluna em que sera associado o checkbox
For Each xCB In ActiveSheet.CheckBoxes
If xCB.Value = 1 Then
    Cells(i, xCChar).Value = True
Else
    Cells(i, xCChar).Value = False
End If
xCB.LinkedCell = Cells(i, xCChar).Address
i = i + 1
Next xCB
End Sub
