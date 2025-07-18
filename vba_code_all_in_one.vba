'=============================
' Sub: CalculateNumberOfMovements
'=============================

Sub CalculateNumberOfMovements()
     dim i as long
     dim j as integer
     j=0
     for i=2 to 4670
          if worksheets("Position Movements").cells(i,1).value=1 then
     j=j+1
             worksheets("sheet1").cells(j,6).value=worksheets("Position Movements").cells(i-1,1).value
          end if
     next
End Sub

'=============================
' Sub: CountDuplicates
'=============================
    
Sub CountDuplicates()
    Dim i As Long
    Dim j As Long
    Dim z As Long
  For i = 2 To 1686
     z = 0
     For j = 2 To 4670
        If Worksheets("sheet1").Cells(i, 1).Value = Worksheets("Position Movements").Cells(j, 2).Value Then
        z = z + 1
        End If
     Next j
       Worksheets("sheet1").Cells(i, 6).Value = z
  Next i
End Sub

'=============================
' Sub: DetermineStatusChangeAndLastStatus
'=============================

 Sub DetermineStatusChangeAndLastStatus()
    Dim i As Long
    Dim j As Long
    Dim d As Long
     Dim z As Long
     Dim k As Long
     Dim w As String
     z = 1
  For i = 2 To 1686
     d = Worksheets("sheet1").Cells(i, 6).Value
     w = Worksheets("Position Movements").Cells(z + 1, 34).Value
       For j = z + 1 To d + z
          If w = Worksheets("Position Movements").Cells(j, 34).Value Then
             k = 0
             Else
             k = 1
             Exit For
          End If
       Next j
       Worksheets("sheet1").Cells(i, 13).Value = k
       z = d + z
       Worksheets("sheet1").Cells(i, 14).Value = Worksheets("Position Movements").Cells(z, 34).Value
  Next i

End Sub

'=============================
' Sub: FindTerminated
'=============================

Sub findterminated()
Dim i As Long
Dim j As Long
Dim k As Long
j = 2
For i = 2 To 1686
    If Worksheets("sheet1").Cells(i, 5).Value = "Terminated" Then
         Worksheets("sheet2").Cells(j, 1).Value = Worksheets("sheet1").Cells(i, 1).Value
         j = j + 1
    End If
Next
End Sub

'=============================
' Sub: FindTerminatedMovements
'=============================

Sub findforterminated()
Dim i As Long
Dim j As Integer
j = 0
For i = 2 To 447
   For j = 2 To 1686
     If Worksheets("sheet2").Cells(i, 1).Value = Worksheets("sheet1").Cells(j, 1).Value Then
        Worksheets("sheet2").Cells(i, 2).Value = Worksheets("sheet1").Cells(j, 6).Value
     End If
   Next j
Next
End Sub

'=============================
' Sub: FindTerminationDate
'=============================

Sub findforterminated()
     Dim i As Long
     Dim j As Long
     Dim k As Long
     Dim z As Long
     z = 1
     k = 1
     For i = 2 To 447
        For j = 1 To 4670
          If Worksheets("sheet2").Cells(i, 1).Value = Worksheets("Position Movements").Cells(j, 2).Value Then
             k = Worksheets("sheet2").Cells(i, 2).Value
             z = j + k - 1
             Worksheets("sheet2").Cells(i, 3).Value = Worksheets("Position Movements").Cells(z, 12).Value
             Exit For
          End If
        Next j
     Next
End Sub

'=============================
' Sub: CountGRN_GRM_Starters
'=============================

Sub CountGN()
  Dim i As Long
  Dim j As Long
  For i = 2 To 143
     For j = 2 To 447
        If Worksheets("sheet3").Cells(i, 1).Value = Worksheets("sheet2").Cells(j, 1).Value Then
               Worksheets("sheet2").Cells(j, 5).Value = 1
        End If
     Next j
  Next i
End Sub
