Sub TrierParMinDate_QuickSort()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' =========================
    ' PARAMÈTRES
    ' =========================
    Dim col1 As Long: col1 = 1
    Dim col2 As Long: col2 = 2
    Dim col3 As Long: col3 = 3
    Dim col4 As Long: col4 = 4
    
    Dim firstRow As Long: firstRow = 2
    ' =========================
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col1).End(xlUp).Row
    
    Dim minVals() As Double
    Dim idx() As Long
    Dim i As Long
    
    ReDim minVals(firstRow To lastRow)
    ReDim idx(firstRow To lastRow)
    
    ' 🔎 Calcul des min + index d'origine
    For i = firstRow To lastRow
        minVals(i) = GetMinDate(ws, i, col1, col2, col3, col4)
        idx(i) = i ' ordre initial
    Next i
    
    ' ⚡ QuickSort
    QuickSort minVals, idx, firstRow, lastRow
    
    ' 🔄 Réorganisation des lignes avec INSERT (formats conservés)
    ReorderRows ws, idx, firstRow, lastRow

End Sub


' =========================
' 🔽 QUICK SORT (stable via index)
' =========================
Sub QuickSort(ByRef arr() As Double, ByRef idx() As Long, ByVal first As Long, ByVal last As Long)

    Dim i As Long, j As Long
    Dim pivot As Double
    Dim pivotIdx As Long
    
    i = first
    j = last
    
    pivot = arr((first + last) \ 2)
    pivotIdx = idx((first + last) \ 2)
    
    Do While i <= j
        
        Do While Compare(arr(i), idx(i), pivot, pivotIdx) < 0
            i = i + 1
        Loop
        
        Do While Compare(arr(j), idx(j), pivot, pivotIdx) > 0
            j = j - 1
        Loop
        
        If i <= j Then
            SwapDouble arr(i), arr(j)
            SwapLong idx(i), idx(j)
            i = i + 1
            j = j - 1
        End If
        
    Loop
    
    If first < j Then QuickSort arr, idx, first, j
    If i < last Then QuickSort arr, idx, i, last

End Sub


' 🔁 Comparaison avec stabilité (ordre d'origine si égalité)
Function Compare(val1 As Double, idx1 As Long, val2 As Double, idx2 As Long) As Long
    
    If val1 < val2 Then
        Compare = -1
    ElseIf val1 > val2 Then
        Compare = 1
    Else
        ' égalité → garder ordre initial
        If idx1 < idx2 Then
            Compare = -1
        Else
            Compare = 1
        End If
    End If

End Function


' =========================
' 🔽 MIN DATE (ignore vides)
' =========================
Function GetMinDate(ws As Worksheet, rowNum As Long, c1 As Long, c2 As Long, c3 As Long, c4 As Long) As Double
    
    Dim vals(1 To 4) As Variant
    Dim i As Long
    Dim minVal As Double
    Dim found As Boolean
    
    vals(1) = ws.Cells(rowNum, c1).Value
    vals(2) = ws.Cells(rowNum, c2).Value
    vals(3) = ws.Cells(rowNum, c3).Value
    vals(4) = ws.Cells(rowNum, c4).Value
    
    found = False
    
    For i = 1 To 4
        
        If Not IsEmpty(vals(i)) Then
            If IsDate(vals(i)) Then
                
                If Not found Then
                    minVal = CDbl(vals(i))
                    found = True
                ElseIf CDbl(vals(i)) < minVal Then
                    minVal = CDbl(vals(i))
                End If
                
            End If
        End If
        
    Next i
    
    If Not found Then
        GetMinDate = 1E+99 ' pas de date → fin
    Else
        GetMinDate = minVal
    End If

End Function


' =========================
' 🔽 RÉORGANISATION AVEC INSERT (100% formats conservés)
' =========================
Sub ReorderRows(ws As Worksheet, idx() As Long, firstRow As Long, lastRow As Long)

    Dim i As Long, targetRow As Long
    
    For i = firstRow To lastRow
        
        targetRow = idx(i)
        
        If targetRow <> i Then
            
            ws.Rows(targetRow).Cut
            ws.Rows(i).Insert Shift:=xlDown
            
            ' Mettre à jour les index après déplacement
            UpdateIndex idx, i, targetRow
            
        End If
        
    Next i

End Sub


Sub UpdateIndex(ByRef idx() As Long, ByVal newPos As Long, ByVal oldPos As Long)

    Dim i As Long
    
    For i = LBound(idx) To UBound(idx)
        
        If idx(i) = newPos Then
            idx(i) = oldPos
        ElseIf idx(i) > newPos And idx(i) <= oldPos Then
            idx(i) = idx(i) - 1
        End If
        
    Next i

End Sub


' =========================
' 🔽 UTILITAIRES
' =========================
Sub SwapDouble(ByRef a As Double, ByRef b As Double)
    Dim t As Double
    t = a: a = b: b = t
End Sub

Sub SwapLong(ByRef a As Long, ByRef b As Long)
    Dim t As Long
    t = a: a = b: b = t
End Sub
