Sub CopierColonnesUltraRapide()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim headers As Variant
    Dim srcData As Variant
    Dim result() As Variant
    
    Dim colMap As Object
    Dim i As Long, j As Long
    Dim lastRow As Long, lastCol As Long
    Dim destCol As Long
    
    Set wsSrc = Sheets("Feuil1")
    Set wsDst = Sheets("Feuil2")
    
    ' Liste des entêtes à copier
    headers = Array("Nom", "Age", "Ville")
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    
    ' Charger toutes les données en mémoire
    srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value
    
    ' Dictionnaire pour trouver les colonnes rapidement
    Set colMap = CreateObject("Scripting.Dictionary")
    
    For j = 1 To lastCol
        colMap(srcData(1, j)) = j
    Next j
    
    ' Tableau résultat
    ReDim result(1 To lastRow, 1 To UBound(headers) + 1)
    
    ' Copier colonnes
    For j = 0 To UBound(headers)
        
        If colMap.exists(headers(j)) Then
            
            For i = 1 To lastRow
                result(i, j + 1) = srcData(i, colMap(headers(j)))
            Next i
            
        End If
        
    Next j
    
    ' Ecriture en une seule fois
    wsDst.Range("A1").Resize(lastRow, UBound(headers) + 1).Value = result

End Sub
