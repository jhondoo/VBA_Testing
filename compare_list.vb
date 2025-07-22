Sub TrouverIndicesAvecDico()
    Dim liste1 As Variant
    Dim liste2 As Variant
    Dim dict As Object
    Dim i As Long
    Dim indices As String

    ' Définir les listes
    liste1 = Array("ab", "cd", "ef")
    liste2 = Array("ab", "azer", "aze", "cd", "trza", "ef", "ytr", "hgr")

    ' Créer un dictionnaire pour liste1
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(liste1) To UBound(liste1)
        If Not dict.exists(liste1(i)) Then
            dict.Add liste1(i), True
        End If
    Next i

    ' Parcourir liste2 une seule fois
    For i = LBound(liste2) To UBound(liste2)
        If dict.exists(liste2(i)) Then
            indices = indices & i & ", "
        End If
    Next i

    ' Supprimer la dernière virgule
    If Len(indices) > 0 Then
        indices = Left(indices, Len(indices) - 2)
    End If

    MsgBox "Indices trouvés dans liste2 : " & indices
End Sub
