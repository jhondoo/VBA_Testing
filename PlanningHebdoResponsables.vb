Sub PlanningHebdoResponsables()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim ligneSource As Long, derLigne As Long
    Dim i As Long
    Dim dateFinMax As Date
    Dim currentDate As Date
    Dim dateDebut As Date, dateFin As Date
    Dim ligneEcriture As Long
    Dim colSemaine As Integer

    ' Référence à la feuille contenant les données
    Set wsSource = ThisWorkbook.Sheets("data_alice") ' adapte si besoin

    ' Créer une nouvelle feuille pour le planning
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("gantt").Delete ' on écrase si existe
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = "gantt"

    ' Date actuelle
    currentDate = Date

    ' Trouver la date fin maximale
    derLigne = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    dateFinMax = wsSource.Cells(2, 11).Value

    For i = 2 To derLigne
        If wsSource.Cells(i, 11).Value > dateFinMax Then
            dateFinMax = wsSource.Cells(i, 11).Value
        End If
    Next i

    ' Calcul du nombre de semaines entre maintenant et la date fin maximale
    Dim nbSemaines As Long
    nbSemaines = DateDiff("ww", currentDate, dateFinMax, vbMonday, vbFirstFourDays)

    ' Écrire en-tête des colonnes
    wsDest.Cells(1, 1).Value = "Responsable"
    wsDest.Cells(1, 2).Value = "Projet"
    colSemaine = 3

    Dim j As Long
    For j = 0 To nbSemaines
        Dim dSemaine As Date
        dSemaine = DateAdd("ww", j, currentDate)
        wsDest.Cells(1, colSemaine).Value = "Semaine " & Format(dSemaine, "ww") & " (" & Year(dSemaine) & ")"
        colSemaine = colSemaine + 1
    Next j

    ' Initialisation
    ligneEcriture = 2

    ' Une boucle sur les données source
    For i = 2 To derLigne
        Dim responsable As String, projet As String
        responsable = wsSource.Cells(i, 10).Value
        projet = wsSource.Cells(i, 2).Value
        dateDebut = wsSource.Cells(i, 13).Value
        dateFin = wsSource.Cells(i, 11).Value

        ' Ligne de base
        wsDest.Cells(ligneEcriture, 1).Value = responsable
        wsDest.Cells(ligneEcriture, 2).Value = projet

        ' Marquer les semaines concernées
        For j = 0 To nbSemaines
            Dim dateSemaine As Date
            dateSemaine = DateAdd("ww", j, currentDate)

            If dateFin >= dateSemaine And dateDebut <= dateSemaine Then
                wsDest.Cells(ligneEcriture, j + 3).Value = "✓"
            End If
        Next j

        ligneEcriture = ligneEcriture + 1
    Next i

    ' Formatage
    wsDest.Columns.AutoFit

    MsgBox "Planning généré avec succès !"

End Sub
