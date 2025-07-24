Sub PlanningHebdoResponsables()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim ligneSource As Long, derLigne As Long
    Dim semaineCourante As Integer, semaineMax As Integer
    Dim i As Long, semaine As Integer
    Dim dateFinMax As Date
    Dim currentDate As Date
    Dim dateDebut As Date, dateFin As Date
    Dim dictResponsables As Object
    Dim ligneEcriture As Long
    Dim projetCle As String
    Dim colSemaine As Integer

    ' Référence à la feuille contenant les données
    Set wsSource = ThisWorkbook.Sheets("Feuil1") ' adapte si besoin

    ' Créer une nouvelle feuille pour le planning
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Planning").Delete ' on écrase si existe
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = "Planning"

    ' Date actuelle et semaine courante
    currentDate = Date
    semaineCourante = Application.WorksheetFunction.WeekNum(currentDate, 21)

    ' Trouver la date fin maximale
    derLigne = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    dateFinMax = wsSource.Cells(2, 4).Value

    For i = 2 To derLigne
        If wsSource.Cells(i, 4).Value > dateFinMax Then
            dateFinMax = wsSource.Cells(i, 4).Value
        End If
    Next i

    semaineMax = Application.WorksheetFunction.WeekNum(dateFinMax, 21)

    ' Écrire en-tête des colonnes
    wsDest.Cells(1, 1).Value = "Responsable"
    wsDest.Cells(1, 2).Value = "Projet"
    colSemaine = 3

    For semaine = semaineCourante To semaineMax
        wsDest.Cells(1, colSemaine).Value = "Semaine " & semaine
        colSemaine = colSemaine + 1
    Next semaine

    ' Initialisation
    ligneEcriture = 2

    ' Une boucle sur les données source
    For i = 2 To derLigne
        Dim responsable As String, projet As String
        responsable = wsSource.Cells(i, 1).Value
        projet = wsSource.Cells(i, 2).Value
        dateDebut = wsSource.Cells(i, 3).Value
        dateFin = wsSource.Cells(i, 4).Value

        ' Ligne de base
        wsDest.Cells(ligneEcriture, 1).Value = responsable
        wsDest.Cells(ligneEcriture, 2).Value = projet

        ' Marquer les semaines concernées
        For semaine = semaineCourante To semaineMax
            Dim dateSemaine As Date
            dateSemaine = DateAdd("ww", semaine - semaineCourante, currentDate)

            If dateFin >= dateSemaine And dateDebut <= dateSemaine Then
                wsDest.Cells(ligneEcriture, semaine - semaineCourante + 3).Value = "✓"
            End If
        Next semaine

        ligneEcriture = ligneEcriture + 1
    Next i

    ' Formatage
    wsDest.Columns.AutoFit
    MsgBox "Planning généré avec succès !"

End Sub
