Attribute VB_Name = "Module4"
Sub CreerGraphiqueFluxMensuel()
    Dim ws As Worksheet
    Dim wsBilan As Worksheet, wsParametres As Worksheet, wsLivrable As Worksheet
    Dim chartObj As ChartObject
    Dim Chart As Chart
    Dim i As Long, j As Long, k As Integer
    Dim lastRow As Long, lastRow2 As Long
    Dim dict As Object
    Dim key As Variant
    Dim dateDebut As Date
    Dim dureeEtage As Integer
    Dim nombreMois As Integer, repartitionActuelle As Double, repartitionReste As Double
    Dim moisAnnee As String
    Dim volume As Double
    Dim camions As Long, camionsCCC As Long
    Dim tempArray As Variant
    Dim moisArray() As Date
    Dim sorted As Boolean

    ' Définir la feuille source et créer une feuille pour les flux mensuels
    Set wsBilan = ThisWorkbook.Sheets("Bilan") ' Remplacez "Données" par le nom de votre feuille source
    Set ws = ThisWorkbook.Sheets("Bilan Graphique")
    Set wsParametres = ThisWorkbook.Sheets("Paramétrage")
    Set wsLivrable = ThisWorkbook.Sheets("Livrable")

    ' Initialiser un dictionnaire pour l'agrégation
    Set dict = CreateObject("Scripting.Dictionary")

    ' Parcourir les données de livraison pour regrouper par mois/année
    lastRow = wsParametres.Cells(ws.Rows.Count, "I").End(xlUp).Row
    lastRow2 = wsBilan.Cells(ws.Rows.Count, 3).End(xlUp).Row
    
    j = 2 ' Index pour les lignes des camions
    
    For i = 3 To lastRow
        ' Date de début et durée
        dateDebut = wsParametres.Cells(i, "H").Value - wsParametres.Cells(i, "J").Value ' Date de début
        dureeEtage = wsParametres.Cells(i, "K").Value ' Durée en jours
        ' Calcul du nombre de mois (arrondi à l'entier supérieur)
        nombreMois = Application.WorksheetFunction.RoundUp(dureeEtage / 30, 0)
        ' Récupération des valeurs
        volume = ws.Cells(i - 1, "C").Value
        camions = ws.Cells(j, "G").Value
        camionsCCC = ws.Cells(j + 1, "I").Value
        j = j + 2 ' Incrémentation pour les prochaines lignes de camions
    
        ' Répartition des livraisons :
        ' - 50% dès le premier mois
        ' - Le reste réparti uniformément sur les mois suivants
        If nombreMois = 1 Then
            ' Si un seul mois, on livre tout en une fois
            repartitionReste = 1
        Else
            repartitionReste = 0.5 / (nombreMois - 1) ' Répartition du reste sur les mois suivants
        End If
    
        ' Répartition des livraisons sur plusieurs mois
        For k = 0 To nombreMois - 1
            moisAnnee = Format(DateAdd("m", k, dateDebut), "yyyy-mm")
            If nombreMois = 1 Then
                repartitionActuelle = 1 ' Tout livrer en un seul mois
            ElseIf k = 0 Then
                repartitionActuelle = 0.5 ' 50% au premier mois
            Else
                repartitionActuelle = repartitionReste ' Répartition du reste sur les autres mois
            End If
    
            ' Ajouter au dictionnaire ou mettre à jour les valeurs
            If Not dict.exists(moisAnnee) Then
                dict.Add moisAnnee, Array(volume * repartitionActuelle, camions * repartitionActuelle, camionsCCC * repartitionActuelle)
            Else
                tempArray = dict(moisAnnee)
                tempArray(0) = tempArray(0) + (volume * repartitionActuelle)   ' Mettre à jour le volume
                tempArray(1) = tempArray(1) + (camions * repartitionActuelle)  ' Mettre à jour les camions
                tempArray(2) = tempArray(2) + (camionsCCC * repartitionActuelle) ' Mettre à jour les camions CCC
                dict(moisAnnee) = tempArray ' Réassigner le tableau mis à jour au dictionnaire
            End If
        Next k
    Next i
    
    j = 2
    For i = 3 To lastRow
        ' Date de début et durée
        dateDebut = wsParametres.Cells(i, "I").Value - wsParametres.Cells(i, "J").Value ' Date de début
        dureeEtage = wsParametres.Cells(i, "L").Value ' Durée en jours
    
        ' Calcul du nombre de mois (arrondi à l'entier supérieur)
        nombreMois = Application.WorksheetFunction.RoundUp(dureeEtage / 30, 0)
    
        ' Récupération des valeurs
        volume = ws.Cells(i - 1, "D").Value
        camions = ws.Cells(j, "H").Value
        camionsCCC = ws.Cells(j + 1, "J").Value
        j = j + 2 ' Incrémentation pour les prochaines lignes de camions
        
        ' Répartition des livraisons :
        ' - 50% dès le premier mois
        ' - Le reste réparti uniformément sur les mois suivants
        If nombreMois = 1 Then
            ' Si un seul mois, on livre tout en une fois
            repartitionReste = 1
        Else
            repartitionReste = 0.5 / (nombreMois - 1) ' Répartition du reste sur les mois suivants
        End If
    
        ' Répartition des livraisons sur plusieurs mois
        For k = 0 To nombreMois - 1
            moisAnnee = Format(DateAdd("m", k, dateDebut), "yyyy-mm")
            If nombreMois = 1 Then
                repartitionActuelle = 1 ' Tout livrer en un seul mois
            ElseIf k = 0 Then
                repartitionActuelle = 0.5 ' 50% au premier mois
            Else
                repartitionActuelle = repartitionReste ' Répartition du reste sur les autres mois
            End If
            
            ' Ajouter au dictionnaire ou mettre à jour les valeurs
            If Not dict.exists(moisAnnee) Then
                dict.Add moisAnnee, Array(volume * repartitionActuelle, camions * repartitionActuelle, camionsCCC * repartitionActuelle)
            Else
                tempArray = dict(moisAnnee)
                tempArray(0) = tempArray(0) + (volume * repartitionActuelle)   ' Mettre à jour le volume
                tempArray(1) = tempArray(1) + (camions * repartitionActuelle)  ' Mettre à jour les camions
                tempArray(2) = tempArray(2) + (camionsCCC * repartitionActuelle) ' Mettre à jour les camions CCC
                dict(moisAnnee) = tempArray ' Réassigner le tableau mis à jour au dictionnaire
            End If
        Next k
    Next i
    
    
    
    ' Créer un tableau temporaire pour stocker et trier les dates
    ReDim moisArray(dict.Count - 1)
    
    ' Copier les clés (mois/année) sous forme de dates dans le tableau
    i = 0
    For Each key In dict.keys
        moisArray(i) = DateValue(key & "-01") ' Ajouter un jour pour convertir en date
        i = i + 1
    Next key
    
    ' Trier le tableau de dates
    Do
        sorted = True
        For i = LBound(moisArray) To UBound(moisArray) - 1
            If moisArray(i) > moisArray(i + 1) Then
                ' Échanger les valeurs
                Dim temp As Date
                temp = moisArray(i)
                moisArray(i) = moisArray(i + 1)
                moisArray(i + 1) = temp
                sorted = False
            End If
        Next i
    Loop Until sorted
    
    ' Remplir la feuille des flux mensuels avec les données agrégées
    ws.Cells(1, 13).Value = "Mois"
    ws.Cells(1, 14).Value = "Volume (nombre de palettes équivalentes)"
    ws.Cells(1, 15).Value = "Nombre de Camions"
    ws.Cells(1, 16).Value = "Nombre de Camions CCC"
    
    ' Remplir la feuille des flux mensuels avec les données triées par mois
    j = 2
    For i = LBound(moisArray) To UBound(moisArray)
        Dim sortedKey As String
        sortedKey = Format(moisArray(i), "yyyy-mm")
        
        ' Remplir les données agrégées dans la feuille
        ws.Cells(j, 13).Value = sortedKey
        ws.Cells(j, 14).Value = dict(sortedKey)(0) ' Volume total
        ws.Cells(j, 15).Value = dict(sortedKey)(1) ' Nombre de camions total
        ws.Cells(j, 16).Value = dict(sortedKey)(2) ' Nombre de camions total CCC
        j = j + 1
    Next i

    ' Créer un graphique de flux mensuel
    Set chartObj = wsBilan.ChartObjects.Add(Left:=50, Width:=500, Top:=350 + wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Configurer les données du graphique
    Chart.SetSourceData Source:=ws.Range("M1:O" & j - 1)
    Chart.ChartType = xlColumnClustered ' Histogramme groupé pour comparer les volumes et camions par mois
    
    ' Définir le type du graphique : histogramme pour le volume et courbe pour le nombre de camions
    Chart.ChartType = xlColumnClustered ' Histogramme pour le volume (par défaut)
    Chart.SeriesCollection(1).Name = "Nombre de palettes"
    Chart.SeriesCollection(2).Name = "Nombre de Camions"

    ' Configurer la série "Nombre de Camions" comme une courbe
    With Chart.SeriesCollection(2)
    .ChartType = xlLine ' Série "Nombre de Camions" en courbe
    .Smooth = True  ' Appliquer un adoucissement pour obtenir une courbe
    .AxisGroup = 2  ' Assigner à un axe secondaire
    End With
    
    ' Configurer les titres et les axes
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Flux Mensuel de Livraison"
    Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Mois"
    
    ' Ajouter des titres aux axes primaires et secondaires
    Chart.Axes(xlValue, xlPrimary).HasTitle = True
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Nombre de palettes"
    Chart.Axes(xlValue, xlSecondary).HasTitle = True
    Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = "Nombre de Camions"
   
    ' Copiez le graphique
    chartObj.Copy
    
    ' Collez-le temporairement sur la feuille Livrable (par exemple, en A1)
    wsLivrable.Activate
    wsLivrable.Range("A1").Select
    wsLivrable.Paste
    
    ' Ajustez la position et la taille du graphique collé
    With wsLivrable.ChartObjects(wsLivrable.ChartObjects.Count)
        .Left = 1
        .Top = 522
        .Width = 478
        .Height = 188.5
        With .Chart
            ' Titre principal
            .ChartTitle.Font.Size = 12
        
            ' Titres des axes
            .Axes(xlCategory).AxisTitle.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 7
            .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 7
            .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 7
            .Axes(xlValue, xlSecondary).AxisTitle.Font.Size = 7
            .Axes(xlValue, xlSecondary).TickLabels.Font.Size = 7
        
            ' Légende
            .Legend.Font.Size = 7
        End With
    End With
    
    ' Créer un graphique de flux mensuel
    Set chartObj = wsBilan.ChartObjects.Add(Left:=50, Width:=500, Top:=350 + wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Configurer les données du graphique
    Chart.SetSourceData Source:=Union(ws.Range("M1:M" & j - 1), ws.Range("N1:N" & j - 1), ws.Range("P1:P" & j - 1))
    Chart.ChartType = xlColumnClustered ' Histogramme groupé pour comparer les volumes et camions par mois
    
    ' Définir le type du graphique : histogramme pour le volume et courbe pour le nombre de camions
    Chart.ChartType = xlColumnClustered ' Histogramme pour le volume (par défaut)
    Chart.SeriesCollection(1).Name = "Nombre de palettes"
    Chart.SeriesCollection(2).Name = "Nombre de Camions"

    ' Configurer la série "Nombre de Camions" comme une courbe
    With Chart.SeriesCollection(2)
    .ChartType = xlLine ' Série "Nombre de Camions" en courbe
    .Smooth = True  ' Appliquer un adoucissement pour obtenir une courbe
    .AxisGroup = 2  ' Assigner à un axe secondaire
    End With
    
    ' Configurer les titres et les axes
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Flux Mensuel de Livraison"
    Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Mois"
    
    ' Ajouter des titres aux axes primaires et secondaires
    Chart.Axes(xlValue, xlPrimary).HasTitle = True
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Nombre de palettes"
    Chart.Axes(xlValue, xlSecondary).HasTitle = True
    Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = "Nombre de Camions"
   
    ' Copiez le graphique
    chartObj.Copy
    
    ' Collez-le temporairement sur la feuille Livrable (par exemple, en A1)
    wsLivrable.Activate
    wsLivrable.Range("A1").Select
    wsLivrable.Paste
    
    ' Ajustez la position et la taille du graphique collé
    With wsLivrable.ChartObjects(wsLivrable.ChartObjects.Count)
        .Left = 482
        .Top = 522
        .Width = 477
        .Height = 188.5
        With .Chart
            ' Titre principal
            .ChartTitle.Font.Size = 12
        
            ' Titres des axes
            .Axes(xlCategory).AxisTitle.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 7
            .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 7
            .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 7
            .Axes(xlValue, xlSecondary).AxisTitle.Font.Size = 7
            .Axes(xlValue, xlSecondary).TickLabels.Font.Size = 7
        
            ' Légende
            .Legend.Font.Size = 7
        End With
    End With
    
   
    ' Créer un graphique de flux mensuel AVEC r2SO
    Set chartObj = wsBilan.ChartObjects.Add(Left:=50, Width:=500, Top:=700 + wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Configurer les données du graphique
    Chart.SetSourceData Source:=ws.Range("M1:P" & j - 1)
    Chart.ChartType = xlColumnClustered ' Histogramme groupé pour comparer les volumes et camions par mois
    
    ' Définir le type du graphique : histogramme pour le volume et courbe pour le nombre de camions
    Chart.SeriesCollection(1).Name = "Nombre de palettes"
    Chart.SeriesCollection(2).Name = "Non Optimisée"
    Chart.SeriesCollection(3).Name = "Optimisée"

    ' Configurer la série "Nombre de Camions sans CCC" comme une courbe
    With Chart.SeriesCollection(1)
    .XValues = ws.Range("M2:M" & j - 1) ' Mois
    .Values = ws.Range("N2:N" & j - 1) ' Nombre de palettes
    .ChartType = xlColumnClustered  ' Série "Nombre de palettes" en histogramme
    .AxisGroup = 1
    End With
    
    ' Configurer la série "Nombre de Camions sans CCC" comme une courbe
    With Chart.SeriesCollection(2)
    .XValues = ws.Range("M2:M" & j - 1) ' Mois
    .Values = ws.Range("O2:O" & j - 1)
    .ChartType = xlLine ' Série "Nombre de Camions" en courbe
    .Smooth = True  ' Appliquer un adoucissement pour obtenir une courbe
    .AxisGroup = 2  ' Assigner à un axe secondaire
    End With
    
    ' Configurer la série "Nombre de Camions avec CCC" comme une courbe
    With Chart.SeriesCollection(3)
    .XValues = ws.Range("M2:M" & j - 1) ' Mois
    .Values = ws.Range("P2:P" & j - 1)
    .ChartType = xlLine ' Série "Nombre de Camions" en courbe
    .Smooth = True  ' Appliquer un adoucissement pour obtenir une courbe
    .AxisGroup = 2  ' Assigner à un axe secondaire
    End With
    
    ' Configurer les titres et les axes
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Comparatif Flux Mensuel de Livraison avec ou sans Optimisation"
    Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Mois"
    
    ' Ajouter des titres aux axes primaires et secondaires
    Chart.Axes(xlValue, xlPrimary).HasTitle = True
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Nombre de palettes"
    Chart.Axes(xlValue, xlSecondary).HasTitle = True
    Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = "Nombre de Camions"


    ' Copiez le graphique
    chartObj.Copy
    
    ' Collez-le temporairement sur la feuille Livrable (par exemple, en A1)
    wsLivrable.Activate
    wsLivrable.Range("A1").Select
    wsLivrable.Paste
    
    ' Ajustez la position et la taille du graphique collé
    With wsLivrable.ChartObjects(wsLivrable.ChartObjects.Count)
        .Left = 1080
        .Top = 377
        .Width = 359
        .Height = 188.5
        With .Chart
            ' Titre principal
            .ChartTitle.Font.Size = 12
        
            ' Titres des axes
            .Axes(xlCategory).AxisTitle.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 7
            .Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 7
            .Axes(xlValue, xlPrimary).TickLabels.Font.Size = 7
            .Axes(xlValue, xlSecondary).AxisTitle.Font.Size = 7
            .Axes(xlValue, xlSecondary).TickLabels.Font.Size = 7

            ' Légende
            .Legend.Font.Size = 7
        End With
    End With

End Sub






