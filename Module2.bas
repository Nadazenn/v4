Attribute VB_Name = "Module2"
Sub CreerHistogrammeNombrePalettes()
    Dim ws As Worksheet, wsBilan As Worksheet, wsLivrable As Worksheet
    Dim chartObj As ChartObject
    Dim Chart As Chart
    Dim i As Long
    Dim lastRow As Long, lastRow2 As Long
    
    ' Définir la feuille
    Set ws = ThisWorkbook.Sheets("Bilan Graphique") ' Remplacez "Sheet1" par le nom de votre feuille
    Set wsBilan = ThisWorkbook.Sheets("Bilan")
    Set wsLivrable = ThisWorkbook.Sheets("Livrable")

    ' Dernière ligne de données
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastRow2 = wsBilan.Cells(ws.Rows.Count, 3).End(xlUp).Row

    ' Créer un nouvel objet graphique
    Set chartObj = wsBilan.ChartObjects.Add(Left:=50, Width:=500, Top:=wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Définir le type de graphique (histogramme empilé)
    Chart.ChartType = xlColumnStacked

    ' Ajouter les séries de données
    Chart.SetSourceData Source:=ws.Range("C2:D" & lastRow)
    
    ' Configurer l'axe des catégories pour correspondre aux étages
    Chart.Axes(xlCategory).CategoryNames = ws.Range("B2:B" & lastRow)

    ' Configurer la légende (pour avoir "Production" et "Terminaux")
    Chart.HasLegend = True
    Chart.SeriesCollection(1).Name = "Production"
    Chart.SeriesCollection(2).Name = "Terminaux"
    
    ' Ajouter un titre au graphique
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Palettes équivalentes par étage"
    
    ' Ajouter un titre à l'axe des abscisses (axe des catégories)
    Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Étage et Zone" ' Titre de l'axe des abscisses

    ' Ajouter un titre à l'axe des ordonnées (axe des valeurs)
    Chart.Axes(xlValue, xlPrimary).HasTitle = True
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Nombre de Palettes" ' Titre de l'axe des ordonnées

    ' Copiez le graphique
    chartObj.Copy
    
    ' Collez-le temporairement sur la feuille Livrable (par exemple, en A1)
    wsLivrable.Activate
    wsLivrable.Range("A1").Select
    wsLivrable.Paste
    
    ' Ajustez la position et la taille du graphique collé
    With wsLivrable.ChartObjects(wsLivrable.ChartObjects.Count)
        .Left = 180
        .Top = 145
        .Width = 299
        .Height = 130.5
        
        With .Chart
            ' Titre principal
            .ChartTitle.Font.Size = 12
        
            ' Titres des axes
            .Axes(xlCategory).AxisTitle.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 7
            .Axes(xlValue).AxisTitle.Font.Size = 7
            .Axes(xlValue).TickLabels.Font.Size = 7
        
            ' Légende
            .Legend.Font.Size = 7
        End With
    End With

End Sub



Sub CreerHistogrammeRemplissageCamions()
    Dim ws As Worksheet, wsBilan As Worksheet
    Dim chartObj As ChartObject
    Dim Chart As Chart
    Dim lastRow As Long

    ' Définir les feuilles
    Set ws = ThisWorkbook.Sheets("Bilan Graphique")
    Set wsBilan = ThisWorkbook.Sheets("Bilan")

    ' Dernière ligne de données
    lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    lastRow2 = wsBilan.Cells(ws.Rows.Count, 3).End(xlUp).Row


    ' Créer un nouvel objet graphique
    Set chartObj = wsBilan.ChartObjects.Add(Left:=700, Width:=500, Top:=wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Définir les plages de données
    Dim xLabelsRange As Range, yValuesRange1 As Range, yValuesRange2 As Range
    Set xLabelsRange = ws.Range("B2:B" & lastRow) ' Colonne des catégories (par exemple "Étage - Zone")
    Set yValuesRange1 = ws.Range("K2:K" & lastRow) ' Remplissage sans CCC
    Set yValuesRange2 = ws.Range("L2:L" & lastRow) ' Remplissage avec CCC

    ' Ajouter une série pour l'histogramme
    Chart.SeriesCollection.NewSeries
    Chart.SeriesCollection(1).Name = "Remplissage camions sans CCC"
    Chart.SeriesCollection(1).Values = yValuesRange1
    Chart.SeriesCollection(1).XValues = xLabelsRange
    Chart.SeriesCollection(1).ChartType = xlColumnClustered ' Histogramme

    ' Ajouter une série pour la courbe
    Chart.SeriesCollection.NewSeries
    Chart.SeriesCollection(2).Name = "Remplissage camions avec CCC"
    Chart.SeriesCollection(2).Values = yValuesRange2
    Chart.SeriesCollection(2).XValues = xLabelsRange
    Chart.SeriesCollection(2).ChartType = xlColumnClustered ' Ligne avec marqueurs

    ' Ajouter un titre au graphique
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Comparaison du remplissage des camions par étage et zone"

    ' Configurer les axes
    Chart.Axes(xlCategory).HasTitle = True
    Chart.Axes(xlCategory).AxisTitle.Text = "Étage et Zone"
    Chart.Axes(xlValue).HasTitle = True
    Chart.Axes(xlValue).AxisTitle.Text = "Remplissage (%)"
End Sub


