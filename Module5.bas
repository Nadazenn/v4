Attribute VB_Name = "Module5"
Sub CreerHistogrammeNombreCamions()
    Dim ws As Worksheet, wsBilan As Worksheet, wsLivrable As Worksheet
    Dim chartObj As ChartObject
    Dim Chart As Chart
    Dim i As Long
    Dim lastRow As Long, lastRow2 As Long
    
    ' Définir la feuille
    Set ws = ThisWorkbook.Sheets("Bilan Graphique")
    Set wsBilan = ThisWorkbook.Sheets("Bilan")
    Set wsLivrable = ThisWorkbook.Sheets("Livrable")

    ' Dernière ligne de données
    lastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
    lastRow2 = wsBilan.Cells(ws.Rows.Count, 3).End(xlUp).Row

    ' Créer un nouvel objet graphique
    Set chartObj = wsBilan.ChartObjects.Add(Left:=700, Width:=500, Top:=350 + wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Définir le type de graphique (histogramme empilé)
    Chart.ChartType = xlColumnStacked

    ' Ajouter les séries de données
    Chart.SetSourceData Source:=ws.Range("G2:H" & lastRow)
    
    ' Configurer l'axe des catégories pour correspondre aux étages
    Chart.Axes(xlCategory).CategoryNames = ws.Range("F2:F" & lastRow)

    ' Configurer la légende (pour avoir "Production" et "Terminaux")
    Chart.HasLegend = True
    Chart.SeriesCollection(1).Name = "Camions Production"
    Chart.SeriesCollection(2).Name = "Camions Terminaux"

    
    ' Ajouter un titre au graphique
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Camions par étage"
    

    ' Ajouter un titre à l'axe des abscisses (axe des catégories)
    Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Étage et Zone" ' Titre de l'axe des abscisses

    ' Ajouter un titre à l'axe des ordonnées (axe des valeurs)
    Chart.Axes(xlValue, xlPrimary).HasTitle = True
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Nombre de camions" ' Titre de l'axe des ordonnées
   
    ' Copiez le graphique
    chartObj.Copy
    
    ' Collez-le temporairement sur la feuille Livrable (par exemple, en A1)
    wsLivrable.Activate
    wsLivrable.Range("A1").Select
    wsLivrable.Paste
    
    ' Ajustez la position et la taille du graphique collé
    With wsLivrable.ChartObjects(wsLivrable.ChartObjects.Count)
        .Left = 1
        .Top = 304.5
        .Width = 300
        .Height = 174
        With .Chart
            ' Titre principal
            .ChartTitle.Font.Size = 12
        
            ' Titres des axes
            .Axes(xlCategory).AxisTitle.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 7
            .Axes(xlValue).AxisTitle.Font.Size = 7
            .Axes(xlValue).TickLabels.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 5
        
            ' Légende
            .Legend.Font.Size = 7
        End With
    End With
    
    ' Créer un nouvel objet graphique
    Set chartObj = wsBilan.ChartObjects.Add(Left:=700, Width:=500, Top:=350 + wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Définir le type de graphique (histogramme empilé)
    Chart.ChartType = xlColumnStacked

    ' Ajouter les séries de données
    Chart.SetSourceData Source:=ws.Range("I3:J" & lastRow)
    
    ' Configurer l'axe des catégories pour correspondre aux étages
    Chart.Axes(xlCategory).CategoryNames = ws.Range("F2:F" & lastRow)

    ' Configurer la légende (pour avoir "Production" et "Terminaux")
    Chart.HasLegend = True
    Chart.SeriesCollection(1).Name = "Camions Production"
    Chart.SeriesCollection(2).Name = "Camions Terminaux"

    
    ' Ajouter un titre au graphique
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Camions par étage"
    

    ' Ajouter un titre à l'axe des abscisses (axe des catégories)
    Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Étage et Zone" ' Titre de l'axe des abscisses

    ' Ajouter un titre à l'axe des ordonnées (axe des valeurs)
    Chart.Axes(xlValue, xlPrimary).HasTitle = True
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Nombre de camions" ' Titre de l'axe des ordonnées
   
    ' Copiez le graphique
    chartObj.Copy
    
    ' Collez-le temporairement sur la feuille Livrable (par exemple, en A1)
    wsLivrable.Activate
    wsLivrable.Range("A1").Select
    wsLivrable.Paste
    
    ' Ajustez la position et la taille du graphique collé
    With wsLivrable.ChartObjects(wsLivrable.ChartObjects.Count)
        .Left = 482
        .Top = 304.5
        .Width = 300
        .Height = 174
        With .Chart
            ' Titre principal
            .ChartTitle.Font.Size = 12
        
            ' Titres des axes
            .Axes(xlCategory).AxisTitle.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 7
            .Axes(xlValue).AxisTitle.Font.Size = 7
            .Axes(xlValue).TickLabels.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 5
            
            
            ' Légende
            .Legend.Font.Size = 7
        End With
    End With
   
   
    ' Créer un nouvel objet graphique
    Set chartObj = wsBilan.ChartObjects.Add(Left:=700, Width:=500, Top:=700 + wsBilan.Cells(lastRow2 + 2, 1).Top, Height:=300)
    Set Chart = chartObj.Chart

    ' Définir le type de graphique (histogramme empilé)
    Chart.ChartType = xlColumnStacked

    ' Ajouter les séries de données
    Chart.SetSourceData Source:=ws.Range("G2:J" & lastRow)
    
    ' Configurer l'axe des catégories pour correspondre aux étages
    Chart.Axes(xlCategory).CategoryNames = ws.Range("F2:F" & lastRow)

    ' Configurer la légende (pour avoir "Production" et "Terminaux")
    Chart.HasLegend = True
    Chart.SeriesCollection(1).Name = "Production"
    Chart.SeriesCollection(2).Name = "Terminaux"
    Chart.SeriesCollection(3).Name = "Production Opti"
    Chart.SeriesCollection(4).Name = "Terminaux Opti"
    
    ' Ajouter un titre au graphique
    Chart.HasTitle = True
    Chart.ChartTitle.Text = "Comparatif Nombre de camions par étage avec ou sans Optimisation"
    

    ' Ajouter un titre à l'axe des abscisses (axe des catégories)
    Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Étage et Zone" ' Titre de l'axe des abscisses

    ' Ajouter un titre à l'axe des ordonnées (axe des valeurs)
    Chart.Axes(xlValue, xlPrimary).HasTitle = True
    Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Nombre de camions" ' Titre de l'axe des ordonnées
    
    ' Copiez le graphique
    chartObj.Copy
    
    ' Collez-le temporairement sur la feuille Livrable (par exemple, en A1)
    wsLivrable.Activate
    wsLivrable.Range("A1").Select
    wsLivrable.Paste
    
    ' Ajustez la position et la taille du graphique collé
    With wsLivrable.ChartObjects(wsLivrable.ChartObjects.Count)
        .Left = 1080
        .Top = 174
        .Width = 359
        .Height = 188.5
        With .Chart
            ' Titre principal
            .ChartTitle.Font.Size = 12
        
            ' Titres des axes
            .Axes(xlCategory).AxisTitle.Font.Size = 7
            .Axes(xlValue).AxisTitle.Font.Size = 7
            .Axes(xlValue).TickLabels.Font.Size = 7
            .Axes(xlCategory).TickLabels.Font.Size = 5
        
            ' Légende
            .Legend.Font.Size = 7
        End With
    End With

End Sub


