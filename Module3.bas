Attribute VB_Name = "Module3"
Sub CreerBilanZones()
    Dim wsSource As Worksheet
    Dim wsBilan As Worksheet
    Dim wsParametres As Worksheet
    Dim wsBilanGraphique As Worksheet, wsLivrable As Worksheet
    Dim lastRowDonnees As Long
    Dim currentZone As String, currentEtage As String
    Dim totalPalettes As Long, totalPalettesProd As Long, totalPalettesTerm As Long
    Dim i As Long, nextRow As Long, nextRow2 As Long
    Dim dernier As Long
    Dim sommeProd As Double, somme As Double
    Dim indexMatch As Variant, ligneStockCCC As Variant, ligneStockCCCProd As Variant, ligneStockCCCTerm As Variant
    Dim listeEtages As Range
    Dim chartObj As ChartObject
    Dim Chart As Chart
    Dim dictListeCamion As Object, dictListeCamionCCC As Object, dictListeMaterielCCC As Object
    Dim camionType As String, nombreCamion As Double, materiel As String, nombremateriel As Double
    Dim key As Variant
    Dim keys() As String
    Dim denominateur As Double
    
    
    ' Définir la feuille source contenant les données
    Set wsSource = ThisWorkbook.Sheets("Tableau Source") ' Remplacer par le nom de ta feuille
    lastRowDonnees = wsSource.Cells(wsSource.Rows.Count, "C").End(xlUp).Row ' Dernière ligne des données
    
    ' Créer une nouvelle feuille pour le résumé
    On Error Resume Next ' Eviter les erreurs si la feuille existe déjà
    Set wsBilan = ThisWorkbook.Sheets("Bilan")
    On Error GoTo 0
    
    If wsBilan Is Nothing Then
        Set wsBilan = ThisWorkbook.Sheets.Add
        wsBilan.Name = "Sortie"
    End If
    
    Set wsParametres = ThisWorkbook.Sheets("Paramétrage")
       
    dernier = wsParametres.Cells(wsSource.Rows.Count, "G").End(xlUp).Row
    Set listeZones = wsParametres.Range(wsParametres.Cells(3, 7), wsParametres.Cells(dernier, 7))
    
    
    'Configuration feuille Livrable
    
    Set wsLivrable = ThisWorkbook.Sheets("Livrable")
    For Each chartObj In wsLivrable.ChartObjects
        chartObj.Delete
    Next chartObj
    
    'Configuration feuille graphiques
    
    Set wsBilanGraphique = ThisWorkbook.Sheets("Bilan Graphique")
    wsBilanGraphique.Range("A:AB").ClearContents
    
    wsBilanGraphique.Cells(1, 2).Value = "Étage - Zone"
    wsBilanGraphique.Cells(1, 3).Value = "Production"
    wsBilanGraphique.Cells(1, 4).Value = "Terminaux"
    wsBilanGraphique.Cells(1, 6).Value = "Étage - Zone"
    wsBilanGraphique.Cells(1, 7).Value = "Camions Production sans CCC"
    wsBilanGraphique.Cells(1, 8).Value = "Camions Terminaux sans CCC"
    wsBilanGraphique.Cells(1, 9).Value = "Camions Production avec CCC"
    wsBilanGraphique.Cells(1, 10).Value = "Camions Terminaux avec CCC"
    wsBilanGraphique.Cells(1, 11).Value = "Remplissage camions sans CCC"
    wsBilanGraphique.Cells(1, 12).Value = "Remplissage camions avec CCC"
    
    
    
    ' Effacer tout le contenu et les formats
    wsBilan.Cells.ClearContents
    wsBilan.Cells.ClearFormats
    wsBilan.Cells.Validation.Delete
    
    For Each chartObj In wsBilan.ChartObjects
        chartObj.Delete
    Next chartObj
    
    ' Résumé volume
    ' Ajouter les en-têtes dans la feuille Résumé
    wsBilan.Cells(1, 1).Value = "Étage"
    wsBilan.Cells(1, 2).Value = "Zone"
    wsBilan.Cells(1, 3).Value = "Phase"
    wsBilan.Cells(1, 4).Value = "Total Palettes Équivalent"

    wsBilan.Cells.Columns(1).Font.Bold = True
    wsBilan.Cells.Columns(1).Font.Color = RGB(255, 255, 255)
    wsBilan.Cells.Columns(2).Font.Color = RGB(255, 255, 255)
    wsBilan.Cells.Rows(1).Font.Bold = True
    wsBilan.Cells.Rows(1).Font.Color = RGB(255, 255, 255)
    wsBilan.Cells.Range(wsBilan.Cells(1, 1), wsBilan.Cells(1, 4)).Interior.Color = RGB(0, 32, 96)
    
    
    nextRow = 2 ' La ligne suivante après les en-têtes
    
    ' Initialiser les variables pour stocker les étages et zones actuels
    currentEtage = ""
    currentZone = ""
    
    totalPalettes = 0
    totalPalettesProd = 0
    totalPalettesTerm = 0
    
    ' Initialiser les dictionnaires
    Set dictListeCamion = CreateObject("Scripting.Dictionary")
    Set dictListeCamionCCC = CreateObject("Scripting.Dictionary")
    Set dictListeMaterielCCC = CreateObject("Scripting.Dictionary")

    ' Boucler sur les lignes de la feuille source
    For i = 2 To lastRowDonnees

        currentEtage = wsSource.Cells(i, 1).Value
        currentZone = wsSource.Cells(i, 2).Value
        
        ' Ajouter les palettes de la ligne courante au total pour la zone
        If wsSource.Cells(i, 16).Value <> "" Then
            totalPalettes = totalPalettes + wsSource.Cells(i, 11).Value
        End If
        
        If wsSource.Cells(i, 4).Value = "Production" Then
            totalPalettesProd = totalPalettesProd + wsSource.Cells(i, 11).Value
        End If
        
        If wsSource.Cells(i, 4).Value = "Terminaux" Then
            totalPalettesTerm = totalPalettesTerm + wsSource.Cells(i, 11).Value
        End If
        
        
        ' Vérifier si on passe à une nouvelle zone ou si c'est la dernière ligne
        If wsSource.Cells(i + 1, 1).Value <> currentEtage Or wsSource.Cells(i + 1, 2).Value <> currentZone Or i = lastRowDonnees Then
            ' Remplir les données dans la feuille Résumé UNE SEULE FOIS PAR ZONE
            wsBilan.Cells(nextRow, 1).Value = currentEtage ' Étage
            wsBilan.Cells(nextRow, 1).Interior.Color = RGB(0, 32, 96)
            wsBilan.Cells(nextRow, 2).Value = currentZone ' Zone
            wsBilan.Cells(nextRow, 2).Interior.Color = RGB(0, 32, 96)
            wsBilan.Cells(nextRow, 3).Value = "Production" ' Total palettesProd
            wsBilan.Cells(nextRow + 1, 3).Value = "Terminaux" ' Total palettesTerm
            wsBilan.Cells(nextRow + 2, 3).Value = "Total" ' Total palettes
            wsBilan.Cells(nextRow + 2, 3).Font.Bold = True
            wsBilan.Cells(nextRow, 4).Value = totalPalettesProd ' Total palettesProd
            wsBilan.Cells(nextRow + 1, 4).Value = totalPalettesTerm ' Total palettesTerm
            wsBilan.Cells(nextRow + 2, 4).Value = totalPalettes ' Total palettes
            wsBilan.Cells(nextRow + 2, 4).Font.Bold = True
            wsBilan.Cells.Range(wsBilan.Cells(nextRow + 1, 3), wsBilan.Cells(nextRow + 1, 4)).Interior.Color = wsSource.Cells(i, 2).Interior.Color
            
            
            nextRow = nextRow + 3
            
            'Stocker dans la feuille graphique
            wsBilanGraphique.Cells(nextRow / 3, 2).Value = currentEtage & " - " & currentZone
            wsBilanGraphique.Cells(nextRow / 3, 3).Value = totalPalettesProd
            wsBilanGraphique.Cells(nextRow / 3, 4).Value = totalPalettesTerm
            
            ' Réinitialiser pour la zone suivante
            totalPalettes = 0
            totalPalettesProd = 0
            totalPalettesTerm = 0
            
        End If
        If wsSource.Cells(i, 16).Value <> "" Then
            camionType = wsSource.Cells(i, 12).Value
            nombreCamion = wsSource.Cells(i, 13).Value
            key = currentEtage & currentZone & "|" & camionType
            If Not dictListeCamion.exists(key) Then
                dictListeCamion.Add key, nombreCamion
            Else
                dictListeCamion(key) = dictListeCamion(key) + nombreCamion
            End If
        End If
        
        If wsSource.Cells(i, 16).Value <> "Oui" Then
            camionType = wsSource.Cells(i, 12).Value
            nombreCamion = wsSource.Cells(i, 13).Value
            key = currentEtage & currentZone & "|" & camionType
            If Not dictListeCamionCCC.exists(key) Then
                dictListeCamionCCC.Add key, nombreCamion
            Else
                dictListeCamionCCC(key) = dictListeCamionCCC(key) + nombreCamion
            End If
        End If
        
        If wsSource.Cells(i, 16).Value = "Oui" Then
            materiel = wsSource.Cells(i, 5).Value
            nombremateriel = wsSource.Cells(i, 7).Value
            key = materiel
            If Not dictListeMaterielCCC.exists(key) Then
                dictListeMaterielCCC.Add key, nombreCamion
            Else
                dictListeMaterielCCC(key) = dictListeMaterielCCC(key) + nombremateriel
            End If
        End If
    Next i
    
    
    ' Insérer les résultats dans la feuille
    wsBilanGraphique.Cells(1, 18).Value = "Étage"
    wsBilanGraphique.Cells(1, 19).Value = "Zone"
    wsBilanGraphique.Cells(1, 20).Value = "Type de Camion"
    wsBilanGraphique.Cells(1, 21).Value = "Nombre de Camions"

    i = 2
    For Each key In dictListeCamion.keys
        keys = Split(key, "|")
        wsBilanGraphique.Cells(i, 19).Value = keys(0) ' Étages et zones
        wsBilanGraphique.Cells(i, 20).Value = keys(1) ' Type de camion
        wsBilanGraphique.Cells(i, 21).Value = dictListeCamion(key) ' Nombre de camions
        i = i + 1
    Next key
    
    ' Insérer les résultats dans la feuille
    wsBilanGraphique.Cells(1, 23).Value = "Étage"
    wsBilanGraphique.Cells(1, 24).Value = "Type de Camion"
    wsBilanGraphique.Cells(1, 25).Value = "Nombre de Camions avec CCC"
    
    i = 2
    For Each key In dictListeCamionCCC.keys
        keys = Split(key, "|")
        wsBilanGraphique.Cells(i, 23).Value = keys(0) ' Étages
        wsBilanGraphique.Cells(i, 24).Value = keys(1) ' Type de camion
        wsBilanGraphique.Cells(i, 25).Value = dictListeCamionCCC(key) ' Nombre de camions
        i = i + 1
    Next key
    
    wsBilanGraphique.Cells(1, 27).Value = "Matériel CCC"
    wsBilanGraphique.Cells(1, 28).Value = "Nombre de matériels CCC"
    i = 2
    For Each key In dictListeMaterielCCC.keys
        keys = Split(key, "|")
        wsBilanGraphique.Cells(i, 27).Value = keys(0) ' matériel
        wsBilanGraphique.Cells(i, 28).Value = dictListeMaterielCCC(key) ' Nombre de matériels
        i = i + 1
    Next key
    
    
    
    wsBilan.Cells(nextRow, 2).Value = "Total"
    wsBilan.Cells(nextRow, 2).Font.Bold = True
    wsBilan.Cells(nextRow, 3).Value = "Production" ' Total palettesProd
    wsBilan.Cells(nextRow + 1, 3).Value = "Terminaux" ' Total palettesTerm
    wsBilan.Cells(nextRow + 2, 3).Value = "Total" ' Total palettes
    wsBilan.Cells(nextRow, 4).Value = WorksheetFunction.SumIf(wsBilan.Range(wsBilan.Cells(2, 3), wsBilan.Cells(nextRow - 1, 3)), "Production", _
            wsBilan.Range(wsBilan.Cells(2, 4), wsBilan.Cells(nextRow - 1, 4)))
    wsBilan.Cells(nextRow + 1, 4).Value = WorksheetFunction.SumIf(wsBilan.Range(wsBilan.Cells(2, 3), wsBilan.Cells(nextRow - 1, 3)), "Terminaux", _
            wsBilan.Range(wsBilan.Cells(2, 4), wsBilan.Cells(nextRow - 1, 4)))
    wsBilan.Cells(nextRow + 2, 4).Value = wsBilan.Cells(nextRow, 4).Value + wsBilan.Cells(nextRow + 1, 4).Value


    wsBilan.Cells.Range(wsBilan.Cells(nextRow, 1), wsBilan.Cells(nextRow + 2, 4)).Interior.Color = RGB(0, 112, 192)
    wsBilan.Cells.Range(wsBilan.Cells(nextRow + 2, 3), wsBilan.Cells(nextRow + 2, 4)).Font.Bold = True
    wsBilan.Cells.Range(wsBilan.Cells(nextRow, 2), wsBilan.Cells(nextRow + 2, 4)).Font.Color = RGB(255, 255, 255)
    
    ' Résumé Camions

    ' Ajouter les en-têtes dans la feuille Résumé
    wsBilan.Cells(1, 6).Value = "Étage"
    wsBilan.Cells(1, 7).Value = "Zone"
    wsBilan.Cells(1, 8).Value = "Phase"
    wsBilan.Cells(1, 9).Value = "Nombre total de camions sans CCC"
    wsBilan.Cells(1, 10).Value = "Nombre total de camions avec CCC"
    wsBilan.Cells(1, 11).Value = "Remplissage moyen sans CCC"
    wsBilan.Cells(1, 12).Value = "Remplissage moyen avec CCC"
    
    wsBilan.Cells.Columns(6).Font.Bold = True
    wsBilan.Cells.Columns(6).Font.Color = RGB(255, 255, 255)
    wsBilan.Cells.Columns(7).Font.Color = RGB(255, 255, 255)
    wsBilan.Cells.Range(wsBilan.Cells(1, 6), wsBilan.Cells(1, 12)).Interior.Color = RGB(0, 32, 96)
    
    
    nextRow = 2 ' La ligne suivante après les en-têtes
    nextRow2 = 2
    
    
    
    ' Boucler sur les lignes de la feuille source
    For Each zone In listeZones
        etage = wsBilan.Cells(nextRow, 1).Value
        wsBilan.Cells(nextRow, 6).Value = etage ' Étage
        wsBilan.Cells(nextRow, 6).Interior.Color = RGB(0, 32, 96)
        
        wsBilan.Cells(nextRow, 7).Value = zone ' Étage
        wsBilan.Cells(nextRow, 7).Interior.Color = RGB(0, 32, 96)
        
        wsBilan.Cells(nextRow, 8).Value = "Production" ' Total palettesProd
        wsBilan.Cells(nextRow + 1, 8).Value = "Terminaux"
        wsBilan.Cells(nextRow + 2, 8).Value = "Total"
        wsBilan.Cells(nextRow + 2, 8).Font.Bold = True
        
        wsBilan.Cells(nextRow, 9).Value = WorksheetFunction.SumIfs( _
                                                            wsSource.Range("M:M"), _
                                                            wsSource.Range("A:A"), etage, _
                                                            wsSource.Range("B:B"), zone, _
                                                            wsSource.Range("P:P"), "<>", _
                                                            wsSource.Range("D:D"), "Production")
        wsBilan.Cells(nextRow + 1, 9).Value = WorksheetFunction.SumIfs( _
                                                            wsSource.Range("M:M"), _
                                                            wsSource.Range("A:A"), etage, _
                                                            wsSource.Range("B:B"), zone, _
                                                            wsSource.Range("P:P"), "<>", _
                                                            wsSource.Range("D:D"), "Terminaux")
        wsBilan.Cells(nextRow + 2, 9).Value = wsBilan.Cells(nextRow, 9).Value + wsBilan.Cells(nextRow + 1, 9).Value
        wsBilan.Cells(nextRow + 2, 9).Font.Bold = True
        
        ligneStockCCCProd = wsSource.Evaluate("MATCH(1, INDEX((A:A=" & etage & ")*(B:B=""" & zone & """)*(P:P=""""), 0), 0)")
        ligneStockCCCTerm = ligneStockCCCProd + 1
        
        
        wsBilan.Cells(nextRow, 10).Value = WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Production") + _
                                                            stockCCCProdM
        wsBilan.Cells(nextRow + 1, 10).Value = WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Terminaux") + _
                                                            stockCCCTermM
        wsBilan.Cells(nextRow + 2, 10).Value = wsBilan.Cells(nextRow, 10).Value + wsBilan.Cells(nextRow + 1, 10).Value
        wsBilan.Cells(nextRow + 2, 10).Font.Bold = True
        
        wsBilan.Cells(nextRow, 11).Value = WorksheetFunction.RoundUp(( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("O:O"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>", wsSource.Range("D:D"), "Production") + _
                                                            WorksheetFunction.SumIfs(wsSource.Range("N:N"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>", wsSource.Range("D:D"), "Production")) / _
                                                            WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>", wsSource.Range("D:D"), "Production"), 2)
        
        denominateur = WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>", wsSource.Range("D:D"), "Terminaux")

        If denominateur = 0 Then
            wsBilan.Cells(nextRow + 1, 11).Value = 0
        Else
            wsBilan.Cells(nextRow + 1, 11).Value = WorksheetFunction.RoundUp( _
                (WorksheetFunction.SumIfs(wsSource.Range("O:O"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>", wsSource.Range("D:D"), "Terminaux") + _
                 WorksheetFunction.SumIfs(wsSource.Range("N:N"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>", wsSource.Range("D:D"), "Terminaux")) / denominateur, 2)
        End If
        wsBilan.Cells(nextRow + 2, 11).Value = WorksheetFunction.RoundUp(( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("O:O"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>") + _
                                                            WorksheetFunction.SumIfs(wsSource.Range("N:N"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>")) / _
                                                            WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "<>"), 2)
        wsBilan.Cells(nextRow + 2, 11).Font.Bold = True
        
        
        wsBilan.Cells(nextRow, 12).Value = WorksheetFunction.RoundUp(( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("O:O"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Production") + _
                                                            WorksheetFunction.SumIfs(wsSource.Range("N:N"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Production") + _
                                                            stockCCCProdO + _
                                                            stockCCCProdN) / ( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Production") + _
                                                            stockCCCProdM _
                                                           ), 2)
                                                           
        ' Calcul du dénominateur
        denominator = WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Terminaux") + _
              stockCCCTermM
        If denominateur = 0 Then
            wsBilan.Cells(nextRow + 1, 12).Value = 0
        Else
            wsBilan.Cells(nextRow + 1, 12).Value = WorksheetFunction.RoundUp(( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("O:O"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Terminaux") + _
                                                            WorksheetFunction.SumIfs(wsSource.Range("N:N"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Terminaux") + _
                                                            stockCCCTermO + _
                                                            stockCCCTermN) / ( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non", wsSource.Range("D:D"), "Terminaux") + _
                                                            stockCCCTermM _
                                                            ), 2)
        End If
        
        
        wsBilan.Cells(nextRow + 2, 12).Value = WorksheetFunction.RoundUp(( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("O:O"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non") + _
                                                            WorksheetFunction.SumIfs(wsSource.Range("N:N"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non") + _
                                                            stockCCCProdO + _
                                                            stockCCCTermO + _
                                                            stockCCCProdN + _
                                                            stockCCCTermN) / ( _
                                                            WorksheetFunction.SumIfs(wsSource.Range("M:M"), wsSource.Range("A:A"), etage, wsSource.Range("B:B"), zone, wsSource.Range("P:P"), "Non") + _
                                                            stockCCCProdM + _
                                                            stockCCCTermM _
                                                            ), 2)
        wsBilan.Cells(nextRow + 2, 12).Font.Bold = True
    
        
        wsBilan.Cells.Range(wsBilan.Cells(nextRow + 1, 8), wsBilan.Cells(nextRow + 1, 12)).Interior.Color = wsSource.Cells(i, 2).Interior.Color
        
        'Stocker dans la feuille graphique
        wsBilanGraphique.Cells(nextRow2, 6).Value = etage & " - " & zone
        wsBilanGraphique.Cells(nextRow2, 7).Value = wsBilan.Cells(nextRow, 9).Value
        wsBilanGraphique.Cells(nextRow2, 8).Value = wsBilan.Cells(nextRow + 1, 9).Value
        wsBilanGraphique.Cells(nextRow2 + 1, 9).Value = wsBilan.Cells(nextRow, 10).Value
        wsBilanGraphique.Cells(nextRow2 + 1, 10).Value = wsBilan.Cells(nextRow + 1, 10).Value
        wsBilanGraphique.Cells((nextRow + 3) / 3, 11).Value = wsBilan.Cells(nextRow + 2, 11).Value * 100
        wsBilanGraphique.Cells((nextRow + 3) / 3, 12).Value = wsBilan.Cells(nextRow + 2, 12).Value * 100
        
        
        nextRow = nextRow + 3
        nextRow2 = nextRow2 + 2

        

    Next zone
    
    
    
    wsBilan.Cells(nextRow, 7).Value = "Total"
    wsBilan.Cells(nextRow, 8).Value = "Production"
    wsBilan.Cells(nextRow + 1, 8).Value = "Terminaux"
    wsBilan.Cells(nextRow + 2, 8).Value = "Total"
    
    'Totaux Nombre total de camions sans CCC
    
    wsBilan.Cells(nextRow, 9).Value = WorksheetFunction.SumIf(wsBilan.Range(wsBilan.Cells(2, 8), wsBilan.Cells(nextRow - 1, 8)), "Production", _
            wsBilan.Range(wsBilan.Cells(2, 9), wsBilan.Cells(nextRow - 1, 9)))
    wsBilan.Cells(nextRow + 1, 9).Value = WorksheetFunction.SumIf(wsBilan.Range(wsBilan.Cells(2, 8), wsBilan.Cells(nextRow - 1, 8)), "Terminaux", _
            wsBilan.Range(wsBilan.Cells(2, 9), wsBilan.Cells(nextRow - 1, 9)))
    wsBilan.Cells(nextRow + 2, 9).Value = wsBilan.Cells(nextRow, 9).Value + wsBilan.Cells(nextRow + 1, 9).Value
    
    'Totaux Nombre total de camions avec CCC

    wsBilan.Cells(nextRow, 10).Value = WorksheetFunction.SumIf(wsBilan.Range(wsBilan.Cells(2, 8), wsBilan.Cells(nextRow - 1, 8)), "Production", _
            wsBilan.Range(wsBilan.Cells(2, 10), wsBilan.Cells(nextRow - 1, 10)))
    wsBilan.Cells(nextRow + 1, 10).Value = WorksheetFunction.SumIf(wsBilan.Range(wsBilan.Cells(2, 8), wsBilan.Cells(nextRow - 1, 8)), "Terminaux", _
            wsBilan.Range(wsBilan.Cells(2, 10), wsBilan.Cells(nextRow - 1, 10)))
    wsBilan.Cells(nextRow + 2, 10).Value = wsBilan.Cells(nextRow, 10).Value + wsBilan.Cells(nextRow + 1, 10).Value
    
    'Totaux Remplissage moyen sans CCC

    
    ' Initialise les sommes
    sommeProd = 0
    somme = 0
    
    ' Boucle sur chaque ligne de la plage pour appliquer la condition
    For i = 2 To nextRow - 1
        If wsBilan.Cells(i, 8).Value = "Production" Then ' Vérifie si la colonne H (colonne 8) contient "Total"
            sommeProd = sommeProd + (wsBilan.Cells(i, 9).Value * wsBilan.Cells(i, 11).Value)
            somme = somme + wsBilan.Cells(i, 9).Value
        End If
    Next i
    wsBilan.Cells(nextRow, 11).Value = WorksheetFunction.RoundUp(sommeProd / somme, 2)
    
    ' Initialise les sommes
    sommeProd = 0
    somme = 0
    
    ' Boucle sur chaque ligne de la plage pour appliquer la condition
    For i = 2 To nextRow - 1
        If wsBilan.Cells(i, 8).Value = "Terminaux" Then ' Vérifie si la colonne H (colonne 8) contient "Total"
            sommeProd = sommeProd + (wsBilan.Cells(i, 9).Value * wsBilan.Cells(i, 11).Value)
            somme = somme + wsBilan.Cells(i, 9).Value
        End If
    Next i
    
    denominateur = somme
    If denominateur = 0 Then
        wsBilan.Cells(nextRow + 1, 11).Value = 0
    Else
        wsBilan.Cells(nextRow + 1, 11).Value = WorksheetFunction.RoundUp(sommeProd / somme, 2)
    End If
     ' Initialise les sommes
    sommeProd = 0
    somme = 0
    
    ' Boucle sur chaque ligne de la plage pour appliquer la condition
    For i = 2 To nextRow - 1
        If wsBilan.Cells(i, 8).Value = "Total" Then ' Vérifie si la colonne H (colonne 8) contient "Total"
            sommeProd = sommeProd + (wsBilan.Cells(i, 9).Value * wsBilan.Cells(i, 11).Value)
            somme = somme + wsBilan.Cells(i, 9).Value
        End If
    Next i
    wsBilan.Cells(nextRow + 2, 11).Value = WorksheetFunction.RoundUp(sommeProd / somme, 2)
    
    
    
    
    'Totaux Remplissage moyen avec CCC


    ' Initialise les sommes
    sommeProd = 0
    somme = 0
    
    ' Boucle sur chaque ligne de la plage pour appliquer la condition
    For i = 2 To nextRow - 1
        If wsBilan.Cells(i, 8).Value = "Production" Then ' Vérifie si la colonne H (colonne 8) contient "Total"
            sommeProd = sommeProd + (wsBilan.Cells(i, 9).Value * wsBilan.Cells(i, 12).Value)
            somme = somme + wsBilan.Cells(i, 9).Value
        End If
    Next i
    wsBilan.Cells(nextRow, 12).Value = WorksheetFunction.RoundUp(sommeProd / somme, 2)
    
    ' Initialise les sommes
    sommeProd = 0
    somme = 0
    
    ' Boucle sur chaque ligne de la plage pour appliquer la condition
    For i = 2 To nextRow - 1
        If wsBilan.Cells(i, 8).Value = "Terminaux" Then ' Vérifie si la colonne H (colonne 8) contient "Total"
            sommeProd = sommeProd + (wsBilan.Cells(i, 9).Value * wsBilan.Cells(i, 12).Value)
            somme = somme + wsBilan.Cells(i, 9).Value
        End If
    Next i
    denominateur = somme
    If denominateur = 0 Then
        wsBilan.Cells(nextRow + 1, 12).Value = 0
    Else
        wsBilan.Cells(nextRow + 1, 12).Value = WorksheetFunction.RoundUp(sommeProd / somme, 2)
    End If
    
    
    ' Initialise les sommes
    sommeProd = 0
    somme = 0
    
    ' Boucle sur chaque ligne de la plage pour appliquer la condition
    For i = 2 To nextRow - 1
        If wsBilan.Cells(i, 8).Value = "Total" Then ' Vérifie si la colonne H (colonne 8) contient "Total"
            sommeProd = sommeProd + (wsBilan.Cells(i, 9).Value * wsBilan.Cells(i, 12).Value)
            somme = somme + wsBilan.Cells(i, 9).Value
        End If
    Next i
    wsBilan.Cells(nextRow + 2, 12).Value = WorksheetFunction.RoundUp(sommeProd / somme, 2)




    wsBilan.Cells.Range(wsBilan.Cells(nextRow + 2, 6), wsBilan.Cells(nextRow, 12)).Interior.Color = RGB(0, 112, 192)
    wsBilan.Cells.Range(wsBilan.Cells(nextRow + 2, 7), wsBilan.Cells(nextRow + 2, 12)).Font.Bold = True
    wsBilan.Cells.Range(wsBilan.Cells(nextRow + 2, 7), wsBilan.Cells(nextRow, 12)).Font.Color = RGB(255, 255, 255)
    
    wsBilan.Range("K2:L" & (nextRow + 2)).NumberFormat = "0%"
    
    ' === Copier Materiel Global depuis Données vers Bilan Graphique ===
    With ThisWorkbook
        .Sheets("Bilan Graphique").Range("AC2:AC4998").Value = .Sheets("Données").Range("B4:B5000").Value
        .Sheets("Bilan Graphique").Range("AD2:AD4998").Value = .Sheets("Données").Range("C4:C5000").Value
    End With
    
    ' Ajuster la mise en forme (facultatif)
    wsBilan.Columns("A:L").AutoFit
    
    Call CreerHistogrammeNombrePalettes
    Call CreerHistogrammeNombreCamions
    Call CreerHistogrammeRemplissageCamions
    Call CreerGraphiqueFluxMensuel
    Call CreerLivrable

End Sub




