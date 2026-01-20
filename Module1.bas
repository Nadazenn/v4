Attribute VB_Name = "Module1"
Sub CreerTableauSource()
    Dim wsDonnees As Worksheet, wsSource As Worksheet, wsMateriel As Worksheet, wsParametres As Worksheet, wsManutention As Worksheet, wsCamion As Worksheet
    Dim lastRowDonnees As Long, lastRowMateriel As Long, lastRowSource As Long
    Dim i As Long, j As Long, k As Long
    Dim categorie As String, materiel As String, unite As String, typeCamion As String, nbPalettes As Long
    Dim etage As String
    Dim quantiteParCond As Double, quantite As Double, nbCond As Double
    Dim materielTrouve As Boolean
    Dim listPhase As Variant
    Dim filteredList As Object
    Dim cell As Range
    Dim resultString As String

    ' Définir les feuilles de travail
    Set wsDonnees = ThisWorkbook.Sheets("Données")
    Set wsSource = ThisWorkbook.Sheets("Tableau Source")
    Set wsMateriel = ThisWorkbook.Sheets("Matériel")
    Set wsManutention = ThisWorkbook.Sheets("Manutention")
    Set wsParametres = ThisWorkbook.Sheets("Paramétrage")
    Set wsCamion = ThisWorkbook.Sheets("Camion")
    
    ' Trouver la dernière ligne dans le bordereau
    lastRowDonnees = wsDonnees.Cells(wsDonnees.Rows.Count, "C").End(xlUp).Row
    lastRowMateriel = wsMateriel.Cells(wsMateriel.Rows.Count, "C").End(xlUp).Row
    lastRowManutention = wsManutention.Cells(wsManutention.Rows.Count, "A").End(xlUp).Row
    lastRowCamion = wsCamion.Cells(wsCamion.Rows.Count, "A").End(xlUp).Row
    
    
    
    ' Effacer tout le contenu et les formats
    wsSource.Cells.ClearContents
    wsSource.Cells.ClearFormats
    wsSource.Cells.Validation.Delete
    wsSource.Columns("P").FormatConditions.Delete
    wsSource.Columns("O").FormatConditions.Delete
    
    ' En-têtes
    wsSource.Cells(1, 1).Value = "Etage"
    wsSource.Cells(1, 2).Value = "Zone"
    wsSource.Cells(1, 3).Value = "Lot"
    wsSource.Cells(1, 4).Value = "Phase de traveaux"
    wsSource.Cells(1, 5).Value = "Nom de l'élément"
    wsSource.Cells(1, 6).Value = "Unité"
    wsSource.Cells(1, 7).Value = "Quantité"
    wsSource.Cells(1, 8).Value = "Conditionnement"
    wsSource.Cells(1, 9).Value = "Quantité par UM"
    wsSource.Cells(1, 10).Value = "Nombre d'UM nécessaires"
    wsSource.Cells(1, 11).Value = "Nombre palettes equivalent total"
    wsSource.Cells(1, 12).Value = "Type de camion requis"
    wsSource.Cells(1, 13).Value = "Nombre de camions nécessaires"
    wsSource.Cells(1, 14).Value = "Dont camions pleins"
    wsSource.Cells(1, 15).Value = "Remplissage camion non plein"
    wsSource.Cells(1, 16).Value = "Utilisation d'une CCC"

    
    ' Mise en forme : mettre les en-têtes en gras
    wsSource.Rows(1).Font.Bold = True
    wsSource.Rows(1).Font.Color = RGB(255, 255, 255)
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, 16)).Interior.Color = RGB(0, 32, 96)
    wsSource.Columns(1).Font.Bold = True
    wsSource.Columns(1).Font.Color = RGB(255, 255, 255)
    wsSource.Columns(1).Interior.Color = RGB(0, 32, 96)
    wsSource.Columns(2).Font.Bold = True
    
    lastColumnDonnees = wsDonnees.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    Debug.Print lastColumnDonnees
        ' Boucle sur chaque étage dans le tableau source (à partir de la 3e colonne)
    For j = 5 To (lastColumnDonnees - 2)
        If wsDonnees.Cells(1, j).Value <> "" Then
            etage = wsDonnees.Cells(1, j).Value
        End If
        zone = wsDonnees.Cells(2, j).Value
        quantite = 0
        ' Boucle sur chaque matériel dans le tableau source
        For i = 4 To lastRowDonnees
            categorie = wsDonnees.Cells(i, 2).Value
            quantite = wsDonnees.Cells(i, j).Value
            
            
            If quantite > 0 Then
                ' Rechercher les informations complémentaires dans le tableau base de données
                For k = 2 To lastRowMateriel
                    If wsMateriel.Cells(k, 1).Value = categorie Then
                        lot = wsMateriel.Cells(k, 2).Value
                        
                        If lot = wsParametres.Cells(1, 2).Value Then
                             materiel = wsMateriel.Cells(k, 1).Value
                             unite = wsMateriel.Cells(k, 3).Value
                             phase = wsMateriel.Cells(k, 4).Value
                             utilisationCCC = wsMateriel.Cells(k, 5).Value
                             
                             lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row + 1
                             wsSource.Cells(lastRowSource, 1).Value = etage
                             wsSource.Cells(lastRowSource, 2).Value = zone
                             wsSource.Cells(lastRowSource, 3).Value = lot
                             wsSource.Cells(lastRowSource, 4).Value = phase
                             wsSource.Cells(lastRowSource, 5).Value = materiel
                             wsSource.Cells(lastRowSource, 6).Value = unite
                             wsSource.Cells(lastRowSource, 7).Value = quantite
                             
                             If Not IsEmpty(wsMateriel.Cells(k, 7).Value) Then
                                wsSource.Cells(lastRowSource, 8).Value = wsMateriel.Cells(k, 7).Value
                             ElseIf Not IsEmpty(wsMateriel.Cells(k, 9).Value) Then
                                wsSource.Cells(lastRowSource, 8).Value = wsMateriel.Cells(k, 9).Value
                             ElseIf Not IsEmpty(wsMateriel.Cells(k, 11).Value) Then
                                wsSource.Cells(lastRowSource, 8).Value = wsMateriel.Cells(k, 11).Value
                             Else
                                wsSource.Cells(lastRowSource, 8).Value = "Palette"
                             End If
                             ' Ajoute la liste déroulante avec des options modifiables
                             With wsSource.Cells(lastRowSource, 8).Validation
                                 .Delete ' Supprime toute validation existante
                                 .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                 xlBetween, Formula1:="='Manutention'!A2:A" & lastRowManutention ' Options de la liste déroulante
                                 .IgnoreBlank = True
                                 .InCellDropdown = True
                             End With
                             
                             
                             wsSource.Cells(lastRowSource, 9).Formula = "=INDEX(Matériel!A:Z, MATCH(E" & lastRowSource & ", Matériel!A:A, 0), MATCH(H" & lastRowSource & ", INDEX(Matériel!A:Z, MATCH(E" & lastRowSource & ", Matériel!A:A, 0), 0), 0) + 1)"

                             wsSource.Cells(lastRowSource, 10).Formula = "=ROUNDUP(G" & lastRowSource & " / I" & lastRowSource & "  , 0)"
            
                             ' Utiliser RECHERCHEV pour les informations complémentaires (nombre palettes equivalent, type camion)
                             wsSource.Cells(lastRowSource, 11).Formula = "=VLOOKUP(H" & lastRowSource & ",'Manutention'!A:E, 3, FALSE) * J" & lastRowSource
                             
                             Application.EnableEvents = False ' Désactive les événements pour éviter les boucles

                             ' Détermine le type de camion requis
                             nbPalettes = wsSource.Cells(lastRowSource, 11).Value
                             typeCamion = Application.WorksheetFunction.VLookup(wsSource.Cells(lastRowSource, 8).Value, wsManutention.Range("A:E"), 2, False)  'Type de camion requis
                                                             
                             ' Appeler la fonction OptimiserRemplissage et remplir la cellule de la colonne 12
                             wsSource.Cells(lastRowSource, 12).Value = OptimiserRemplissage(lastRowSource, typeCamion, nbPalettes)
                             
                             Application.EnableEvents = True ' Réactive les événements
                            
                             ' Initialisation des objets
                             Set filteredList = CreateObject("Scripting.Dictionary") ' Utilisation d'un dictionnaire pour éviter les doublons
                             resultString = ""
                            
                             ' Parcourir la feuille Camion pour trouver les valeurs correspondantes au type de camion
                             For Each cell In wsCamion.Range("A2:A" & lastRowCamion)
                                If wsCamion.Cells(cell.Row, "B").Value = typeCamion Then
                                    filteredList(wsCamion.Cells(cell.Row, "A").Value) = True ' Ajoute les valeurs à un dictionnaire
                                End If
                             Next cell
                            
                             ' Construire la chaîne pour la validation des données
                             If filteredList.Count > 0 Then
                                For Each key In filteredList.keys
                                    resultString = resultString & key & ","
                                Next key
                                resultString = Left(resultString, Len(resultString) - 1) ' Retirer la dernière virgule
                             Else
                                resultString = "Aucun" ' Si aucune correspondance trouvée
                             End If
                            
                             ' Ajoute la liste déroulante avec les options filtrées
                             With wsSource.Cells(lastRowSource, 12).Validation
                                .Delete ' Supprime toute validation existante
                                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                xlBetween, Formula1:=resultString ' Options de la liste déroulante
                                .IgnoreBlank = True
                                .InCellDropdown = True
                             End With

                             
                             wsSource.Cells(lastRowSource, 13).Formula = "=ROUNDUP(K" & lastRowSource & "/VLOOKUP(L" & lastRowSource & ",'Camion'!A:D, 3, FALSE),0)"
                             wsSource.Cells(lastRowSource, 14).Formula = "=INT(K" & lastRowSource & "/VLOOKUP(L" & lastRowSource & ",'Camion'!A:D, 3, FALSE))"
                             
                             
                             wsSource.Cells(lastRowSource, 15).Formula = "=ROUND(K" & lastRowSource & "/VLOOKUP(L" & lastRowSource & ",'Camion'!A:D, 3, FALSE) - N" & lastRowSource & ",2)"
                             wsSource.Cells(lastRowSource, 15).NumberFormat = "0%"
                             
                             Application.EnableEvents = False ' Désactive les événements pour éviter les boucles
                             wsSource.Cells(lastRowSource, 16).Value = utilisationCCC
                             Application.EnableEvents = True ' Réactive les événements
                     
                             ' Ajoute la liste déroulante avec des options modifiables
                             With wsSource.Cells(lastRowSource, 16).Validation
                                 .Delete ' Supprime toute validation existante
                                 .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                 xlBetween, Formula1:="Oui,Non"  ' Options de la liste déroulante
                                 .IgnoreBlank = True
                                 .InCellDropdown = True
                             End With
                             
                             wsSource.Range(wsSource.Cells(lastRowSource, 2), wsSource.Cells(lastRowSource, 16)).Interior.Color = RGB(255, 255, 255)
                         End If
                         Exit For
                    End If
                Next k
            End If
        Next i
        
        listPhase = Array("Production", "Terminaux")
        For Each phase In listPhase
            lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row + 1
            wsSource.Cells(lastRowSource, 1).Value = etage
            wsSource.Cells(lastRowSource, 2).Value = zone
            wsSource.Cells(lastRowSource, 5).Value = "Stock CCC " & phase
            wsSource.Cells(lastRowSource, 8).Value = "Palette"
            
            wsSource.Cells(lastRowSource, 11).Formula = "=SUMIFS(K:K, A:A, " & etage & ", B:B, """ & zone & """, P:P, ""Oui"", D:D, """ & phase & """)"
            
            Application.EnableEvents = False
            
            ' Détermine le type de camion requis
            nbPalettes = wsSource.Cells(lastRowSource, 11).Value
            typeCamion = Application.WorksheetFunction.VLookup(wsSource.Cells(lastRowSource, 8).Value, wsManutention.Range("A:E"), 2, False)  'Type de camion requis
                                                             
            ' Appeler la fonction OptimiserRemplissage et remplir la cellule de la colonne 12
            wsSource.Cells(lastRowSource, 12).Value = OptimiserRemplissage(lastRowSource, typeCamion, nbPalettes)
            
            Application.EnableEvents = True ' Réactive les événements
            
            ' Ajoute la liste déroulante avec des options modifiables
            With wsSource.Cells(lastRowSource, 12).Validation
                .Delete ' Supprime toute validation existante
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='Camion'!A2:A11"   ' Options de la liste déroulante
                .IgnoreBlank = True
                .InCellDropdown = True
            End With
            
            wsSource.Cells(lastRowSource, 13).Formula = "=ROUNDUP(K" & lastRowSource & "/VLOOKUP(L" & lastRowSource & ",'Camion'!A:D, 3, FALSE),0)"
            wsSource.Cells(lastRowSource, 14).Formula = "=INT(K" & lastRowSource & "/VLOOKUP(L" & lastRowSource & ",'Camion'!A:D, 3, FALSE))"
                                
            wsSource.Cells(lastRowSource, 15).Formula = "=ROUND(K" & lastRowSource & "/VLOOKUP(L" & lastRowSource & ",'Camion'!A:D, 3, FALSE) - N" & lastRowSource & ",2)"
            wsSource.Cells(lastRowSource, 15).NumberFormat = "0%"
            
            wsSource.Range(wsSource.Cells(lastRowSource, 3), wsSource.Cells(lastRowSource, 16)).Interior.Color = RGB(169, 208, 142)
        Next phase

    Next j
    
     ' Règle pour "Oui" : Remplissage vert, texte vert foncé
    With wsSource.Columns("P").FormatConditions.Add(Type:=xlTextString, String:="Oui", TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206) ' Vert clair
        .Font.Color = RGB(0, 97, 0) ' Vert foncé
    End With

    ' Règle pour "Non" : Remplissage rouge, texte rouge foncé
    With wsSource.Columns("P").FormatConditions.Add(Type:=xlTextString, String:="Non", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 199, 206) ' Rouge
        .Font.Color = RGB(156, 0, 6) ' Rouge foncé
    End With

    ' Configurer les propriétés de la barre de données
    With wsSource.Columns("O").FormatConditions.AddDatabar
        ' Définir les valeurs minimales et maximales
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1

        ' Configurer la couleur de la barre
        .BarColor.Color = RGB(68, 114, 196) ' Bleu

        ' Activer ou désactiver les options d'affichage
        .ShowValue = True ' Afficher les valeurs dans les cellules
        .BarFillType = xlDataBarFillSolid ' Barres pleines
        .Direction = xlContext ' Direction par défaut selon la langue
    End With

    ' Adapter la largeur des colonnes
    wsSource.Columns.AutoFit
    wsSource.Range("A1").CurrentRegion.AutoFilter
    
End Sub









