Attribute VB_Name = "Module7"
Function OptimiserRemplissage(lastRowDonnees As Long, typeCamion As String, nbPalettes As Long) As String
    Dim wsCamion As Worksheet
    Dim i As Long, nbCamions As Long
    Dim capaciteCamion As Long
    Dim meilleurNbCamions As Long
    Dim meilleurTauxRemplissage As Double
    Dim tauxRemplissage As Double
    Dim palettesDernierCamion As Long
    Dim tauxDernierCamion As Double
    Dim camionChoisi As String
    Dim lastRowCamion As Long
    
    ' Définir la feuille "Camion"
    Set wsCamion = ThisWorkbook.Sheets("Camion")
    lastRowCamion = wsCamion.Cells(wsCamion.Rows.Count, "A").End(xlUp).Row
    
    ' Initialiser les variables
    meilleurNbCamions = 9999   ' Valeur élevée pour comparer
    meilleurTauxRemplissage = 0
    camionChoisi = ""
    
    ' Parcourir les camions
    For i = 2 To lastRowCamion
        ' Vérifier si le type de camion correspond
        If wsCamion.Cells(i, 2).Value = typeCamion Then
            capaciteCamion = wsCamion.Cells(i, 3).Value
            
            ' Calculer le nombre de camions nécessaires (au moins 1)
            nbCamions = Application.WorksheetFunction.Max(1, _
                          Application.WorksheetFunction.Ceiling(nbPalettes / capaciteCamion, 1))
            
            ' Calculer le taux de remplissage moyen
            tauxRemplissage = (nbPalettes / (nbCamions * capaciteCamion)) * 100
            
            ' Calculer le nombre de palettes dans le dernier camion
            palettesDernierCamion = nbPalettes Mod capaciteCamion
            If palettesDernierCamion = 0 Then
                tauxDernierCamion = 100
            Else
                tauxDernierCamion = (palettesDernierCamion / capaciteCamion) * 100
            End If
            
            ' Si le nombre de palettes est égal à 1, on sélectionne le camion qui offre le meilleur taux de remplissage
            If nbPalettes = 1 Then
                If tauxRemplissage > meilleurTauxRemplissage Then
                    camionChoisi = wsCamion.Cells(i, 1).Value
                    meilleurTauxRemplissage = tauxRemplissage
                End If
            Else
                ' Pour nbPalettes > 1, on choisit le camion en fonction du nombre de camions nécessaires et du taux de remplissage
                If nbCamions < meilleurNbCamions Or _
                   (nbCamions = meilleurNbCamions And tauxRemplissage > meilleurTauxRemplissage And tauxDernierCamion > 50) Then
                    camionChoisi = wsCamion.Cells(i, 1).Value
                    meilleurNbCamions = nbCamions
                    meilleurTauxRemplissage = tauxRemplissage
                End If
            End If
        End If
    Next i
    
    ' Retourner le camion choisi
    OptimiserRemplissage = camionChoisi
End Function




