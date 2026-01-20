Attribute VB_Name = "Module6"
Sub CreerLivrable()
    Dim ws As Worksheet, wsBilan As Worksheet, wsParametrage As Worksheet, wsBilanGraphique As Worksheet, wsSource As Worksheet
    Dim total As Double, totalProd As Double, totalTerm As Double
    Dim lastRowSource As Long, lastRowParametrage As Long, lastRowBilanGraphique As Long, lastRowBilanGraphique2 As Long
    Dim moyenneDelai As Double
    Dim totalCamion As Double, totalCamionProd As Double, totalCamionTerm As Double
    Dim remplissageMoyen As Double, moisPic As String, volumePic As Double, camionPic As Double
    Dim totalCamionCCC As Double, totalCamionProdCCC As Double, totalCamionTermCCC As Double
    Dim remplissageMoyenCCC As Double, moisPicCCC As String, volumePicCCC As Double, camionPicCCC As Double
    Dim stockCCC As Double, remplissageAmelioration  As Double, camionAmelioration As Double
    Dim plateauGruePetit As Double, plateauGrueGrand As Double, camionnette As Double, fourgon8 As Double, fourgon12 As Double, camion20 As Double, camion30 As Double, porteur12 As Double, porteur19 As Double, semiremorque As Double
    Dim i As Long
    Dim ListeMaterielOpti As String
    Dim Cellule As Range
    
    ' Définir les feuilles et les variables
    Set ws = ThisWorkbook.Sheets("Livrable")
    Set wsBilan = ThisWorkbook.Sheets("Bilan")
    Set wsParametrage = ThisWorkbook.Sheets("Paramétrage")
    Set wsBilanGraphique = ThisWorkbook.Sheets("Bilan Graphique")
    Set wsSource = ThisWorkbook.Sheets("Tableau Source")
    lastRowSource = wsBilan.Range("C" & wsBilan.Rows.Count).End(xlUp).Row
    lastRowBilanGraphique = wsBilanGraphique.Range("M" & wsBilanGraphique.Rows.Count).End(xlUp).Row
    lastRowBilanGraphique2 = wsBilanGraphique.Range("AA" & wsBilanGraphique.Rows.Count).End(xlUp).Row
    



    With Union(ws.Range("H1:H4"), ws.Range("P1:P4"), ws.Range("X1:X4"))
        .Merge
        .Characters(1, 15).Font.Bold = False
        .Value = Date
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 10  ' Définit la police des autres éléments à 14
    End With
    
    total = wsBilan.Range("D" & lastRowSource).Value
    totalProd = wsBilan.Range("D" & (lastRowSource - 2)).Value
    totalTerm = wsBilan.Range("D" & (lastRowSource - 1)).Value
    lastRowParametrage = wsParametrage.Range("J" & wsParametrage.Rows.Count).End(xlUp).Row
    
    ' Calcul de la moyenne des délais de livraison
    moyenneDelai = Application.WorksheetFunction.Average(wsParametrage.Range("J3:J" & lastRowParametrage))

    ' Préparer le texte avec les valeurs de variables
    Dim texte As String
    texte = total & " palettes équivalentes" & vbCrLf & vbCrLf & _
            Format(total * 1.2 * 0.8, "0.00") & " m² occupé au sol" & vbCrLf & vbCrLf & _
            " Palette Européenne (80 x 120 cm) :"

    ' Fusionner et formater la zone de texte dans la plage C22:H42
    With ws.Range("A13:C19")
        .Merge
        .Value = texte
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 10  ' Définit la police des autres éléments à 14
    End With


    
    
    totalCamion = wsBilan.Range("I" & lastRowSource).Value
    totalCamionProd = wsBilan.Range("I" & (lastRowSource - 2)).Value
    totalCamionTerm = wsBilan.Range("I" & (lastRowSource - 1)).Value
    remplissageMoyen = wsBilan.Range("K" & lastRowSource).Value
    
    camionPic = Application.WorksheetFunction.Max(wsBilanGraphique.Range("O2:O" & lastRowBilanGraphique)) ' Colonne O pour les camions
    
    ' Recherchez la ligne où le nombre de camions est maximum et obtenez les valeurs correspondantes
    For i = 2 To lastRowBilanGraphique ' On part de la ligne 2
        If wsBilanGraphique.Cells(i, 15).Value = camionPic Then
            moisPic = wsBilanGraphique.Cells(i, 13).Value2      ' Colonne M pour les dates/mois
            volumePic = wsBilanGraphique.Cells(i, 14).Value   ' Colonne N pour les volumes
            Exit For
        End If
    Next i

    Dim texteBase As String, texteBase2 As String
    texteBase = "Pic de livraison :" & vbCrLf & _
            "En " & Format(moisPic, "mmmm yyyy") & ", " & Application.WorksheetFunction.RoundUp(camionPic / 4, 0) & " camions/semaine"

    texteBase2 = totalCamion & " camions" & vbCrLf & _
            "Remplissage moyen : " & Format(remplissageMoyen, "0%")
    
          
    With ws.Range("A34:E36")
        .Merge
        .Characters.Font.Bold = True
        .Value = texteBase
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 12  ' Définit la police des autres éléments à 14
    End With
    
    With ws.Range("F34:H36")
        .Merge
        .Characters.Font.Bold = True
        .Value = texteBase2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 12  ' Définit la police des autres éléments à 14
    End With


    

    ' Initialiser la liste
    ListeMateriel = "Matériels stockés en CCC : "
    If wsBilanGraphique.Range("AA2") <> "" Then
        ' Parcourir les cellules de la colonne pour construire la liste
        For Each Cellule In wsBilanGraphique.Range("AA2:AA" & lastRowBilanGraphique2)
                ListeMateriel = ListeMateriel & Cellule.Value & ", " ' Ajouter avec une virgule
        Next Cellule
        ListeMateriel = Left(ListeMateriel, Len(ListeMateriel) - 2)
    Else
        ListeMateriel = ListeMateriel & "Aucun matériel stocké en CCC"
    End If
    
    With ws.Range("I16:P20")
        .Merge
        .Characters.Font.Bold = False
        .Value = ListeMateriel
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 10  ' Définit la police des autres éléments à 14
    End With
    
    With ws.Range("I16")
        .Characters(1, 26).Font.Bold = True
        .Font.Size = 10  ' Définit la police des autres éléments à 14
    End With
    
    totalCamionCCC = wsBilan.Range("J" & lastRowSource).Value
    totalCamionProdCCC = wsBilan.Range("J" & (lastRowSource - 2)).Value
    totalCamionTermCCC = wsBilan.Range("J" & (lastRowSource - 1)).Value
    remplissageMoyenCCC = wsBilan.Range("L" & lastRowSource).Value
    
    
    camionPicCCC = Application.WorksheetFunction.Max(wsBilanGraphique.Range("P2:P" & lastRowBilanGraphique)) ' Colonne O pour les camions
    
    ' Recherchez la ligne où le nombre de camions est maximum et obtenez les valeurs correspondantes
    For i = 2 To lastRowBilanGraphique ' On part de la ligne 2
        If wsBilanGraphique.Cells(i, 15).Value = camionPic Then
            moisPicCCC = wsBilanGraphique.Cells(i, 13).Value2      ' Colonne M pour les dates/mois
            volumePicCCC = wsBilanGraphique.Cells(i, 14).Value   ' Colonne N pour les volumes
            Exit For
        End If
    Next i
    
    Dim texteOpti As String, texteOpti2 As String
    texteOpti = "Pic de livraison :" & vbCrLf & _
            "En " & Format(moisPicCCC, "mmmm yyyy") & ", " & Application.WorksheetFunction.RoundUp(camionPicCCC / 4, 0) & " camions/semaine"

    texteOpti2 = totalCamionCCC & " camions" & vbCrLf & _
            "Remplissage moyen : " & Format(remplissageMoyenCCC, "0%")
    
          
    With ws.Range("I34:M36")
        .Merge
        .Characters.Font.Bold = True
        .Value = texteOpti
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 12  ' Définit la police des autres éléments à 14
    End With
    
    With ws.Range("N34:P36")
        .Merge
        .Characters.Font.Bold = True
        .Value = texteOpti2
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 12  ' Définit la police des autres éléments à 14
    End With

    
    
    
    
    ws.Range("H23").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Plateau Grue Petit")
    ws.Range("H24").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Plateau Grue Grand")
    ws.Range("H25").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Camionnette 6m3")
    ws.Range("H26").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Fourgon 7-9m3")
    ws.Range("H27").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Fourgon 10-12m3")
    ws.Range("H28").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Camion 20-22 m3")
    ws.Range("H29").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Camion 30 m3")
    ws.Range("H30").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Porteur 12 tonnes")
    ws.Range("H31").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Porteur 19 tonnes")
    ws.Range("H32").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(21), wsBilanGraphique.Columns(20), "Semi-remorque 90 m3")


    With ws.Range("H23:H32")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 10
    End With
    
    
    ws.Range("P23").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Plateau Grue Petit")
    ws.Range("P24").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Plateau Grue Grand")
    ws.Range("P25").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Camionnette 6m3")
    ws.Range("P26").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Fourgon 7-9m3")
    ws.Range("P27").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Fourgon 10-12m3")
    ws.Range("P28").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Camion 20-22 m3")
    ws.Range("P29").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Camion 30 m3")
    ws.Range("P30").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Porteur 12 tonnes")
    ws.Range("P31").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Porteur 19 tonnes")
    ws.Range("P32").Value = Application.WorksheetFunction.SumIfs(wsBilanGraphique.Columns(25), wsBilanGraphique.Columns(24), "Semi-remorque 90 m3")

    With ws.Range("P23:P32")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 10
    End With
    
    
    
    stockCCC = Application.WorksheetFunction.SumIfs(wsSource.Columns(11), wsSource.Columns(5), "Stock CCC Production") + _
                Application.WorksheetFunction.SumIfs(wsSource.Columns(11), wsSource.Columns(5), "Stock CCC Terminaux")
    camionAmelioration = Abs(((totalCamionCCC - totalCamion) / totalCamion))
    remplissageAmelioration = ((remplissageMoyenCCC - remplissageMoyen) / remplissageMoyen)
    dureeCCC = wsParametrage.Range("B4").Value
    prixstockageCCC = (wsParametrage.Range("B5").Value * dureeCCC + wsParametrage.Range("B6").Value) * stockCCC
    prixlivraisonCCC = wsParametrage.Range("B7").Value * stockCCC / 9
    prixCCC = prixstockageCCC + prixlivraisonCCC
    
    
    With wsBilanGraphique
        ' ===== LIBELLÉS (ligne 1) =====
        .Range("AE1").Value = "% Stock CCC"
        .Range("AF1").Value = "% réduction Camions"
        .Range("AG1").Value = "% remplissage moyen des camions"
        .Range("AH1").Value = "Coût CCC stockage"
        .Range("AI1").Value = "Coût CCC livraison"
        .Range("AJ1").Value = "Coût CCC Total"
    
        ' ===== VALEURS (ligne 2) =====
        .Range("AE2").Value = stockCCC / total
        .Range("AF2").Value = camionAmelioration
        .Range("AG2").Value = remplissageAmelioration
        .Range("AH2").Value = prixstockageCCC
        .Range("AI2").Value = prixlivraisonCCC
        .Range("AJ2").Value = prixCCC
    
        ' ===== FORMATS =====
        .Range("AE2:AG2").NumberFormat = "0%"
        .Range("AH2:AJ2").NumberFormat = "#,##0 €"
    End With

    
    
    
    
    
    Dim texteAmeliorations As String
    texteAmeliorations = "Avec " & Format(stockCCC / total, "0%") & " du matériel stocké" & vbCrLf & _
            "dans un CCC pendant " & vbCrLf & _
            dureeCCC & " mois :" & vbCrLf & vbCrLf & _
            " - " & Format(camionAmelioration, "0%") & " de camions" & vbCrLf & vbCrLf & _
            " " & IIf(remplissageAmelioration >= 0, "+", "-") & Format(Abs(remplissageAmelioration), "0%") & " du" & vbCrLf & _
            "remplissage moyen"

    
    With ws.Range("Q15:R29")
        .Merge
        .Characters.Font.Bold = True
        .Value = texteAmeliorations
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 14  ' Définit la police des autres éléments à 14
    End With
    
    Dim texteCCC As String
    texteCCC = "Coût CCC : " & vbCrLf & vbCrLf & _
            "Stockage : " & Format(prixstockageCCC, 0) & "€" & vbCrLf & _
            "Livraison : " & Format(prixlivraisonCCC, 0) & "€" & vbCrLf & vbCrLf & _
            "Total : " & Format(prixCCC, 0) & "€"

    
    With ws.Range("Q30:R37")
        .Merge
        .Characters.Font.Bold = False
        .Value = texteCCC
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Font.Size = 12  ' Définit la police des autres éléments à 14
    End With
    
    With ws.Range("Q30")
        .Characters(1, 10).Font.Bold = True
        .Characters(1, 10).Font.Size = 14  ' Définit la police des autres éléments à 14
    End With


End Sub










