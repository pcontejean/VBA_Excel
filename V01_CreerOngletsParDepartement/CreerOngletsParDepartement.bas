Attribute VB_Name = "Module1"
Sub CreerOnglets()

    Dim wsData As Worksheet          ' Variable pour l’onglet Données
    Dim wsNew As Worksheet           ' Variable pour les nouveaux onglets
    Dim lastRow As Long              ' Dernière ligne utilisée
    Dim dept As String               ' Nom du département
    Dim i As Long                    ' Compteur de boucle
    
    ' Définir l’onglet Données
    Set wsData = Worksheets("Données")
    
    ' Aller sur l’onglet Données
    wsData.Activate
    
    ' Trouver la dernière ligne remplie dans la colonne I
    lastRow = wsData.Cells(wsData.Rows.Count, "I").End(xlUp).Row
    
    ' Trier les données par la colonne I (Département)
    wsData.Sort.SortFields.Clear
    wsData.Sort.SortFields.Add Key:=wsData.Range("I2:I" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsData.Sort
        .SetRange wsData.Range("A1:I" & lastRow) ' Plage à trier
        .Header = xlYes                           ' Il y a une ligne d’en-tête
        .Apply                                   ' Appliquer le tri
    End With
    
    ' Boucle sur chaque ligne pour détecter les départements
    For i = 2 To lastRow
        
        ' Récupérer le nom du département
        dept = wsData.Cells(i, "I").Value
        
        ' Vérifier que le département n’est pas vide
        If dept <> "" Then
            
            ' Appliquer un filtre sur le département
            wsData.Range("A1:I" & lastRow).AutoFilter Field:=9, Criteria1:=dept
            
            ' Créer un nouvel onglet à droite de l’onglet actif
            Set wsNew = Worksheets.Add(After:=ActiveSheet)
            
            ' Renommer le nouvel onglet avec le nom du département
            wsNew.Name = dept
            
            ' Copier les données filtrées (avec en-têtes)
            wsData.Range("A1:I" & lastRow).SpecialCells(xlCellTypeVisible).Copy
            
            ' Coller les données dans le nouvel onglet
            wsNew.Range("A1").PasteSpecial xlPasteValues
            
            ' Désactiver le mode copie
            Application.CutCopyMode = False
            
            ' Supprimer le filtre
            wsData.AutoFilterMode = False
            
            ' Sauter les lignes du même département déjà traité
            Do While i < lastRow And wsData.Cells(i + 1, "I").Value = dept
                i = i + 1
            Loop
            
        End If
        
    Next i
    
    ' Message de fin
    MsgBox "Création des onglets par département terminée.", vbInformation




 
End Sub
Sub SupprimerOnglets()

    Dim ws As Worksheet   ' Variable pour parcourir les onglets
    
    ' Désactiver les messages de confirmation Excel
    ' (évite le message "Voulez-vous vraiment supprimer cette feuille ?")
    Application.DisplayAlerts = False
    
    ' Boucle sur tous les onglets du classeur
    For Each ws In ThisWorkbook.Worksheets
        
        ' Vérifier que l’onglet n’est PAS Menu ni Données
        If ws.Name <> "Menu" And ws.Name <> "Données" Then
            
            ' Supprimer l’onglet
            ws.Delete
            
        End If
        
    Next ws
    
    ' Réactiver les messages de confirmation Excel
    Application.DisplayAlerts = True
    
    ' Message de fin
    MsgBox "Tous les onglets de départements ont été supprimés.", vbInformation

End Sub


