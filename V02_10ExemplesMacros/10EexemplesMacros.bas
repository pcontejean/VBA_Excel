Attribute VB_Name = "Module1"
Option Explicit 'sert à vous obliger à déclarer vos variables

'Macro 1: Affiche un message à l'écran
Sub Bonjour()
    
    'MsgBox est une boîte de dialogue dans Excel qui sert à afficher un message à l'écran
    MsgBox "Bonjour et bienvenue sur Excel !"
End Sub

'Macro 2: Écrire du texte dans une cellule
Sub EcrireDansCellule()
    Range("A1").Value = "Bonjour Excel"
    MsgBox ("Fin")
End Sub

'Macro 3: Mettre en gras les cellules sélectionnées
Sub MettreEnGras()
    Selection.Font.Bold = True
End Sub

'Macro 4: Mettre une couleur de fond jaune dans la cellule B2
Sub ColorerCellule()
    Range("B2").Interior.Color = vbYellow
End Sub

'Macro 5: Effacer le contenu d'une plage
Sub EffacerContenu()
    Range("A1:A10").ClearContents
End Sub

'Macro 6 : Écrire la date du jour dans la cellule sélectionnée
Sub DateDuJour()
    ActiveCell.Value = Date
End Sub
'Macro 7 : Sub CopierCellule()
Sub CopierCellule()
    Range("B6").Copy Destination:=Range("C6")
End Sub
'Macro 8 : Remplit A1 à A5 avec des nombres
Sub RemplirColonne()
    Dim i As Integer
    For i = 1 To 5
        Cells(i, 1).Value = i
    Next i
End Sub
'Macro 9: Macro avec condition
Sub ConditionSimple()
    If Range("A1").Value > 10 Then
        Range("B1").Value = "Plus grand que 10"
    Else
        Range("B1").Value = "Inférieur à 10"
    End If
End Sub
'Macro 10: Supprime les espaces inutiles dans la colonne A
Sub NettoyerEspaces()
    Dim c As Range
    For Each c In Range("A1:A10")
        c.Value = Trim(c.Value)
    Next c
End Sub



