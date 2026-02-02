Attribute VB_Name = "Module1"
Option Explicit

Sub UniformiserFormatDates()

    Dim ws As Worksheet
    Dim DerniereLigne As Long
    Dim i As Long

    ' Définir la feuille de travail active
    Set ws = ActiveSheet

    ' Trouver la dernière ligne avec des données dans la colonne A
    DerniereLigne = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Parcourir toutes les cellules dans la colonne F
    For i = 2 To DerniereLigne
    ' Vérifier si la cellule contient une date ou un texte convertible en date
        If IsDate(ws.Cells(i, "F").Value) Then
            ' Convertir en vraie date (au cas où c'était du texte)
            ws.Cells(i, "F").Value = CDate(ws.Cells(i, "F").Value)
            ' Appliquer le format jj/mm/aaaa
            ws.Cells(i, "F").NumberFormat = "dd/mm/yyyy"
    End If
Next i

MsgBox ("Fin !")

End Sub

Sub UniformiserNombres()

    Dim DerniereLigne As Long
    Dim c As Range
    
    'Chercher la dernière ligne remplie
    DerniereLigne = Cells(Rows.Count, "E").End(xlUp).Row
    
    'Parcourir les lignes de la colonne E, une par une
    For Each c In Range("E1:E" & DerniereLigne)
        
        'Si dans la cellule on a un nombre qui est stocké en texte
        If IsNumeric(c.Value) And VarType(c.Value) = vbString Then
            'alors on convertit ce nombre stockée texte en un vrai nombre
            c.Value = CDbl(c.Value)  ' convertit en nombre
        End If

    Next c
    
    MsgBox ("Fin!")

End Sub

Sub MarquerAdressesCourrielsIncorrectes()

    Dim DerniereLigne As Long
    Dim c As Range

    DerniereLigne = Cells(Rows.Count, "D").End(xlUp).Row

    For Each c In Range("D2:D" & DerniereLigne)

        ' Si la cellule n'est pas vide
        If Trim(c.Value) <> "" Then

            ' Si le texte ne contient pas "@"
            If InStr(1, c.Value, "@") = 0 Then
                c.Interior.Color = RGB(255, 165, 0)    ' orange   ' met en jaune
            Else
                c.Interior.ColorIndex = xlNone ' remet en blanc si OK
            End If

        End If

    Next c
    MsgBox ("Fin!")

End Sub

Sub ConcatenerNomPrenom()

    Dim DerniereLigneRemplie As Long
    Dim i As Long

    ' Cells(Rows.Count, "A"): retourne la dernière ligne de la feuille Excel
    ' ! Ce n'est pas la dernière ligne remplie avec des données,
    ' Exemple: avec Excel 2007 ce serait la ligne 1 048 576
    '.End(xlUp).Row: remonte jusqu'a la dernière ligne remplie
    DerniereLigneRemplie = Cells(Rows.Count, "A").End(xlUp).Row
        
    ' Boucle à partir de la ligne 2
    For i = 2 To DerniereLigneRemplie
        Cells(i, "C").Value = Cells(i, "B").Value & " " & Cells(i, "A").Value
    Next i
    
    MsgBox ("Fin !")

End Sub




