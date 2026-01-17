Attribute VB_Name = "modStockModificationSub"
' ==============================================================================================
' Procédure : addItem
' Objectif  : Ajouter un nouvel matériel dans le tableau "stock"
' Paramètre : itemLabel ? texte du libellé du matériel à ajouter
' ==============================================================================================
Public Sub addItem(itemLabel As String)
     ' Classeur contenant la macro
    Set wb = ThisWorkbook
    ' Feuille de calcul "stock"
    Set wsStock = wb.Worksheets("stock")
    ' Tableau structuré nommé "stock"
    Set tabStock = wsStock.ListObjects("stock")
    ' Plage physique correspondant au tableau
    Set rangeStock = wsStock.Range("stock")

    ' Ajout d'une nouvelle ligne à la fin du tableau
    Set newStockRow = tabStock.ListRows.Add

    ' Remplissage des cellules de la nouvelle ligne
    ' Colonne 1 : libellé du matériel
    newStockRow.Range.Cells(1).Value = itemLabel

    ' Colonne 8 : numéro de ligne relatif dans le tableau (ROW() - 2 car en-têtes + ligne départ)
    newStockRow.Range.Cells(8).Formula = "=ROW()-2"

    ' Colonne 9 : numéro de ligne absolu dans la feuille
    newStockRow.Range.Cells(9).Formula = "=ROW()"
End Sub

' ==============================================================================================
' Procédure : addMovement
' Objectif  : Ajouter un nouvel enregistrement dans le tableau "movement"
' Paramètres:
'   moveDate        ? Date du mouvement
'   moveType        ? Type de mouvement ("Entrée" ou "Sortie")
'   moveValue       ? Quantité du mouvement
'   moveDescription ? Commentaire ou description
'   moveItem        ? Libellé du matériel concerné
' ==============================================================================================
Public Sub addMovement(moveDate As Date, moveType As String, moveValue As Variant, moveDescription As String, moveItem As String)
    ' Classeur contenant la macro
    Set wb = ThisWorkbook
    ' Feuille "mouvement"
    Set wsMovement = wb.Worksheets("mouvement")
    ' Tableau structuré "movement"
    Set tabMovement = wsMovement.ListObjects("movement")
    ' Plage physique du tableau
    Set rangeMovement = wsMovement.Range("movement")

    ' Ajout d'une nouvelle ligne
    Set newMovementRow = tabMovement.ListRows.Add

    ' Remplissage des cellules de la nouvelle ligne
    ' Colonne 1 : date du mouvement
    newMovementRow.Range.Cells(1).Value = moveDate

    ' Colonne 2 : type de mouvement
    newMovementRow.Range.Cells(2).Value = moveType

    ' Colonne 3 : quantité associée au mouvement
    newMovementRow.Range.Cells(3).Value = moveValue

    ' Colonne 4 : description
    newMovementRow.Range.Cells(4).Value = moveDescription

    ' Colonne 5 : libellé du matériel concerné
    newMovementRow.Range.Cells(5).Value = moveItem
End Sub
