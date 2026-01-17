Attribute VB_Name = "ModConstant"
' Référence au classeur principal
Public wb As Workbook

' Référence à la feuille contenant les données de stock
Public wsStock As Worksheet

' Tableau structuré nommé (ListObject) dans la feuille de stock
Public tabStock As ListObject

' Ligne nouvellement ajoutée dans le tableau de stock
Public newStockRow As ListRow

' Plage représentant le tableau de stock
Public rangeStock As Range

' Plage représentant l'adresse du tableau de stock
Public rangeStockAddress As String

' Référence à la feuille contenant les mouvements de stock
Public wsMovement As Worksheet

' Tableau structuré nommé (ListObject) dans la feuille de mouvement
Public tabMovement As ListObject

' Ligne nouvellement ajoutée dans le tableau des mouvement
Public newMovementRow As ListRow

' Plage représentant le tableau des mouvement
Public rangeMovement As Range

' Plage représentant l'adresse du tableau de mouvement
Public rangeMovementAdress As String
