VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStock 
   Caption         =   "UserForm1"
   ClientHeight    =   12225
   ClientLeft      =   -165
   ClientTop       =   -570
   ClientWidth     =   17160
   OleObjectBlob   =   "frmStock.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================================
' Procédure : UserForm_Initialize
' Objectif  : Initialiser l’interface et les variables du formulaire de gestion du stock
' Déclenchement : Automatique à l'ouverture du UserForm
' ==============================================================================================

Private Sub lblItemUpdateDate_Click()

End Sub

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------------------------
' Section pour déclaration des variables et initialisation des références
' ----------------------------------------------------------------------------------------------

' Référence au classeur actif (celui qui contient le code)
Set wb = ThisWorkbook

' Référence à la feuille "stock" (données des matériels en stock)
Set wsStock = wb.Worksheets("stock")

' Référence à la feuille "mouvement" (historique des entrées/sorties de stock)
Set wsMovement = wb.Worksheets("mouvement")

' Référence au tableau structuré nommé "stock" présent dans wsStock
Set tabStock = wsStock.ListObjects("stock")

' Référence à la plage de cellules couvrant le tableau "stock"
Set rangeStock = wsStock.Range("stock")

' Référence au tableau structuré nommé "movement" présent dans wsMovement
Set tabMovement = wsMovement.ListObjects("movement")

' Référence à la plage de cellules couvrant le tableau "movement"
Set rangeMovement = wsMovement.Range("movement")

' Variables servant au découpage d'adresse de plage et calcul de lignes pour "stock"
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long

' Variables servant au découpage d'adresse de plage et calcul de lignes pour "movement"
Dim rangeMovementAddressPart() As String
Dim rangeMovementLastLine As Long

' ----------------------------------------------------------------------------------------------
' Section pour définir le front-end du formulaire (dimensions, titre, couleurs)
' ----------------------------------------------------------------------------------------------

With Me
    ' Largeur totale du formulaire
    .Width = 900
    ' Hauteur totale du formulaire
    .Height = 520
    ' Titre affiché dans la barre du formulaire
    .Caption = "Stock du service informatique"
    ' Couleur d’arrière-plan générale
    .BackColor = COLOR_GRAY_DARK
End With

' ----------------------------------------------
' Section : Contrôles de recherche et filtres
' ----------------------------------------------

' Zone de texte de recherche matériel
With txtSearchItem
    .Left = 20
    .Top = 17
    .Width = 265
    .Height = 25
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_IRON
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Bouton de lancement de la recherche
With btnSearchItem
    .Left = 295
    .Top = 17
    .Width = 125
    .Height = 25
    .Caption = "Rechercher"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton filtrant les matériels en faible quantité
With btnFilterLowQuantity
    .Left = 430
    .Top = 17
    .Width = 120
    .Height = 25
    .Caption = "Quantités faibles"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_WHITE
End With

' ----------------------------------------------
' Section : Boutons de tri des données
' ----------------------------------------------

' Tri par libellé (nom matériel)
With btnSortItemLabel
    .Left = 20
    .Top = 50
    .Width = 125
    .Height = 25
    .Caption = "Trier/nom"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Tri par quantité en stock
With btnSortItemQuantity
    .Left = 155
    .Top = 50
    .Width = 130
    .Height = 25
    .Caption = "Trier/quantité"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Tri par catégorie
With btnSortItemCategory
    .Left = 295
    .Top = 50
    .Width = 125
    .Height = 25
    .Caption = "Trier/catégorie"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Tri par date de mise à jour
With btnSortItemUpdateDate
    .Left = 430
    .Top = 50
    .Width = 120
    .Height = 25
    .Caption = "Trier/date de MAJ"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' ----------------------------------------------
' Section : Liste des items (ListBox principale)
' ----------------------------------------------

With lstItems
    .Left = 20
    .Top = 80
    .Width = 531
    .Height = 380
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_IRON
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
    .SpecialEffect = fmSpecialEffectFlat
    
    ' Structure de la ListBox : nombre et largeur des colonnes
    .ColumnCount = 4
    .ColumnWidths = "190;50;170;115"
   
    ' Récupération de l'adresse du tableau "stock" pour déterminer la dernière ligne
    rangeStockAddress = rangeStock.Address
    ' Découpe l'adresse en parties ($A$1 ? {"","A","1"})
    rangeStockAddressPart = Split(rangeStockAddress, "$")
    ' Récupère le numéro de ligne de fin (le 4ème élément de la chaine découpée)
    rangeStockLastLine = CLng(rangeStockAddressPart(4))

    ' Remplissage de la ListBox avec les données du tableau
    ' On commence à la ligne 3 pour ignorer l’en-tête du tableau
    For i = 3 To rangeStockLastLine
        ' Colonne 0 : Libellé
        .addItem tabStock.Range.Cells(i - 1, 1)
        ' Colonne 1 : Quantité
        .List(.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2)
        ' Colonne 2 : Catégorie
        .List(.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3)
        ' Colonne 3 : Date / autre info
        .List(.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4)
    Next i
End With

' ----------------------------------------------
' Section : Détails du matériel (panneau de droite)
' ----------------------------------------------

' Bouton pour sauvegarder les modifications sur un item
With btnSaveItemUpdate
    .Left = 780
    .Top = 45
    .Width = 75
    .Height = 25
    .Caption = "Sauvegarder"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_SILVER_GLINT
    .ForeColor = COLOR_GRAY_DARK
    ' Désactivé par défaut tant qu'aucune modification n’est en cours
    .Enabled = False
End With

' Titre du panneau de détails
With lblItemDetail
    .Left = 580
    .Top = 50
    .Width = 200
    .Height = 25
    .Caption = "Détail du matériel"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Label du champ "Libellé"
With lblItemLabel
    .Left = 580
    .Top = 80
    .Width = 80
    .Height = 20
    .Caption = "Libellé"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés de la zone de texte pour le libellé du matériel
With txtItemLabel
    .Left = 675
    .Top = 80
    .Width = 180
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés du label pour indiquer la catégorie
With lblItemCategory
    .Left = 580
    .Top = 118
    .Width = 80
    .Height = 20
    .Caption = "Catégorie"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés de la liste déroulante pour choisir la catégorie
With cmbItemCategory
    .Left = 675
    .Top = 118
    .Width = 180
    .Height = 20
    .Style = fmStyleDropDownList
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
    .RowSource = "category"
End With

' Propriétés du label pour indiquer la sous-catégorie
With lblItemSubcategory
    .Left = 580
    .Top = 156
    .Width = 80
    .Height = 20
    .Caption = "Sous-catégorie"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés de la liste déroulante pour sélectionner la sous-catégorie
With cmbItemSubcategory
    .Left = 675
    .Top = 156
    .Width = 180
    .Height = 20
    .Style = fmStyleDropDownList
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés du label pour afficher le texte "En stock"
With lblItemCurrentQuantity
    .Left = 580
    .Top = 194
    .Width = 80
    .Height = 20
    .Caption = "En stock"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés de la zone de texte pour saisir la quantité actuelle
With txtItemCurrentQuantity
    .Left = 675
    .Top = 194
    .Width = 50
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés du label pour symboliser "quantité minimale" (>=)
With lblItemMinQuantity
    .Left = 750
    .Top = 194
    .Width = 60
    .Height = 20
    .Caption = ">="
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés de la zone de texte pour saisir la quantité minimale
With txtItemMinQuantity
    .Left = 804
    .Top = 194
    .Width = 50
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés du label pour la date de mise à jour
With lblItemUpdateDate
    .Left = 580
    .Top = 232
    .Width = 80
    .Height = 20
    .Caption = "Date de MAJ"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés de la zone de texte pour saisir ou afficher la date de mise à jour
With txtItemUpdateDate
    .Left = 675
    .Top = 232
    .Width = 180
    .Height = 22
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés du label pour le champ commentaire
With lblItemComment
    .Left = 580
    .Top = 270
    .Width = 80
    .Height = 20
    .Caption = "Commentaire"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propriétés de la zone de texte pour saisir un commentaire libre
With txtItemComment
    .Left = 675
    .Top = 270
    .Width = 180
    .Height = 22
    .MaxLength = 30
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' ----------------------------------------------
' Section : Historique des mouvements
' ----------------------------------------------

' Label titre de la section historique
With lblItemHistorical
    .Left = 580
    .Top = 310
    .Width = 200
    .Height = 25
    .Caption = "Historique des mouvements"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' ListBox affichant l'historique
With lstItemHistorical
    .Left = 580
    .Top = 335
    .Width = 275
    .Height = 155
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_EXTRA_SMALL
    .BackColor = COLOR_GRAY_IRON
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
    .SpecialEffect = fmSpecialEffectFlat
    .ColumnCount = 4
    .ColumnWidths = "70;45;30;125"
End With

' ----------------------------------------------
' Section : Boutons d'action
' ----------------------------------------------

' Bouton pour ajouter un nouveau matériel
With btnAddItem
    .Left = 20
    .Top = 450
    .Width = 170
    .Height = 35
    .Caption = "Nouveau"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_FOREST_GREEN
    .ForeColor = COLOR_WHITE
End With

' Bouton pour supprimer un matériel sélectionné
With btnDeleteItem
    .Left = 200
    .Top = 450
    .Width = 170
    .Height = 35
    .Caption = "Supprimer"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_CRIMSON_DARK
    .ForeColor = COLOR_WHITE
End With

' Bouton pour enregistrer un mouvement (entrée ou sortie) sur un matériel
With btnAddMovement
    .Left = 380
    .Top = 450
    .Width = 170
    .Height = 35
    .Caption = "Mouvement"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_NAVY_SLATE
    .ForeColor = COLOR_WHITE
End With
End Sub

' ----------------------------------------------------------------------------------------------
' Action : Ouvre le formulaire de gestion d'un nouveau matériel
' ----------------------------------------------------------------------------------------------
Private Sub btnAddItem_Click()
' Affiche le formulaire frmItem en mode modal
 frmItem.Show
End Sub

' ----------------------------------------------------------------------------------------------
' Action : Ouvre le formulaire d'enregistrement d'un nouveau mouvement (entrée ou sortie)
' ----------------------------------------------------------------------------------------------
Private Sub btnAddMovement_Click()
' Affiche le formulaire frmMovement en mode modal
frmMovement.Show
End Sub

' ----------------------------------------------------------------------------------------------
' Objectif : Vider et recharger la liste principale (lstItems) avec les données du tableau "stock"
' ----------------------------------------------------------------------------------------------
Private Sub displayItems()

' Déclaration des variables pour gérer les lignes et découper l'adresse du tableau
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long
Dim i As Long

' Récupère l'adresse de la plage "stock"
rangeStockAddress = rangeStock.Address

' Sépare l'adresse en parties (ex: "$A$1:$D$20" -> {"","A","1","","D","20"})
rangeStockAddressPart = Split(rangeStockAddress, "$")

' Convertit en nombre la dernière ligne du tableau (élément n°4 du tableau après split)
rangeStockLastLine = CLng(rangeStockAddressPart(4))

' Vide la liste avant de la recharger
lstItems.Clear

' Boucle à partir de la 3e ligne (pour sauter l’en-tête du tableau)
For i = 3 To rangeStockLastLine
    ' Ajoute un nouvel élément dans la première colonne
    lstItems.addItem tabStock.Range.Cells(i - 1, 1)
    ' Remplit les colonnes 2 à 4 avec les données correspondantes
    lstItems.List(lstItems.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2) ' Quantité
    lstItems.List(lstItems.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3) ' Catégorie
    lstItems.List(lstItems.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4) ' Date ou autre info
Next i

End Sub

' ----------------------------------------------------------------------------------------------
' Objectif : Afficher les détails du matériel sélectionné + son historique
' ----------------------------------------------------------------------------------------------
Private Sub lstItems_Change()

 ' Variables pour stocker les informations de l’matériel actif
 Dim activeItemLabel As String
 Dim activeItemCategory As String
 Dim activeItemSubcategory As String
 Dim activeItemCurrentQuantity As Integer
 Dim activeItemMinQuantity As Integer
 Dim activeItemUpdateDate As Date
 Dim activeItemComment As String
 Dim lastRow As Long

 ' Active le bouton de sauvegarde (un élément est sélectionné)
 btnSaveItemUpdate.Enabled = True

 ' ----------------------------
 ' Préparation de la liste historique
 ' ----------------------------

On Error Resume Next
 ' Récupère l'adresse du tableau "movement"
 rangeMovementAdress = rangeMovement.Address

 ' Sépare l'adresse en parties pour identifier la dernière ligne
 rangeMovementAddressPart = Split(rangeMovementAdress, "$")
 rangeMovementLastLine = CLng(rangeMovementAddressPart(4)) + 1

 ' ----------------------------
 ' Récupération des infos du matériel sélectionné
 ' ----------------------------
 ' Évite les erreurs si matériel est introuvable
 On Error Resume Next

 ' Libellé de l’matériel sélectionné
 activeItemLabel = lstItems.Value

 ' Permet d'actualiser le formulaire dynamique après enregistrement
 lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
 Set rangeStock = wsStock.Range("A3:G" & lastRow)

 If activeItemLabel = "" Then
     ' Aucun élément sélectionné réinitialise les champs
     txtItemLabel.Value = ""
     cmbItemCategory.Value = ""
     cmbItemSubcategory.Value = ""
     txtItemCurrentQuantity.Value = ""
     txtItemMinQuantity = ""
     txtItemUpdateDate = ""
     txtItemComment = ""
 Else
     ' Recherche des infos dans le tableau "stock" avec VLOOKUP
     activeItemCategory = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 3, False)
     activeItemSubcategory = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 6, False)
     activeItemCurrentQuantity = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 2, False)
     activeItemMinQuantity = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 5, False)
     activeItemUpdateDate = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 4, False)
     activeItemComment = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 7, False)

     ' Affichage dans les contrôles du formulaire
     txtItemLabel.Value = activeItemLabel
     cmbItemCategory.Value = activeItemCategory
     cmbItemSubcategory.Value = activeItemSubcategory
     
     ' N'affiche la quantité en stock du matériel seulement si un mouvement a déjà eu lieu
     If activeItemCurrentQuantity = 0 And activeItemUpdateDate = "00:00:00" Then
        txtItemCurrentQuantity.Value = ""
     Else
         txtItemCurrentQuantity.Value = activeItemCurrentQuantity
     End If
    
     txtItemMinQuantity.Value = activeItemMinQuantity
     
     ' N'affiche la date de la dernière mise à jour du matériel seulement si un mouvement a déjà eu lieu
     If activeItemUpdateDate = "00:00:00" Then
        txtItemUpdateDate.Value = ""
     Else
        txtItemUpdateDate.Value = activeItemUpdateDate
     End If
    
     txtItemComment = activeItemComment
     
     ' Remplissage de l’historique
     lstItemHistorical.Clear
     For i = 2 To rangeMovementLastLine + 2
         ' Vérifie si la colonne 5 du mouvement correspond à matériel sélectionné
         If tabMovement.Range.Cells(i - 1, 5) = activeItemLabel Then
            ' Date mouvement
             lstItemHistorical.addItem tabMovement.Range.Cells(i - 1, 1)
             ' Type (entrée/sortie)
             lstItemHistorical.List(lstItemHistorical.ListCount - 1, 1) = tabMovement.Range.Cells(i - 1, 2)
             ' Quantité
             lstItemHistorical.List(lstItemHistorical.ListCount - 1, 2) = tabMovement.Range.Cells(i - 1, 3)
             ' Commentaire
             lstItemHistorical.List(lstItemHistorical.ListCount - 1, 3) = tabMovement.Range.Cells(i - 1, 4)
         End If
     Next i
 End If
 
 Exit Sub
 ' Réinitialise la gestion des erreurs
 On Error GoTo 0

End Sub

' ----------------------------------------------------------------------------------------------
' Objectif : Modifie les options de sous-catégorie selon la catégorie choisie
' ----------------------------------------------------------------------------------------------
Private Sub cmbItemCategory_Change()

Dim categoryChoiced As String

categoryChoiced = cmbItemCategory.Value

If categoryChoiced = "accessoire" Then
    cmbItemSubcategory.RowSource = "accessorie"
ElseIf categoryChoiced = "composant interne" Then
    cmbItemSubcategory.RowSource = "internal_component"
ElseIf categoryChoiced = "connectique/câblage" Then
    cmbItemSubcategory.RowSource = "connector_cabling"
ElseIf categoryChoiced = "consommable" Then
    cmbItemSubcategory.RowSource = "consumable"
ElseIf categoryChoiced = "imprimante/scanner" Then
    cmbItemSubcategory.RowSource = "printer_scanner"
ElseIf categoryChoiced = "logiciel/licence" Then
    cmbItemSubcategory.RowSource = "software_licence"
ElseIf categoryChoiced = "matériel de bureau" Then
    cmbItemSubcategory.RowSource = "office_equipment"
ElseIf categoryChoiced = "matériel mobile" Then
    cmbItemSubcategory.RowSource = "mobile_hardware"
ElseIf categoryChoiced = "matériel réseau" Then
    cmbItemSubcategory.RowSource = "network_hardware"
ElseIf categoryChoiced = "périphérique" Then
    cmbItemSubcategory.RowSource = "peripheral"
ElseIf categoryChoiced = "stockage" Then
    cmbItemSubcategory.RowSource = "storage"
End If
End Sub

' ----------------------------------------------------------------------------------------------
' Objectif : Réinitilise le champ de sous-catégorie
' ----------------------------------------------------------------------------------------------
Private Sub cmbItemCategory_Click()
cmbItemSubcategory.Value = ""
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par nom" ? tri ascendant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemLabel_Click()
SortStockByLabelAscending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par nom" ? tri descendant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SortStockByLabelDescending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par quantité" ? croissant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemQuantity_Click()
SortStockByCurrentQuantityAscending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par quantité" ? décroissant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemQuantity_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SortStockByCurrentQuantityDescending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par catégorie" ? A ? Z
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemCategory_Click()
SortStockByCategoryAscending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par catégorie" ? Z ? A
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemCategory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 SortStockByCategoryDescending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par date" ? plus ancien au plus récent
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemUpdateDate_Click()
 SortStockByUpdateDateAscending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par date" ? plus récent au plus ancien
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemUpdateDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SortStockByUpdateDateDescending
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur "Quantités faibles" : affiche uniquement les matériels dont la quantité = quantité mini
' ----------------------------------------------------------------------------------------------
Private Sub btnFilterLowQuantity_Click()
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long

' Récupère l’adresse et le numéro de ligne max de la table "stock"
rangeStockAddress = rangeStock.Address
rangeStockAddressPart = Split(rangeStockAddress, "$")
rangeStockLastLine = CLng(rangeStockAddressPart(4))

' Vide la liste avant de la remplir
lstItems.Clear

' Vide la liste des historiques de mouvements du matériel
lstItemHistorical.Clear

' Parcours des matériels
For i = 3 To rangeStockLastLine
    ' Col2 = Qté actuelle / Col5 = Qté minimum ? filtre
    If tabStock.Range.Cells(i - 1, 2) <= tabStock.Range(i - 1, 5) Then
        lstItems.addItem tabStock.Range.Cells(i - 1, 1)
        lstItems.List(lstItems.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2)
        lstItems.List(lstItems.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3)
        lstItems.List(lstItems.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4)
    End If
Next i
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur "Quantités faibles" : réinitialise l’affichage
' ----------------------------------------------------------------------------------------------
Private Sub btnFilterLowQuantity_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur recherche" : filtre la liste selon le texte saisi dans txtSearchItem
' ----------------------------------------------------------------------------------------------
Private Sub btnSearchItem_Click()
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long
Dim userSearch As String
Dim itemLabelValue As String

' Détermine la plage "stock"
rangeStockAddress = rangeStock.Address
rangeStockAddressPart = Split(rangeStockAddress, "$")
rangeStockLastLine = CLng(rangeStockAddressPart(4))

' Récupère et prépare le texte recherché
userSearch = CStr(LCase(Trim(txtSearchItem.Value)))

' Vide la liste avant de la remplir avec les résultats
lstItems.Clear

' Vide la liste des historiques de mouvements du matériel
lstItemHistorical.Clear

' Vide la liste
' Boucle sur les matériels
For i = 3 To rangeStockLastLine
    itemLabelValue = CStr(LCase(Trim(tabStock.Range(i - 1, 1).Value)))
    ' InStr(..., ..., vbTextCompare) = 1 ? commence par le texte recherché
    If InStr(1, itemLabelValue, userSearch, vbTextCompare) = 1 Then
        lstItems.addItem tabStock.Range.Cells(i - 1, 1)
        lstItems.List(lstItems.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2)
        lstItems.List(lstItems.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3)
        lstItems.List(lstItems.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4)
    End If
Next i

    ' Réinitialise la zone de recherche
    txtSearchItem.Value = ""
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur "Recherche" : réaffiche tous les matériels
' ----------------------------------------------------------------------------------------------
Private Sub btnSearchItem_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Rafraîchit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur "Supprimer" : efface l’matériel sélectionné après confirmation
' ----------------------------------------------------------------------------------------------
Private Sub btnDeleteItem_Click()
Dim activeItemLabel As String
Dim rowToDelete As Variant
Dim lastRow As Long
Dim i As Long

On Error Resume Next
activeItemLabel = lstItems.Value
If activeItemLabel = "" Then Exit Sub

' Recalcul plage avant recherche
lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
Set rangeStock = wsStock.Range("A3:A" & lastRow)

On Error Resume Next
rowToDelete = Application.Match(activeItemLabel, rangeStock, 0)
On Error GoTo 0

If IsError(rowToDelete) Then
    MsgBox "Élément introuvable."
    Exit Sub
End If

' Demande de confirmation avant suppression
If MsgBox("Confirmer la suppression de : " & activeItemLabel, vbYesNo) = vbYes Then
    ' +2 car plage commence à A3
    wsStock.Rows(rowToDelete + 2).Delete
End If

' Recalcul plage après suppression
lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
Set rangeStock = wsStock.Range("A3:A" & lastRow)

' Recharge la liste après suppression
lstItems.Clear
For i = 3 To lastRow
    lstItems.addItem wsStock.Cells(i, 1).Value
    lstItems.List(lstItems.ListCount - 1, 1) = wsStock.Cells(i, 2).Value
    lstItems.List(lstItems.ListCount - 1, 2) = wsStock.Cells(i, 3).Value
    lstItems.List(lstItems.ListCount - 1, 3) = wsStock.Cells(i, 4).Value
Next i

' Sauvegarde le classeur
wb.Save

End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur "Sauvegarder" : enregistre les modifications apportées au matériel sélectionné
' ----------------------------------------------------------------------------------------------
Private Sub btnSaveItemUpdate_Click()
    Dim activeItemLabel As String
    Dim rowToUpdate As Variant
    Dim lastRow As Long
    Dim saveConfirmation As VbMsgBoxResult
    Dim i As Long
    Dim itemMinQuantity As Variant

    ' Vérifier la sélection de l'élément dans la ListBox
    If lstItems.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un matériel à modifier.", vbExclamation
        Exit Sub
    End If
    
    ' Conserver la valeur de l'élément sélectionné
    activeItemLabel = lstItems.Value
    
    '  Re-calculer la plage de recherche
    lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
    Dim rangeStock As Range
    Set rangeStock = wsStock.Range("A3:A" & lastRow)
    
    ' Chercher la ligne correspondante avec Application.Match
    On Error Resume Next
    rowToUpdate = Application.Match(activeItemLabel, rangeStock, 0)
    On Error GoTo 0
    
    ' Gérer l'erreur si matériel n'est pas trouvé
    If IsError(rowToUpdate) Then
        MsgBox "Erreur : matériel sélectionné n'a pas été trouvé dans le tableau.", vbCritical
        Exit Sub
    End If
    
    ' Demander confirmation avant la sauvegarde
    saveConfirmation = MsgBox("Confirmer la sauvegarde des modifications", vbYesNo)
    
    If saveConfirmation = vbYes Then
        ' Mettre à jour les données sur la feuille de calcul
        ' Le +2 est nécessaire car la plage commence à la ligne 3 (rowToUpdate est un index basé sur la plage)
        wsStock.Cells(rowToUpdate + 2, 1).Value = txtItemLabel.Value
        wsStock.Cells(rowToUpdate + 2, 3).Value = cmbItemCategory.Value
        wsStock.Cells(rowToUpdate + 2, 6).Value = cmbItemSubcategory.Value
        
        ' Récupération de la valeur du seuil
        itemMinQuantity = txtItemMinQuantity.Value
        
        ' Vérification que la valeur est bien numérique sinon elle sera égale à 0
        If Not IsNumeric(itemMinQuantity) Then
            itemMinQuantity = 0
        End If
        
        ' Vérification que la valeur est bien positive sinon elle sera convertie en nombre positif : -5 deviendra 5
        If itemMinQuantity < 0 Then
            itemMinQuantity = itemMinQuantity - (itemMinQuantity * 2)
        End If
            
        wsStock.Cells(rowToUpdate + 2, 5).Value = itemMinQuantity
        wsStock.Cells(rowToUpdate + 2, 7).Value = CStr(LCase(Trim(txtItemComment.Value)))
        
        MsgBox "Modifications sauvegardées avec succès !", vbInformation
    End If

    ' Recharger la ListBox
    lstItems.Clear
    For i = 3 To lastRow
        lstItems.addItem wsStock.Cells(i, 1).Value
        lstItems.List(lstItems.ListCount - 1, 1) = wsStock.Cells(i, 2).Value
        lstItems.List(lstItems.ListCount - 1, 2) = wsStock.Cells(i, 3).Value
        lstItems.List(lstItems.ListCount - 1, 3) = wsStock.Cells(i, 4).Value
    Next i
    
    ' Resélectionner l'élément dans la ListBox
    For i = 0 To lstItems.ListCount - 1
        If lstItems.List(i, 0) = activeItemLabel Then
            lstItems.Selected(i) = True
            Exit For
        End If
    Next i
End Sub

' ----------------------------------------------------------------------------------------------
' Objectif : Sauvegarde le classeur et propose l'affichage du dlasseur avant la fermeture de l'application
' ----------------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
' Sauvegarde le classeur
wb.Save
' MsgBox qui permet de demander à l'utilisateur s'il veut afficher ou non le classeur Excel après la fermeture de l'application
choice = MsgBox("Ouvrir le classeur Excel ?" & vbCrLf & "Oui : Fermer l'application et afficher le classeur Excel" & vbCrLf & "Non : Fermer l'application", vbYesNo)
' Si le choix est "oui", l'application se ferme et le classeur devient visible, si le choix est "non", l'application se ferme et le classeur reste invisible
If choice = vbYes Then
    Application.Visible = True
End If
End Sub
