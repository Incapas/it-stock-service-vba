VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMovement 
   Caption         =   "UserForm1"
   ClientHeight    =   7860
   ClientLeft      =   30
   ClientTop       =   135
   ClientWidth     =   10215
   OleObjectBlob   =   "frmMovement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================================
' Procédure : UserForm_Initialize
' Objectif  : Initialiser le formulaire "Ajouter un mouvement" en définissant les références
'             aux données et en configurant l'interface graphique (front-end)
' ==============================================================================================

Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------------------------
' Section pour déclaration des variables et initialisation des références
' ----------------------------------------------------------------------------------------------

' Référence au classeur qui contient la macro
Set wb = ThisWorkbook

' Référence à la feuille "stock" contenant la liste des matériels
Set wsStock = wb.Worksheets("stock")

' Référence au tableau structuré nommé "stock"
Set tabStock = wsStock.ListObjects("stock")

' Référence à la feuille "mouvement" contenant l’historique des mouvements
Set wsMovement = wb.Worksheets("mouvement")

' Référence au tableau structuré nommé "movement" (historique des entrées/sorties)
Set tabMovement = wsMovement.ListObjects("movement")

' Référence à la plage physique correspondant au tableau "movement"
Set rangeMovement = wsMovement.Range("movement")

' Variables pour découper les adresses de plages et déterminer le nombre de lignes
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long
Dim rangeMovementAddressPart() As String
Dim rangeMovementLastLine As Long
Dim startDate As Date
Dim endDate As Date

' ----------------------------------------------------------------------------------------------
' Section pour définir le front-end du formulaire (dimensions, titre, couleurs)
' ----------------------------------------------------------------------------------------------

' Propriétés générales du formulaire
With Me
    .Width = 420
    .Height = 300
    .Caption = "Ajouter un mouvement"
    .BackColor = COLOR_GRAY_DARK
End With

' Label "Matériel"
With lblMovementItem
    .Left = 30
    .Top = 20
    .Width = 100
    .Height = 20
    .Caption = "Matériel"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Liste déroulante pour choisir le matériel
With cmbMovementItem
    .Left = 140
    .Top = 20
    .Width = 250
    .Height = 20
    .Style = fmStyleDropDownList
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .RowSource = tabStock
End With

' Label "Date"
With lblMovementDate
    .Left = 30
    .Top = 60
    .Width = 100
    .Height = 20
    .Caption = "Date"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Liste déroulante pour la date
With cmbMovementDate
    .Left = 140
    .Top = 60
    .Width = 250
    .Height = 20
    .MaxLength = 10
    .Style = fmStyleDropDownList
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Alimentation de la liste déroulante pour la date à partir d'une période donnée
startDate = CDate("01/01/2025")
endDate = CDate("31/12/2025")

While endDate <> startDate
    cmbMovementDate.addItem (endDate)
    endDate = endDate - 1
Wend

' Label "Type"
With lblMovementType
    .Left = 30
    .Top = 100
    .Width = 100
    .Height = 20
    .Caption = "Type"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton radio "Entrée"
With rdbEntry
    .Left = 140
    .Top = 100
    .Width = 50
    .Height = 20
    .Caption = "Entrée"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton radio "Sortie"
With rdbExit
    .Left = 210
    .Top = 100
    .Width = 50
    .Height = 20
    .Caption = "Sortie"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Label "Valeur"
With lblMovementValue
    .Left = 30
    .Top = 140
    .Width = 115
    .Height = 20
    .Caption = "Valeur"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Zone de saisie pour la valeur
With txtMovementValue
    .Left = 140
    .Top = 140
    .Width = 90
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Label "Description"
With lblMovementDescription
    .Left = 30
    .Top = 180
    .Width = 100
    .Height = 20
    .Caption = "Description"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Zone de saisie pour la description
With txtMovementDescription
    .Left = 140
    .Top = 180
    .Width = 250
    .Height = 20
    .MaxLength = 30
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton "Ajouter" ? Valide et enregistre le mouvement
With btnAddMovement
    .Left = 100
    .Top = 225
    .Width = 100
    .Height = 25
    .Caption = "Ajouter"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_FOREST_GREEN
    .ForeColor = COLOR_WHITE
End With

' Bouton "Annuler" ? Ferme le formulaire sans action
With btnCancelAddMovement
    .Left = 220
    .Top = 225
    .Width = 100
    .Height = 25
    .Caption = "Annuler"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_CRIMSON_DARK
    .ForeColor = COLOR_WHITE
End With
End Sub

' ----------------------------------------------------------------------------------------------
' Bouton "Ajouter" : enregistre un nouveau mouvement et met à jour la liste
' ----------------------------------------------------------------------------------------------
Private Sub btnAddMovement_Click()
 Dim wb As Workbook
 Dim wsStock As Worksheet, wsMovement As Worksheet
 Dim tabStock As ListObject, tabMovement As ListObject
 Dim rangeStock As Range
 Dim moveItemLabel As String, moveType As String, moveDescription As String
 Dim moveDate As Date
 Dim moveValue As Variant
 Dim activeItemRowTab As Variant
 Dim activeItemCurrentQuantity As Variant
 Dim i As Long

' Références de base
' Classeur actif
Set wb = ThisWorkbook
' Feuille stock
Set wsStock = wb.Worksheets("stock")
' Feuille mouvements
Set wsMovement = wb.Worksheets("mouvement")
' Tableau structuré des stocks
Set tabStock = wsStock.ListObjects("stock")
' Tableau structuré mouvements
Set tabMovement = wsMovement.ListObjects("movement")

' Lecture des données saisies par l'utilisateur
' Nom du matériel
 moveItemLabel = Trim(cmbMovementItem.Value)
  
' Si vide ? on sort
If Len(moveItemLabel) = 0 Then Exit Sub

' Conversion en date
moveDate = CDate(cmbMovementDate.Value)

' Détermination du type de mouvement
If rdbEntry.Value Then
    moveType = "entrée"
ElseIf rdbExit.Value Then
    moveType = "sortie"
Else
    Exit Sub
End If

' Récupération de la valeur du mouvement
moveValue = txtMovementValue.Value

' Vérification que la valeur est bien numérique sinon elle sera égale à 0
If Not IsNumeric(moveValue) Then
    moveValue = 0
End If

' Vérification que la valeur est bien positive sinon elle sera convertie en nombre positif : -5 deviendra 5
If moveValue < 0 Then
    moveValue = moveValue - (moveValue * 2)
End If
    
' Description libre
moveDescription = CStr(LCase(Trim(txtMovementDescription.Value)))
 
' Localisation du matériel dans le stock
Set rangeStock = tabStock.ListColumns(1).DataBodyRange
activeItemRowTab = Application.Match(moveItemLabel, rangeStock, 0)
If IsError(activeItemRowTab) Then Exit Sub    ' Si introuvable ? sortie

' Quantité actuelle avant mouvement
activeItemCurrentQuantity = tabStock.DataBodyRange.Cells(activeItemRowTab, 2).Value

' Mise à jour de la quantité en stock
If moveType = "entrée" Then
    tabStock.DataBodyRange.Cells(activeItemRowTab, 2).Value = CLng(activeItemCurrentQuantity) + moveValue
Else
    tabStock.DataBodyRange.Cells(activeItemRowTab, 2).Value = CLng(activeItemCurrentQuantity) - moveValue
End If

' Mise à jour de la date de dernière MAJ
tabStock.DataBodyRange.Cells(activeItemRowTab, 4).Value = moveDate

' Ajout de l'enregistement dans le tableau "movement"
addMovement moveDate, moveType, moveValue, moveDescription, moveItemLabel

' Rafraîchissement de la liste lstITems (vue d'ensemble des stocks)
frmStock.lstItems.Clear
For i = 1 To tabStock.DataBodyRange.Rows.Count
    frmStock.lstItems.addItem tabStock.DataBodyRange.Cells(i, 1).Value
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 1) = tabStock.DataBodyRange.Cells(i, 2).Value
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 2) = tabStock.DataBodyRange.Cells(i, 3).Value
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 3) = tabStock.DataBodyRange.Cells(i, 4).Value
Next i

' Rafraîchissement de la liste lstItemHistorical (vue individualisée des mouvements)
With frmStock.lstItemHistorical
    ' Supprime le lien direct au tableau
    .RowSource = ""
    ' Nombre de colonnes dans la liste
    .ColumnCount = 4
    .Clear
End With

 If Not tabMovement.DataBodyRange Is Nothing Then
        Dim idx As Long, j As Long
        For j = 1 To tabMovement.DataBodyRange.Rows.Count
            ' Filtre sur colonne 5 ? Matériel
            If CStr(tabMovement.DataBodyRange.Cells(j, 5).Value) = CStr(moveItemLabel) Then
                ' Ajout d’une ligne dans la listbox historique
                frmStock.lstItemHistorical.addItem tabMovement.DataBodyRange.Cells(j, 1).Value
                idx = frmStock.lstItemHistorical.ListCount - 1
                ' Colonnes suivantes : Type, Valeur, Description
                frmStock.lstItemHistorical.List(idx, 1) = tabMovement.DataBodyRange.Cells(j, 2).Value
                frmStock.lstItemHistorical.List(idx, 2) = tabMovement.DataBodyRange.Cells(j, 3).Value
                frmStock.lstItemHistorical.List(idx, 3) = tabMovement.DataBodyRange.Cells(j, 4).Value
            End If
        Next j
    End If
    
' Fermeture du formulaire
Unload Me
End Sub

' ----------------------------------------------------------------------------------------------
' Bouton "Annuler" : ferme simplement le formulaire sans rien modifier
' ----------------------------------------------------------------------------------------------
Private Sub btnCancelAddMovement_Click()
    Unload Me
End Sub
