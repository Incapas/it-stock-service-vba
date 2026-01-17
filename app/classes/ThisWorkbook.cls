VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ----------------------------------------------------------------------------------------------
' Objectif : Trie la table "movement" par date décroissante et affiche le formulaire frmStock
' ----------------------------------------------------------------------------------------------
Private Sub Workbook_Open()
    ' Efface les critères de tri existants sur la table "movement" de la feuille "mouvement"
    ActiveWorkbook.Worksheets("mouvement").ListObjects("movement").Sort.SortFields.Clear

    ' Ajoute un nouveau critère de tri sur la colonne "date" en ordre décroissant
    ActiveWorkbook.Worksheets("mouvement").ListObjects("movement").Sort.SortFields. _
        Add2 Key:=Range("movement[[#Headers],[date]]"), SortOn:=xlSortOnValues, _
        Order:=xlDescending, DataOption:=xlSortNormal

    ' Applique les paramètres de tri à la table
    With ActiveWorkbook.Worksheets("mouvement").ListObjects("movement").Sort
        ' Indique que la première ligne contient les en-têtes
        .Header = xlYes
        ' Ignore la casse lors du tri
        .MatchCase = False
        ' Trie de haut en bas
        .Orientation = xlTopToBottom
        ' Méthode de tri utilisée (utile pour les caractères chinois)
        .SortMethod = xlPinYin
        ' Applique le tri
        .Apply
    End With

    ' Masque l'application Excel pour ne laisser apparaître que le formulaire
    Application.Visible = False

    ' Affiche le formulaire principal
    frmStock.Show
End Sub
