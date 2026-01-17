Attribute VB_Name = "modStockSortingSub"
' ==============================================================================================
' Procédure : SortStockByLabelAscending
' Objectif  : Trier le stock par libellé (A ? Z)
' ==============================================================================================
Public Sub SortStockByLabelAscending()
    ' On efface les critères de tri précédents
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    
    ' Ajout du critère : colonne [libellé] en ordre croissant
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Application du tri
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        ' La première ligne contient les en-têtes
        .Header = xlYes
        ' Ne pas tenir compte de la casse
        .MatchCase = False
        ' Tri vertical
        .Orientation = xlTopToBottom
        ' Méthode de tri compatible avec caractères accentués
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Procédure : SortStockByLabelDescending
' Objectif  : Trier le stock par libellé (Z ? A)
' ==============================================================================================
Public Sub SortStockByLabelDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Procédure : SortStockByCurrentQuantityAscending
' Objectif  : Trier par quantité (croissante), puis libellé (A ? Z)
' ==============================================================================================
Public Sub SortStockByCurrentQuantityAscending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier critère : quantité
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[stock]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' Deuxième critère : libellé
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Procédure : SortStockByCurrentQuantityDescending
' Objectif  : Trier par quantité (décroissante), puis libellé (A ? Z)
' ==============================================================================================
Public Sub SortStockByCurrentQuantityDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier critère : quantité décroissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[stock]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ' Deuxième critère : libellé
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Procédure : SortStockByCategoryAscending
' Objectif  : Trier par catégorie (A ? Z), puis libellé (A ? Z)
' ==============================================================================================
Public Sub SortStockByCategoryAscending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier critère : catégorie
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[catégorie]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' Deuxième critère : libellé
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Procédure : SortStockByCategoryDescending
' Objectif  : Trier par catégorie (Z ? A), puis libellé (A ? Z)
' ==============================================================================================
Public Sub SortStockByCategoryDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier critère : catégorie décroissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[catégorie]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ' Deuxième critère : libellé
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Procédure : SortStockByUpdateDateAscending
' Objectif  : Trier par date de mise à jour (ancienne ? récente), puis libellé (A ? Z)
' ==============================================================================================
Public Sub SortStockByUpdateDateAscending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier critère : date croissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[maj]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' Deuxième critère : libellé
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Procédure : SortStockByUpdateDateDescending
' Objectif  : Trier par date de mise à jour (récent ? ancien), puis libellé (A ? Z)
' ==============================================================================================
Public Sub SortStockByUpdateDateDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier critère : date décroissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[maj]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ' Deuxième critère : libellé
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libellé]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
