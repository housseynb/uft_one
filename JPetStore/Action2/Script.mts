' Récupérer le nombre de lignes dans la DataTable globale
rowCount = DataTable.GetRowCount

' Boucle à travers chaque ligne de la DataTable
For i = 1 To rowCount
    ' Sélectionner la ligne actuelle
    DataTable.SetCurrentRow(i)
    
    ' Récupérer la valeur de la colonne "Animals"
    animalType = DataTable.Value("Animals", dtGlobalSheet)
    
    ' Appliquer la valeur dans le champ de recherche et cliquer sur le bouton
    Browser("JPetStore Demo").Page("JPetStore Demo").WebEdit("keyword").Set animalType
    Browser("JPetStore Demo").Page("JPetStore Demo").WebButton("Search").Click
Next
