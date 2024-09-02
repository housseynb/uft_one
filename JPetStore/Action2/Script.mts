' Récupérer la valeur de la colonne "Animals"
animalType = DataTable.Value("Animals", dtGlobalSheet)
    
' Appliquer la valeur dans le champ de recherche et cliquer sur le bouton
Browser("JPetStore Demo").Page("JPetStore Demo").WebEdit("keyword").Set animalType
Browser("JPetStore Demo").Page("JPetStore Demo").WebButton("Search").Click
