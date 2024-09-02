'Variables
Dim nav, pageJPet

'Set des objets
Set nav = Browser("JPetStore Demo")
Set pageJPet = Page("JPetStore Demo")
    
    ' Appliquer la valeur dans le champ de recherche et cliquer sur le bouton
    nav.pageJPet.WebEdit("keyword").Set Parameter("VALEUR_RECHERCHE")
    nav.pageJPet.WebButton("Search").Click
