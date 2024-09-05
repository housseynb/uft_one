' Déclaration des variables
Dim browserObj, pageObj
Dim elementAnimal, resultAnimal
Dim XLHandle, XLBook, InSheet, numLigne, valeurRecherche
Dim maxLignes, compteur

' Initialisation d'Excel
Set XLHandle = CreateObject("Excel.Application")
XLHandle.DisplayAlerts = False
Set XLBook = XLHandle.WorkBooks.Open("C:\Users\hboudjenane\Desktop\git_one\Ressources\Animals.xlsx")
Set InSheet = XLBook.Worksheets("Animals")

' Initialisation du navigateur et de la page
Set browserObj = Browser("JPetStore Demo")
Set pageObj = browserObj.Page("JPetStore Demo")

' Définir la limite de lignes
maxLignes = 100
compteur = 0

' Boucle sur les lignes de la colonne A
numLigne = 1
Do While InSheet.Cells(numLigne, 1).Value <> "" And compteur < maxLignes

    ' Récupérer la valeur de la colonne A
    valeurRecherche = InSheet.Cells(numLigne, 1).Value
    Reporter.ReportEvent micInfo, "Recherche", "Recherche pour : " & valeurRecherche

    ' Rechercher l'animal
    pageObj.WebEdit("keyword").Set valeurRecherche
    pageObj.WebButton("Search").Click 

    ' Modifier et attendre l'élément
    pageObj.WebElement("Animals").SetTOProperty "innertext", valeurRecherche
    
    ' Utiliser la procédure SynchroniserVisibilite
    SynchroniserVisibilite pageObj.WebElement("Animals"), 5

    ' Vérification de l'élément avec la fonction VerifElement
    Set elementAnimal = pageObj.WebElement("Animals")
    resultAnimal = VerifElement(elementAnimal, valeurRecherche, "Passant")

    ' Vérifier l'innertext de l'élément
    innerTextResult = VerifierInnerText(elementAnimal, valeurRecherche)

    ' Rapport des résultats
    If resultAnimal = "OK" And innerTextResult = "OK" Then
        Reporter.ReportEvent micPass, "Test Status", "Le test pour " & valeurRecherche & " est OK"
    Else
        Reporter.ReportEvent micFail, "Test Status", "Le test pour " & valeurRecherche & " est KO"
    End If

    ' Incrémenter la ligne et le compteur
    numLigne = numLigne + 1
    compteur = compteur + 1

Loop

' Libération des objets Excel
XLBook.Close False
XLHandle.Quit
Set InSheet = Nothing
Set XLBook = Nothing
Set XLHandle = Nothing

' Libération des objets UFT
Set elementAnimal = Nothing
Set browserObj = Nothing
Set pageObj = Nothing
