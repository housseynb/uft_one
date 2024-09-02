' Déclaration des variables globales pour le navigateur, la page et les éléments à vérifier
Dim browserObj, pageObj
Dim imgDog2, linkFriendlyDog, elementBulldog
Dim resultDog2, resultLink, resultBulldog

' Initialisation des objets navigateur et page
Set browserObj = Browser("JPetStore Demo") ' Référence à l'objet navigateur
Set pageObj = browserObj.Page("JPetStore Demo") ' Référence à l'objet page

' Initialisation des éléments à vérifier sur la page
Set imgDog2 = pageObj.Image("dog2") ' Référence à l'image du chien (dog2)
Set linkFriendlyDog = pageObj.Link("Friendly dog from England") ' Référence au lien du chien amical d'Angleterre
Set elementBulldog = pageObj.WebElement("Bulldog") ' Référence à l'élément web représentant le Bulldog

' Appliquer la valeur dans le champ de recherche depuis le paramètre VALEUR_RECHERCHE
pageObj.WebEdit("keyword").Set Parameter("VALEUR_RECHERCHE") ' Remplir le champ de recherche avec la valeur paramétrée

' Cliquer sur le bouton de recherche pour lancer la recherche
pageObj.WebButton("Search").Click ' Cliquer sur le bouton de recherche

' Synchroniser l'exécution pour attendre que l'élément Bulldog soit visible
Browser("JPetStore Demo").Page("JPetStore Demo").WebElement("Bulldog").WaitProperty "visible", True, 5000 ' Attendre jusqu'à 5 secondes que l'élément Bulldog soit visible

' Ajouter un Output Checkpoint pour capturer la valeur de l'élément K9-BD-01
Browser("JPetStore Demo").Page("JPetStore Demo").WebElement("K9-BD-01").Output CheckPoint("K9-BD-01") ' Ajouter un point de contrôle de sortie pour l'élément K9-BD-01

' Vérification des éléments sur la page en utilisant la fonction VerifElement
resultDog2 = VerifElement(imgDog2, "dog2", "Passant") ' Vérifier la présence de l'image dog2
resultLink = VerifElement(linkFriendlyDog, "Friendly dog from England", "Passant") ' Vérifier la présence du lien Friendly dog from England
resultBulldog = VerifElement(elementBulldog, "Bulldog", "Passant") ' Vérifier la présence de l'élément Bulldog

' Déterminer le statut global du test basé sur les résultats des vérifications
If resultDog2 = "OK" And resultLink = "OK" And resultBulldog = "OK" Then
    Parameter("TEST_STATUS") = "OK"  ' Tous les éléments sont trouvés, donc le test est OK
Else
    Parameter("TEST_STATUS") = "KO"  ' Un ou plusieurs éléments ne sont pas trouvés, donc le test est KO
End If

' Libération des objets pour éviter les fuites de mémoire
Set imgDog2 = Nothing
Set linkFriendlyDog = Nothing
Set elementBulldog = Nothing
Set browserObj = Nothing
Set pageObj = Nothing

' Fonction VerifElement - Vérifie l'existence d'un élément et retourne un statut
Function VerifElement(Element, nomElement, Etat)
    ' Vérifier si l'élément existe avec un délai d'attente de 2 secondes
    existence = Element.Exist(2)
    
    ' Vérification en fonction de l'état attendu (Passant ou NonPassant)
    Select Case Etat
        Case "Passant"
            ' Cas où l'élément est attendu et doit être trouvé
            Select Case existence
                Case True
                    ' L'élément est trouvé comme attendu, rapporter un succès
                    Reporter.ReportEvent micPass, "Verif Elément", "L'élément " & nomElement & " a été trouvé"
                    VerifElement = "OK"
                Case False
                    ' L'élément n'est pas trouvé mais il était attendu, rapporter un échec
                    Reporter.ReportEvent micFail, "Verif Elément", "L'élément " & nomElement & " n'a pas été trouvé: Le test est en échec: KO"
                    VerifElement = "KO"
            End Select
        Case "NonPassant"
            ' Cas où l'élément n'est pas attendu et ne doit pas être trouvé
            Select Case existence
                Case False
                    ' L'élément n'est pas trouvé comme attendu, rapporter un succès
                    Reporter.ReportEvent micPass, "Verif Elément", "L'élément " & nomElement & " n'a pas été trouvé comme attendu"
                    VerifElement = "OK"
                Case True
                    ' L'élément est trouvé mais il ne devait pas l'être, rapporter un échec
                    Reporter.ReportEvent micFail, "Verif Elément", "L'élément " & nomElement & " a été trouvé contrairement à l'attendu"
                    VerifElement = "KO"
            End Select
    End Select
End Function

