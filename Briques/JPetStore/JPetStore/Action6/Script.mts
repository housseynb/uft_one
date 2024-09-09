' Initialisation du navigateur et de la page
Set browserObj = Browser("JPetStore Demo")
Set home = browserObj.Page("Home")
Set Buy = browserObj.Page("Buy")

' Déclaration des variables
Dim prixInitial, prixTotalMiseAJour, prixTotalConfirme
Dim quantite, montantAttendu
Dim valeur

' Actions à exécuter
home.Image("fish").Click
home.Link("Angelfish").Click
home.Link("AddCart").Click

' Récupérer le prix initial
SynchroniserVisibilite Buy.WebElement("Total Cost"), 5
valeur = Buy.WebElement("Total Cost").GetROProperty("InnerText")
prixInitial = ExtraireValeur(valeur) ' Convertir la valeur récupérée

' Définir la quantité souhaitée
quantite = 3 ' Modifier cette valeur selon le besoin

' Mettre à jour la quantité et le panier
Buy.WebEdit("Quantity").Set quantite
Buy.WebButton("UpdateCart").Click

' Récupérer le prix après mise à jour
SynchroniserVisibilite Buy.WebElement("Total Cost"), 5
valeur = Buy.WebElement("Total Cost").GetROProperty("InnerText")
prixTotalMiseAJour = ExtraireValeur(valeur) ' Convertir la valeur récupérée

' Calculer le montant attendu
montantAttendu = prixInitial * quantite

' Vérifier que le prix après mise à jour est correct
If prixTotalMiseAJour = montantAttendu Then
    Reporter.ReportEvent micPass, "Verification du Prix après Mise à Jour", "Le prix après mise à jour est correct : $" & prixTotalMiseAJour
Else
    Reporter.ReportEvent micFail, "Verification du Prix après Mise à Jour", "Le prix après mise à jour est incorrect : $" & prixTotalMiseAJour
End If

' Continuer avec la commande
Buy.Link("Proceed to Checkout").Click
Buy.WebButton("Continue").Click
Buy.Link("Confirm").Click

' Récupérer le prix confirmé
SynchroniserVisibilite Buy.WebElement("ConfirmTotal"), 5
valeur = Buy.WebElement("ConfirmTotal").GetROProperty("InnerText")
prixTotalConfirme = ExtraireValeur(valeur) ' Convertir la valeur récupérée

' Vérifier que le prix confirmé est correct
If prixTotalConfirme = montantAttendu Then
    Reporter.ReportEvent micPass, "Verification du Prix Confirmé", "Le prix confirmé est correct : $" & prixTotalConfirme
Else
    Reporter.ReportEvent micFail, "Verification du Prix Confirmé", "Le prix confirmé est incorrect : $" & prixTotalConfirme
End If

' Nettoyage
Set browserObj = Nothing
Set home = Nothing
Set Buy = Nothing

