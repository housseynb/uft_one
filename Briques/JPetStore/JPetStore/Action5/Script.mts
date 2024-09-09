' Initialisation du navigateur et de la page
Set browserObj = Browser("JPetStore Demo")
Set home = browserObj.Page("Home")
Set Account = browserObj.Page("Account")

' Variables pour les saisies (langue, favoris, état des checkboxes)
Dim langue, animal
langue = "japanese"
animal = "CATS"

' Actions à exécuter
home.Link("My Account").Click
Account.WebList("Language").Select langue
Account.WebList("Favoris").Select animal
'Account.WebCheckBox("name:=account.listOption").Set False
'Account.WebCheckBox("name:=account.bannerOption").Set False
Account.WebCheckBox("MyList").Set "Off"
Account.WebCheckBox("MyBanner").Set "Off"
Account.WebButton("Save").Click
home.Link("My Account").Click
Wait 3

' Vérification des éléments après sauvegarde

' Vérification de la langue sélectionnée
Call VerifierValeur(Account.WebList("Language"), langue)

' Vérification de l'animal favori sélectionné 
Call VerifierValeur(Account.WebList("Favoris"), animal)

' Vérification de l'état coché ou non coché d'une case à cocher
Call VerifierCheckBox(Account.WebCheckBox("MyList"), False)
Call VerifierCheckBox(Account.WebCheckBox("MyBanner"), False)
